// One M365 Copilot chat turn over SignalR/WebSocket.
//
// Protocol notes that matter (full write-up: the m365-copilot-proxy docs):
//  - The access token rides in the WebSocket URL query string, not a header.
//  - Node's native WebSocket does NOT work — use `ws` with a browser Origin/UA.
//  - Every frame is terminated by a 0x1E (record separator) byte.
//  - A chat turn is the chat invocation PLUS a `Metrics` frame in ONE send, or
//    the turn silently produces nothing.
//  - The model is chosen by a `tone` string, not a model id.
//  - Streaming mixes token deltas with full-text snapshots, so we fold both.

import WebSocket from "ws";
import { createLogger } from "./log.js";
import { decodeJwt } from "./auth.js";

const RS = "\x1E";
const log = createLogger("client");

// Real "Stop generating" frame, captured from the web client. Sent on abort so
// the server discards the partial answer instead of running to completion.
const STOP_FRAME = JSON.stringify({ arguments: [{}], invocationId: "1", target: "stop", type: 1 }) + RS;

// Model id → `tone`. The server VALIDATES tones (unknown → "Failed to invoke
// 'Chat'"), so only add tones confirmed live.
const MODEL_TONES = {
  "m365-copilot": "magic",
  "auto": "magic",
  // NB: "quick" -> "Gpt_Quick" was rejected server-side ("Failed to invoke
  // 'Chat'") when tested, so it is intentionally not offered. Tones are
  // server-validated and Microsoft retires them freely — reprobe before adding.
  "think-deeper": "Gpt_Reasoning",
  "claude": "Claude_Sonnet",
  "claude-sonnet": "Claude_Sonnet",
  "gpt-5.5": "Gpt_5_5_Chat",
  "gpt-5.5-quick": "Gpt_5_5_Chat",
  "gpt-5.5-think-deeper": "Gpt_5_5_Reasoning",
  "gpt-5.6-think-deeper": "Gpt_5_6_Reasoning",
  "gpt-5.4": "Gpt_5_4_Reasoning",
  "gpt-5.4-quick": "Gpt_5_4_Quick",
  "gpt-5.3-quick": "Gpt_5_3_Quick",
  "gpt-5.2-quick": "Gpt_5_2_Quick",
};

export function getToneForModel(model) {
  if (MODEL_TONES[model]) return MODEL_TONES[model];
  if (/^claude/i.test(model)) return "Claude_Sonnet";
  return "magic";
}

export function getAvailableModels() {
  return Object.keys(MODEL_TONES);
}

// Feature flags lifted from a captured web-client session. Removing them is
// untested, so we send the proven list.
const VARIANTS = [
  "EnableMcpServerWidgets",
  "feature.EnableMcpServerWidgets",
  "feature.EnableLuForChatCIQ",
  "feature.enableChatCIQPlugin",
  "EnableRequestPlugins",
  "feature.EnableSensitivityLabels",
  "EnableUnsupportedUrlDetector",
  "feature.IsCustomEngineCopilotEnabled",
  "feature.bizchatfluxv3",
  "feature.enablechatpages",
  "feature.enableCodeCanvas",
  "feature.turnOnWorkTabRecommendation",
  "turnOffWorkTabUpsellFromClient",
  "feature.turnOnDARecommendation",
  "feature.IsStreamingModeInChatRequestEnabled",
  "IncludeSourceAttributionsConcise",
  "SkipPublishEmptyMessage",
  "feature.EnableDeduplicatingSourceAttributions",
  "Enable3PActionProgressMessages",
  "feature.enableClientWebRtc",
  "feature.EnableMeetingRecapOfSeriesMeetingWithCiq",
  "feature.EnableReferencesListCompleteSignal",
  "feature.StorageMessageSplitDisabled",
  "feature.EnableCuaTakeControlApi",
  "feature.cwcallowedos",
  "feature.disabledisallowedmsgs",
  "feature.enableCitationsForSynthesisData",
  "feature.enableGenerateGraphicArtOptionsSet",
  "cdximagen",
  "feature.EnableUpdatedUXForConfirmationDialog",
  "feature.EnableClientFileURLSupportForOfficeWebPaidCopilot",
  "feature.EnableDesignEditorImageGrounding",
  "feature.EnableDesignerEditor",
  "feature.OfficeWebToHelix",
  "feature.OfficeDesktopToHelix",
  "feature.M365TeamsHubToHelix",
  "feature.OwaHubToHelix",
  "feature.MonarchHubToHelix",
  "feature.Win32OutlookHubToHelix",
  "feature.MacOutlookHubToHelix",
  "Agt_bizchat_enableGpt5ForHelix",
].join(",");

// Unlocks M365's real server-side Python sandbox (the model executes Python and
// returns true results). Disable with M365_NO_CODE_INTERPRETER=1.
const CODE_INTERPRETER_OPTIONS_SETS = [
  "cwc_code_interpreter",
  "cwc_code_interpreter_amsfix",
  "cwc_code_interpreter_citation_fix",
  "code_interpreter_interactive_charts",
  "code_interpreter_matplotlib_patching",
];

/**
 * Fold streamed text into the running answer, returning the new answer and the
 * suffix to emit. M365 mixes token deltas with full-text snapshots, and the
 * first token often arrives ONLY as a snapshot — so naive delta concatenation
 * drops the head. We only ever emit a true prefix (already-emitted bytes can't
 * be retracted).
 */
export function foldStreamText(answer, next) {
  if (next.length <= answer.length) return { answer, emit: null };
  if (next.startsWith(answer)) return { answer: next, emit: next.slice(answer.length) };
  return { answer: next, emit: null };
}

/**
 * A persistent M365 Copilot conversation. One WebSocket per turn, but the same
 * sessionId/conversationId are reused so M365 threads the server-side context.
 */
export class CopilotSession {
  constructor({ sessionId, conversationId } = {}) {
    this.sessionId = sessionId ?? crypto.randomUUID();
    this.conversationId = conversationId ?? crypto.randomUUID();
    this.turnCount = 0;
    log.info(`New session: sid=${this.sessionId}, cid=${this.conversationId}`);
  }

  /**
   * Send one message and stream the response.
   * @returns {Promise<CopilotStream>} async-iterable of delta strings; also has
   *   getters fullText/hasContent/messageType/contentOrigin/throttle/scores.
   */
  chat(token, text, model = "m365-copilot", signal) {
    const isFirst = this.turnCount === 0;
    this.turnCount++;
    log.info(`Chat turn ${this.turnCount - 1}: model=${model}, first=${isFirst}`);

    const claims = decodeJwt(token);
    const requestId = crypto.randomUUID();
    const sessionId = this.sessionId;

    const params = new URLSearchParams({
      chatsessionid: requestId,
      clientrequestid: requestId,
      "X-SessionId": sessionId,
      ConversationId: this.conversationId,
      access_token: token,
      variants: VARIANTS,
      source: '"officeweb"',
      product: "Office",
      agentHost: "Bizchat.FullScreen",
      licenseType: "Starter",
      agent: "web",
      scenario: "OfficeWebIncludedCopilot",
    });
    const wsUrl = `wss://substrate.office.com/m365Copilot/Chathub/${claims.oid}@${claims.tid}?${params}`;

    return new Promise((resolve, reject) => {
      let answer = "";
      let hasContent = false;
      let throttle = null;
      let contentOrigin = null;
      let messageType = null;
      let turnCountServer = null;
      const maxScores = {};

      // Delta pump, at Promise scope so tokens arriving before the consumer
      // starts iterating are buffered instead of dropped.
      const queue = [];
      let done = false;
      let streamError = null;
      let waiting = null;
      const onDelta = (t) => {
        if (waiting) {
          const w = waiting;
          waiting = null;
          w.resolve({ value: t, done: false });
        } else queue.push(t);
      };
      const onDone = () => {
        done = true;
        if (waiting) {
          const w = waiting;
          waiting = null;
          w.resolve({ value: undefined, done: true });
        }
      };
      const onError = (err) => {
        streamError = err;
        done = true;
        if (waiting) {
          const w = waiting;
          waiting = null;
          w.reject(err);
        }
      };
      const advance = (next) => {
        const r = foldStreamText(answer, next);
        if (r.answer !== answer) {
          answer = r.answer;
          hasContent = true;
        }
        if (r.emit) onDelta(r.emit);
      };

      const stream = {
        get fullText() { return answer; },
        get hasContent() { return hasContent; },
        get throttle() { return throttle; },
        get contentOrigin() { return contentOrigin; },
        get messageType() { return messageType; },
        get turnCount() { return turnCountServer; },
        get scores() { return Object.keys(maxScores).length ? { ...maxScores } : null; },
        [Symbol.asyncIterator]() {
          return {
            next() {
              if (queue.length > 0) return Promise.resolve({ value: queue.shift(), done: false });
              if (streamError) return Promise.reject(streamError);
              if (done) return Promise.resolve({ value: undefined, done: true });
              return new Promise((res, rej) => { waiting = { resolve: res, reject: rej }; });
            },
          };
        },
      };

      const wsOptions = {
        headers: {
          "Origin": "https://m365.cloud.microsoft",
          "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:148.0) Gecko/20100101 Firefox/148.0",
          "Accept-Language": "en-US,en;q=0.9",
          "Cache-Control": "no-cache",
          "Pragma": "no-cache",
        },
      };
      // Corporate TLS-inspecting proxy (Zscaler/Netskope/…) → "self-signed
      // certificate in certificate chain". Prefer NODE_EXTRA_CA_CERTS with the
      // corp CA; M365_INSECURE=1 is the blunt fallback.
      if (process.env.M365_INSECURE === "1") wsOptions.rejectUnauthorized = false;

      const ws = new WebSocket(wsUrl, wsOptions);
      let handshakeDone = false;
      let stopped = false;

      const onAbort = () => {
        if (stopped) return;
        stopped = true;
        try {
          if (ws.readyState === WebSocket.OPEN && handshakeDone) {
            ws.send(STOP_FRAME);
            setTimeout(() => { try { ws.close(); } catch {} }, 2_000);
          } else {
            ws.close();
          }
        } catch {
          try { ws.close(); } catch {}
        }
      };
      if (signal) {
        if (signal.aborted) onAbort();
        else signal.addEventListener("abort", onAbort, { once: true });
      }
      const clearAbort = () => signal?.removeEventListener("abort", onAbort);

      ws.on("open", () => {
        ws.send(JSON.stringify({ protocol: "json", version: 1 }) + RS);
      });

      ws.on("message", (data) => {
        for (const frame of data.toString().split(RS).filter(Boolean)) {
          let parsed;
          try {
            parsed = JSON.parse(frame);
          } catch {
            if (!handshakeDone) { handshakeDone = true; sendChat(); }
            continue;
          }
          if (!handshakeDone) {
            handshakeDone = true;
            if (parsed && parsed.error) {
              ws.close();
              reject(new Error(`Handshake error: ${parsed.error}`));
              return;
            }
            sendChat();
            continue;
          }
          handleMsg(parsed);
        }
      });

      ws.on("error", (err) => {
        const msg = err.message || "connection failed";
        if (!handshakeDone) reject(new Error(`WebSocket error: ${msg}`));
        else onError(new Error(`WebSocket error: ${msg}`));
      });

      ws.on("close", () => {
        clearAbort();
        onDone();
      });

      function sendChat() {
        const chatMsg = {
          arguments: [{
            source: "officeweb",
            clientCorrelationId: requestId,
            sessionId,
            optionsSets: [
              ...(process.env.M365_NO_CODE_INTERPRETER ? [] : CODE_INTERPRETER_OPTIONS_SETS),
            ],
            streamingMode: "ConciseWithPadding",
            spokenTextMode: "None",
            options: {},
            extraExtensionParameters: {},
            allowedMessageTypes: [
              "Chat", "Suggestion", "InternalSearchQuery", "Disengaged",
              "InternalLoaderMessage", "Progress", "RenderCardRequest", "SemanticSerp",
              "GenerateContentQuery", "SearchQuery", "ConfirmationCard", "DeveloperLogs",
              "EndOfRequest", "ReferencesListComplete", "GeneratedCode",
            ],
            sliceIds: [],
            threadLevelGptId: {},
            traceId: requestId,
            isStartOfSession: isFirst,
            clientInfo: {
              clientPlatform: "mcmcopilot-web",
              clientAppName: "Office",
              clientEntrypoint: "mcmcopilot-officeweb",
              clientSessionId: sessionId,
              clientAppType: "Web",
              deviceOS: "Linux",
              deviceType: "Desktop",
            },
            message: {
              author: "user",
              inputMethod: "Keyboard",
              text,
              entityAnnotationTypes: ["People", "File", "Event", "Email", "TeamsMessage"],
              requestId,
              locationInfo: { timeZoneOffset: 0, timeZone: "UTC" },
              locale: "en-gb",
              messageType: "Chat",
              experienceType: "Default",
              adaptiveCards: [],
              clientPreferences: {},
            },
            plugins: [{ Id: "BingWebSearch", Source: "BuiltIn" }],
            isSbsSupported: true,
            tone: getToneForModel(model),
            renderReferencesBehindEOS: true,
            disconnectBehavior: "continue",
          }],
          invocationId: "0",
          target: "chat",
          type: 4,
        };
        const metrics = {
          arguments: [{
            Timestamps: {
              ConnectionStart: new Date().toISOString(),
              UserInputStart: new Date().toISOString(),
              ConnectionEstablished: new Date().toISOString(),
              UserInputSubmit: new Date().toISOString(),
            },
          }],
          target: "Metrics",
          type: 1,
        };
        ws.send(JSON.stringify(chatMsg) + RS + JSON.stringify(metrics) + RS);
        resolve(stream);
      }

      function noteBot(m) {
        if (m.contentOrigin) contentOrigin = m.contentOrigin;
        if (m.messageType) messageType = m.messageType;
        if (typeof m.turnCount === "number") turnCountServer = m.turnCount;
        if (Array.isArray(m.scores)) {
          for (const s of m.scores) {
            if (typeof s?.component !== "string" || typeof s?.score !== "number") continue;
            if (!(s.component in maxScores) || s.score > maxScores[s.component]) {
              maxScores[s.component] = s.score;
            }
          }
        }
      }

      function handleMsg(raw) {
        const type = raw?.type;

        if (type === 6) { // ping
          ws.send(JSON.stringify({ type: 6 }) + RS);
          return;
        }
        if (type === 7) { // close
          if (raw.error) onError(new Error(`Server close: ${raw.error}`));
          ws.close();
          return;
        }
        if (type === 3) { // completion
          if (raw.error) onError(new Error(`Completion error: ${raw.error}`));
          ws.close();
          return;
        }
        if (type === 2) { // stream item — final conversation state
          const item = raw.item;
          if (item) {
            if (item.throttling) {
              throttle = {
                current: item.throttling.numUserMessagesInConversation,
                max: item.throttling.maxNumUserMessagesInConversation,
              };
            }
            for (const m of item.messages ?? []) {
              if (m.author !== "bot") continue;
              noteBot(m);
              if (m.text && !m.messageType) advance(m.text);
            }
          }
          ws.close();
          return;
        }
        if (type === 1 && raw.target === "update" && Array.isArray(raw.arguments)) {
          for (const arg of raw.arguments) {
            if (typeof arg?.writeAtCursor === "string") {
              advance(answer + arg.writeAtCursor);
              continue;
            }
            if (Array.isArray(arg?.messages)) {
              for (const m of arg.messages) {
                if (m.author === "bot") noteBot(m);
                if (m.author === "bot" && m.text && !m.messageType) advance(m.text);
              }
              continue;
            }
            if (arg?.throttling) {
              throttle = {
                current: arg.throttling.numUserMessagesInConversation,
                max: arg.throttling.maxNumUserMessagesInConversation,
              };
            }
          }
        }
      }
    });
  }
}
