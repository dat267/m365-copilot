// Attachment handling for the ticket-context flow.
//
// M365 Copilot accepts only a few images per message (see
// `MAX_IMAGES_PER_MESSAGE` in chat-api.js), so a ticket with many screenshots
// cannot be delivered in one turn. These helpers keep that decision explicit
// and testable, and make sure nothing is silently dropped.
//
// Pure logic (no I/O) — tested with `node --test`.

/** Extensions M365 treats as images (drives the per-message image limit). */
const IMAGE_EXTENSIONS = new Set([
  "png",
  "jpg",
  "jpeg",
  "jfif",
  "pjpeg",
  "pjp",
  "gif",
  "bmp",
  "webp",
  "tif",
  "tiff",
  "heic",
  "heif",
  "svg",
]);

function extensionOf(name) {
  const s = String(name ?? "");
  const dot = s.lastIndexOf(".");
  if (dot <= 0 || dot === s.length - 1) return "";
  return s.slice(dot + 1).toLowerCase();
}

/** True when M365 will treat this attachment as an image (counts to the 3/message cap). */
export function isImage({ name, contentType } = {}) {
  const type = String(contentType ?? "").trim().toLowerCase();
  if (type) return type.startsWith("image/");
  return IMAGE_EXTENSIONS.has(extensionOf(name));
}

/**
 * Splits a ticket's attachments into images and non-image files — useful for
 * deciding HOW to deliver each (inline vs OneDrive), NOT for how many fit.
 *
 * For packing a message, use `planAttachmentBatches`: M365 caps images and files
 * TOGETHER at 3 per message (verified: 1 image + 2 files = 3 accepted; a 4th of
 * either kind is refused, and 3 images refuse any file).
 */
export function planImageBatches(attachments = [], { maxPerMessage = 3 } = {}) {
  const images = [];
  const files = [];
  for (const a of attachments) (isImage(a) ? images : files).push(a);

  const batches = [];
  for (let i = 0; i < images.length; i += maxPerMessage) {
    batches.push(images.slice(i, i + maxPerMessage));
  }
  return { images, files, batches };
}

/**
 * Packs attachments, in order, into per-message batches that respect the SHARED
 * cap (images and files together). Nothing is dropped or reordered.
 *
 * @returns {object[][]} e.g. 7 attachments at 3 per message -> [3, 3, 1]
 */
export function planAttachmentBatches(attachments = [], { maxPerMessage = 3 } = {}) {
  const batches = [];
  for (let i = 0; i < attachments.length; i += maxPerMessage) {
    batches.push(attachments.slice(i, i + maxPerMessage));
  }
  return batches;
}

/**
 * Renders the attachment inventory as Markdown for the prompt. Always emitted,
 * so the model knows the complete set even when only a few can be delivered.
 */
export function renderAttachmentManifest(attachments = []) {
  const out = ["## Attachments", ""];
  if (attachments.length === 0) {
    out.push("(none)");
    return out.join("\n");
  }
  out.push("| # | Name | Type | Size (bytes) | Source | URL |", "|---|---|---|---|---|---|");
  attachments.forEach((a, i) => {
    const cells = [
      String(i + 1),
      a.name ?? "",
      a.contentType ?? "",
      a.size == null ? "" : String(a.size),
      a.source ?? "",
      a.url ?? "",
    ].map((c) => String(c).replace(/\|/g, "\\|"));
    out.push(`| ${cells.join(" | ")} |`);
  });
  return out.join("\n");
}
