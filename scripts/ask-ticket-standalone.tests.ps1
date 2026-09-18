#!/usr/bin/env pwsh
# Smoke tests for ask-ticket-standalone.ps1 (no Pester required).
#   pwsh -NoProfile -File scripts/ask-ticket-standalone.tests.ps1
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here "ask-ticket-standalone.ps1")

$script:Failures = 0
function Assert-StrEqual {
    param([string]$Name, [string]$Expected, [string]$Actual)
    if ($Expected -cne $Actual) {
        Write-Host "FAIL $Name"
        Write-Host "  expected: [$Expected]"
        Write-Host "  actual:   [$Actual]"
        $script:Failures++
    } else { Write-Host "ok   $Name" }
}
function Assert-True {
    param([string]$Name, [bool]$Condition)
    if (-not $Condition) { Write-Host "FAIL $Name"; $script:Failures++ } else { Write-Host "ok   $Name" }
}


# --- payload budget --------------------------------------------------------
Assert-StrEqual "limit under budget" "short" (Limit-Text -Text "short" -MaxChars 100)
$long = Limit-Text -Text ("H" * 200 + "T" * 200) -MaxChars 100
Assert-True "limit respects budget" ($long.Length -le 100)
Assert-True "limit keeps head" ($long.StartsWith("H"))
Assert-True "limit keeps tail" ($long.EndsWith("T"))

# --- prompt assembly -------------------------------------------------------
$withSystem = New-PromptText -Context "DATA" -Instruction "do it" -System "Be terse."
Assert-True "system prompt first" ($withSystem.StartsWith("Be terse.`n`n"))
Assert-True "context block present" ($withSystem.Contains("<<<CONTEXT`nDATA`nCONTEXT>>>"))
Assert-True "instruction last" ($withSystem.EndsWith("do it"))
Assert-True "empty system omitted" ((New-PromptText -Context "DATA" -Instruction "do it" -System "").StartsWith("The text between"))

# --- stream folding --------------------------------------------------------
$r = Add-StreamText -Answer "" -Next "hi"
Assert-StrEqual "fold snapshot" "hi" $r.Answer
Assert-StrEqual "fold snapshot emit" "hi" $r.Emit
$r = Add-StreamText -Answer "hi" -Next "hi there"
Assert-StrEqual "fold delta" "hi there" $r.Answer
Assert-StrEqual "fold delta emit" " there" $r.Emit
$r = Add-StreamText -Answer "hi there" -Next "hi"
Assert-StrEqual "fold no shrink" "hi there" $r.Answer
Assert-True "fold no emit" ($null -eq $r.Emit)

# --- numeric mapping -------------------------------------------------------
Assert-StrEqual "map urgency" "Low" (Get-MappedName -Key "urgency" -Value "1")
Assert-StrEqual "map status" "Open" (Get-MappedName -Key "status" -Value "2")
Assert-StrEqual "map already named" "High" (Get-MappedName -Key "priority" -Value "High")

# --- rendering -------------------------------------------------------------
$ticket = [pscustomobject]@{
    id = 10100; display_id = 10100; subject = "Printer"; status = 2; urgency = 1
    created_at = "2026-08-01T10:00:00Z"; description_text = "It is broken"
}
$conv = @([pscustomobject]@{
    id = 1; user_id = 2100; incoming = $true
    created_at = "2026-08-01T10:30:00Z"; body_text = "<p>Please fix</p>"
})
$md = Format-TicketContents -Ticket $ticket -Conversations $conv
Assert-True "render title" ($md.StartsWith("# Ticket #10100 — Printer"))
Assert-True "render status name" ($md.Contains("Status     : Open"))
Assert-True "render urgency name" ($md.Contains("Urgency    : Low"))
Assert-True "render conversation" ($md.Contains("### 2100 (incoming, 2026-08-01T10:30:00Z)"))
Assert-True "render body, html stripped" ($md.Contains("Please fix") -and -not $md.Contains("<p>"))

# --- native base64url JWT decode ------------------------------------------
# Encode without Base64Url so this test also runs on .NET 8 / pwsh 7.4.
$jwtPayload = [System.Convert]::ToBase64String(
    [System.Text.Encoding]::UTF8.GetBytes('{"oid":"abc","tid":"xyz"}')).TrimEnd('=').Replace('+', '-').Replace('/', '_')
$claims = ConvertFrom-JwtPayload -Token "header.$jwtPayload.sig"
Assert-StrEqual "jwt oid" "abc" $claims.oid
Assert-StrEqual "jwt tid" "xyz" $claims.tid

# --- native field access ---------------------------------------------------
$obj = [pscustomobject]@{ a = 1; nested = [pscustomobject]@{ name = "x" } }
Assert-True "field present" ((Get-FieldRaw -Object $obj -Key "a") -eq 1)
Assert-True "field missing is null" ($null -eq (Get-FieldRaw -Object $obj -Key "nope"))
Assert-StrEqual "nested value" "x" (Get-FieldValue -Object $obj.nested -Key "name")
Assert-StrEqual "missing value empty" "" (Get-FieldValue -Object $obj -Key "nope")

# --- attachment normalisation ---------------------------------------------
$raw = [pscustomobject]@{
    id = 21117119076; name = "83e5.jpeg"; content_type = "image/jpeg"; size = 279272
    attachment_url = "https://x/a.jpeg"; uploaded_by = "x"
}
$att = ConvertTo-NormalizedAttachment -Raw $raw -Source "ticket"
Assert-StrEqual "att name" "83e5.jpeg" $att.Name
Assert-StrEqual "att type" "image/jpeg" $att.ContentType
Assert-True "att size" ($att.Size -eq 279272)
Assert-StrEqual "att url" "https://x/a.jpeg" $att.Url
Assert-StrEqual "att source" "ticket" $att.Source

# Real captured shape also carries `canonical_url`/`canonical_path`. Prefer the
# signed (cookie-free) `attachment_url`; fall back to `canonical_url` when absent.
$real = ConvertTo-NormalizedAttachment -Raw ([pscustomobject]@{
    id = 21113132612; name = "Adhar card.pdf"; content_type = "application/pdf"; size = 192838
    attachment_url = "https://x.attachments.freshservice.com/data/a.pdf?Signature=abc"
    canonical_url = "https://x/helpdesk/attachments/21113132612"
}) -Source "ticket"
Assert-StrEqual "att prefers signed url" "https://x.attachments.freshservice.com/data/a.pdf?Signature=abc" $real.Url
$canonOnly = ConvertTo-NormalizedAttachment -Raw ([pscustomobject]@{
    name = "Adhar card.pdf"; content_type = "application/pdf"
    canonical_url = "https://x/helpdesk/attachments/21113132612"
}) -Source "ticket"
Assert-StrEqual "att canonical url fallback" "https://x/helpdesk/attachments/21113132612" $canonOnly.Url

# --- attachment plan (size + count limits) --------------------------------
$atts = @(
    (ConvertTo-NormalizedAttachment -Raw ([pscustomobject]@{ name = "a.png"; content_type = "image/png"; size = 100 }) -Source "ticket"),
    (ConvertTo-NormalizedAttachment -Raw ([pscustomobject]@{ name = "b.log"; content_type = "text/plain"; size = 100 }) -Source "ticket"),
    (ConvertTo-NormalizedAttachment -Raw ([pscustomobject]@{ name = "huge.log"; content_type = "text/plain"; size = 999999 }) -Source "ticket"),
    (ConvertTo-NormalizedAttachment -Raw ([pscustomobject]@{ name = "c.pdf"; content_type = "application/pdf"; size = 100 }) -Source "ticket")
)
# NB: do not name these $plan — the script's [switch]$Plan parameter creates a
# type-constrained $Plan variable that dot-sourcing shares with this file.
$attPlan = Get-AttachmentPlan -Attachments $atts -MaxInlineFileBytes 1000 -MaxInlineFiles 20 -TextExtensions @(".log")
Assert-True "plan inlines small text" ($attPlan.Inline.Count -eq 1 -and $attPlan.Inline[0].Name -eq "b.log")
Assert-True "plan lists oversize + binary" (`
    ($attPlan.ListedOnly | ForEach-Object { $_.Name }) -contains "huge.log" -and `
    ($attPlan.ListedOnly | ForEach-Object { $_.Name }) -contains "c.pdf")
$attPlan2 = Get-AttachmentPlan -Attachments $atts -MaxInlineFileBytes 1000 -MaxInlineFiles 0 -TextExtensions @(".log")
Assert-True "count limit respected" ($attPlan2.Inline.Count -eq 0)

# --- manifest --------------------------------------------------------------
$manifest = Format-AttachmentManifest -Attachments @($atts[1])
Assert-True "manifest heading" ($manifest.StartsWith("## Attachments"))
Assert-True "manifest has row" ($manifest.Contains("b.log") -and $manifest.Contains("text/plain") -and $manifest.Contains("100"))
Assert-StrEqual "manifest none" "## Attachments`n`n(none)" (Format-AttachmentManifest -Attachments @())

# --- inlined attachment section -------------------------------------------
$sec = ConvertTo-AttachmentSection -Attachment $att -Bytes ([System.Text.Encoding]::UTF8.GetBytes("line one`nline two")) -MaxChars 1000
Assert-True "inline section heading" ($sec.Contains("### Attachment: 83e5.jpeg"))
Assert-True "inline section fence is open+newline" ($sec.Contains('```' + [char]10 + 'line one'))
Assert-True "inline section fence closes" ($sec.TrimEnd().EndsWith('```'))
$secNull = ConvertTo-AttachmentSection -Attachment $att -Bytes $null -MaxChars 100
Assert-True "inline missing download noted" ($secNull.Contains("(could not download)"))

# --- message packing: the long-ticket path ---------------------------------
$chunks = Split-TextIntoChunks -Text ("line`n" * 500) -MaxChars 100
Assert-True "chunks exist" ($chunks.Count -gt 1)
Assert-True "chunks respect budget" (($chunks | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum -le 100)
$hard = Split-TextIntoChunks -Text ("X" * 250) -MaxChars 100
Assert-True "hard split long line" ($hard.Count -eq 3)
Assert-True "hard split preserves length" ((($hard | ForEach-Object { $_.Length }) | Measure-Object -Sum).Sum -ge 250)

# a long ticket: 200 conversation sections, some oversized
$many = @()
$many += "# Ticket #1 — big"
for ($i = 1; $i -le 200; $i++) { $many += "### conv $i`n`n" + ("body $i " * 30) }
$many += ("Z" * 5000)
$packed = Group-ContextSections -Sections $many -MaxChars 2000
Assert-True "many messages" ($packed.Count -gt 10)
Assert-True "every message within budget" (($packed | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum -le 2000)
$joined = $packed -join "`n"
Assert-True "first section kept" ($joined.Contains("# Ticket #1 — big"))
Assert-True "first conv kept" ($joined.Contains("### conv 1"))
Assert-True "last conv kept" ($joined.Contains("### conv 200"))
Assert-True "oversized tail kept" ($joined.Contains("ZZZZZ"))
Assert-True "nothing dropped" ($joined.Contains("body 200"))

# single short section stays one message
Assert-True "short ticket is one message" ((Group-ContextSections -Sections @("a", "b") -MaxChars 1000).Count -eq 1)

# regression: a single packed message must still be an ARRAY. PowerShell
# unrolls single-element returns, which made $messages[0] the FIRST CHARACTER
# of the context (the whole ticket collapsed to "#").
$oneMsg = Group-ContextSections -Sections @("hello world") -MaxChars 1000
Assert-True "single packed message is an array" ($oneMsg -is [System.Array])
Assert-StrEqual "single packed message keeps full text" "hello world" $oneMsg[0]
$oneChunk = Split-TextIntoChunks -Text "short" -MaxChars 100
Assert-True "single chunk is an array" ($oneChunk -is [System.Array])

# --- conversation mode -----------------------------------------------------
Assert-StrEqual "resolve explicit wins" "conv-explicit" (Resolve-ConversationId -Explicit "conv-explicit" -UseLast $true -Last "conv-last")
Assert-StrEqual "resolve last" "conv-last" (Resolve-ConversationId -Explicit "" -UseLast $true -Last "conv-last")
Assert-StrEqual "resolve new when nothing remembered" "conv-new" (Resolve-ConversationId -Explicit "" -UseLast $true -Last "" -NewId { "conv-new" })
Assert-StrEqual "resolve new when not asked for last" "conv-new2" (Resolve-ConversationId -Explicit "" -UseLast $false -Last "conv-last" -NewId { "conv-new2" })
Assert-True "resolve mints a guid by default" ((Resolve-ConversationId) -match '^[0-9a-f-]{36}$')

$tmpUrl = New-ChatHubUrl -Oid "OID" -Tid "TID" -SessionId "SID" -ConversationId "CID" -RequestId "RID" -Token "TOK" -Temporary $true
Assert-True "temp url disables memory" ($tmpUrl.Contains("disableMemory=1"))
Assert-True "url targets chathub" ($tmpUrl.StartsWith("wss://substrate.office.com/m365Copilot/Chathub/OID@TID?"))
Assert-True "url carries conversation" ($tmpUrl.Contains("ConversationId=CID"))
Assert-True "url carries session" ($tmpUrl.Contains("X-SessionId=SID"))
Assert-True "url carries token" ($tmpUrl.Contains("access_token=TOK"))
$persUrl = New-ChatHubUrl -Oid "OID" -Tid "TID" -SessionId "SID" -ConversationId "CID" -RequestId "RID" -Token "TOK" -Temporary $false
Assert-True "persistent url keeps memory" (-not $persUrl.Contains("disableMemory"))
Assert-True "temporary is the default" ((New-ChatHubUrl -Oid "O" -Tid "T" -SessionId "S" -ConversationId "C" -RequestId "R" -Token "K").Contains("disableMemory=1"))

# --- default config --------------------------------------------------------
Assert-True "temporary chat is on by default" ($Config.Temporary -eq $true)
Assert-True "system prompt is set by default" ($Config.SystemPrompt.Trim().Length -gt 0)
Assert-True "system prompt asks for plain text" ($Config.SystemPrompt -match 'plain text')
$appliedSystem = New-PromptText -Context "CTX" -Instruction "digest" -System $Config.SystemPrompt
Assert-True "system prompt is prepended" ($appliedSystem.StartsWith($Config.SystemPrompt.Trim()))

# a plain frame carries no annotations (text only)
$frPlain = New-ChatFrame -RequestId "R" -SessionId "S" -Text "hi" -IsFirstTurn $true | ConvertFrom-Json
Assert-True "frame omits annotations" ($null -eq $frPlain.arguments[0].message.messageAnnotations)
Assert-True "frame omits image options" (-not ($frPlain.arguments[0].optionsSets -contains "gptvnorm2048"))

if ($script:Failures -gt 0) { Write-Host "$($script:Failures) PowerShell test(s) failed"; exit 1 }
Write-Host "all PowerShell tests passed"
exit 0
