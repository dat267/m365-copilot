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

# --- redaction -------------------------------------------------------------
Assert-StrEqual "redact email" "Email [redacted-email] now" (Redact-PII -Text "Email omar.saleh@example.com now")
Assert-StrEqual "redact NANP" "call [redacted-phone]" (Redact-PII -Text "call 555-123-4567")
Assert-StrEqual "redact parenthesised" "call [redacted-phone]" (Redact-PII -Text "call (555) 123-4567")
Assert-StrEqual "redact international" "call [redacted-phone]" (Redact-PII -Text "call +44 20 7946 0958")
Assert-StrEqual "redact UK mobile" "call [redacted-phone]" (Redact-PII -Text "call 07123456789")
Assert-StrEqual "redact UK spaced" "call [redacted-phone]" (Redact-PII -Text "call 020 7946 0958")
Assert-StrEqual "keep dates/ids" "Created 2026-08-01T10:30:00Z, ticket #10100, ref INC0012345" `
    (Redact-PII -Text "Created 2026-08-01T10:30:00Z, ticket #10100, ref INC0012345")

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

if ($script:Failures -gt 0) { Write-Host "$($script:Failures) PowerShell test(s) failed"; exit 1 }
Write-Host "all PowerShell tests passed"
exit 0