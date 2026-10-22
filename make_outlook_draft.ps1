param(
  [Parameter(Mandatory=$true)][string]$HtmlBody,
  [Parameter(Mandatory=$true)][string]$Attach,
  [Parameter(Mandatory=$true)][string]$Subject,
  [Parameter(Mandatory=$true)][string]$To
)
# Opens an Outlook draft (NOT sent): HTML body pasted from file, report HTML attached, recipients in To.
$ErrorActionPreference = "Stop"
$body = [System.IO.File]::ReadAllText((Resolve-Path $HtmlBody).Path, [System.Text.Encoding]::UTF8)
$ol = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)          # olMailItem
$mail.To = $To
$mail.Subject = $Subject
$mail.HTMLBody = $body
[void]$mail.Attachments.Add((Resolve-Path $Attach).Path)
[void]$mail.Recipients.ResolveAll()
$mail.Display()
"draft opened: $Subject  | To: $($mail.To) | attachments: $($mail.Attachments.Count)"
