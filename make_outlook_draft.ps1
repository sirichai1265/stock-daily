param(
  [Parameter(Mandatory=$true)][string]$HtmlBody,
  [Parameter(Mandatory=$true)][string]$Attach,
  [Parameter(Mandatory=$true)][string]$Subject,
  [Parameter(Mandatory=$true)][string]$To,
  [string]$InlineImage = "",
  [string]$Cid = "halbanner"
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
if ($InlineImage -ne "") {
  # inline picture referenced from the body as <img src="cid:...">; hidden so it is not listed as a file attachment
  $img = $mail.Attachments.Add((Resolve-Path $InlineImage).Path, 1, 0)
  $img.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", $Cid)
  $img.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", $true)
}
[void]$mail.Recipients.ResolveAll()
$mail.Display()
"draft opened: $Subject  | To: $($mail.To) | attachments: $($mail.Attachments.Count)"
