param(
    [Parameter(Mandatory=$true)]
    [string]$HtmlFile,

    [Parameter(Mandatory=$true)]
    [string]$To,

    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [string]$Cc = ""
)

if (-not (Test-Path $HtmlFile)) {
    Write-Error "HTML file not found: $HtmlFile"
    exit 1
}

$html = Get-Content -Path $HtmlFile -Raw -Encoding UTF8

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $To
    if ($Cc -ne "") {
        $mail.CC = $Cc
    }
    $mail.Subject = $Subject
    $mail.HTMLBody = $html
    $mail.Display()
}
catch {
    Write-Error "Could not open Outlook draft: $_"
    exit 1
}
