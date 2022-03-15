# this is all preliminary draft, but reduced my mailbox size enough for now... :-)

param
(
    $fromYear = 2000,
    $toYear = 2018,
    $minMessageSizeMb = 8,
    $attachmentArchivePath = "$PSScriptRoot/attachments"
)

$ErrorActionPreference = "Stop"

$minMessageSizeByte = $minMessageSizeMb * 1024 * 1024;

#$token = (az account get-access-token --resource-type ms-graph --query accessToken -o tsv) # az cli access-token does not have the scopes we need, see https://goodworkaround.com/2020/09/14/easiest-ways-to-get-an-access-token-to-the-microsoft-graph/
$token = ""

# todo: Check token, automatically open graph explorer to sign in and copy if expired
# todo: Retry with exponential backoff 

$stringAsStream = [System.IO.MemoryStream]::new()
$writer = [System.IO.StreamWriter]::new($stringAsStream)
function GetHash ($value)
{
    $writer.write($value)
    $writer.Flush()
    $stringAsStream.Position = 0
    return (Get-FileHash -InputStream $stringAsStream).Hash
}

function ProcessMessage($MessageId)
{
    $messageUrl = "https://graph.microsoft.com/v1.0/me/messages/${MessageId}?`$expand=attachments(`$select=id,name)";
    $result = Invoke-WebRequest -Uri $messageUrl -Method GET -Headers @{"Authorization" = "Bearer $token"; "Content-Type" = "application/json" } 
    $message = $result.Content | ConvertFrom-Json
    if ($message.attachments.Length -gt 0)
    {
        $messageHash = GetHash -Value $message.id
        $mailAttachmentDirPath = [IO.Path]::Combine( $attachmentArchivePath, $messageHash)
        $mailAttachmentInfoPath = [IO.Path]::Combine( $mailAttachmentDirPath, "mail.json")
        New-Item -ItemType Directory -Force -Path $mailAttachmentDirPath | Out-Null
        Set-Content -Path $mailAttachmentInfoPath -Value $result.Content  #"{`n""subject"": ""$($message.subject)"",`n""id"": ""$($message.id)""`n}"
        foreach ($attachment in $message.attachments)
        {
            $attachmentId = $attachment.id
            $attachmentName = $attachment.name
            $attachmentUrl = $messageUrl = "https://graph.microsoft.com/v1.0/me/messages/${MessageId}/attachments/${attachmentId}";
            $attachmentValueUrl = $messageUrl = "${attachmentUrl}/`$value";
            [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {$attachmentName = $attachmentName.replace($_,'.')}
            $attachmentPath = [IO.Path]::Combine( $mailAttachmentDirPath, $attachmentName)
            $ProgressPreference = "SilentlyContinue"
            Invoke-WebRequest -Method GET -Uri $attachmentValueUrl -OutFile $attachmentPath -Headers @{"Authorization" = "Bearer $token" } | Out-Null
            Invoke-WebRequest -Method DELETE -Uri $attachmentUrl -Headers @{"Authorization" = "Bearer $token" } | Out-Null
        }
    }

}



$url = "https://graph.microsoft.com/v1.0/me/messages?`$filter=singleValueExtendedProperties/Any(ep%3A%20ep%2Fid%20eq%20'Long%200x0E08'%20and%20cast(%20ep%2Fvalue,Edm.Int64)%20gt%20$minMessageSizeByte) and sentDateTime ge $fromYear-01-01T00:00:00Z and sentDateTime le $($toYear + 1)-01-01T00:00:00Z&`$expand=singleValueExtendedProperties(`$filter=Id eq 'LONG 0x0E08')&`$orderby=sentDateTime&`$select=subject,hasAttachments,sentDateTime,id&`$count=true"
$result = Invoke-WebRequest -Uri $url -Method GET -Headers @{"Authorization" = "Bearer $token"; "Content-Type" = "application/json" } 
$messageList = $result.Content | ConvertFrom-Json

Write-Host "Found $($messageList."@odata.count") messages"

$messages = @()
$nextUrl = $url
while ($nextUrl)
{
    foreach ($message in $messageList.value)
    {
        $messages += $message
    }

    $nextUrl = $messageList."@odata.nextLink"
    if ($nextUrl)
    {
        $result = Invoke-WebRequest -Uri $nextUrl -Method GET -Headers @{"Authorization" = "Bearer $token"; "Content-Type" = "application/json" }  
        $messageList = $result.Content | ConvertFrom-Json
    }
}

$processedMessageCount = 0
foreach ($message in $messages)
{
    $subject = $message.subject
    $id = $message.id
    $size = $message.singleValueExtendedProperties[0].value
    Write-Host "Processing message $($processedMessageCount + 1) [size: $(($size/1024/1024).ToString('00.00')) MB, subject: $subject]"    
    $processedMessageCount += 1
    ProcessMessage -MessageId $id
}