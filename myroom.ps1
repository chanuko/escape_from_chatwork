Param([string]$chatworkApiKey)

$headers = @{
    "X-ChatWorkToken" = $chatworkApiKey
}


# マイチャットのroom_idを抽出する
$response = Invoke-RestMethod -Headers $headers "https://api.chatwork.com/v2/rooms"
$roomId = $response | Where-Object type -EQ "my" | Select-Object -ExpandProperty room_id

# マイチャットから直近100件を取得する ※API上100件以上は取れない。 
$messages = Invoke-RestMethod -Headers $headers "https://api.chatwork.com/v2/rooms/$roomId/messages?force=1" 

# 整形してCSV出力
$messages | Select-Object `
    @{Name="datetime"; Expression={ConvertFromUnixtoJst -unixTIme $_.send_time}} `
    , @{Name="message"; Expression={$_.body}} `
    | ConvertTo-Csv `
    | Out-File -Encoding default -FilePath "$HOME\Desktop\mychat_log.csv"


<#
.SYNOPSIS
#

.DESCRIPTION
Unix時間をJSTに変換します。
UTC+9変換はベタ打ち

.PARAMETER unixTime
1576485433

.EXAMPLE
ConvertFromUnixtoJst -unixTIme 1576485433

.NOTES
General notes
#>
function ConvertFromUnixtoJst {
    param (
        [int]$unixTime
    )
    #UTC+9変換はベタ打ち
    $currentTime = (Get-Date -Year 1970 -Month 1 -Date 1 -Hour 9 -Second 0) + [System.TimeSpan]::FromSeconds($unixTime)
    return $currentTime
}