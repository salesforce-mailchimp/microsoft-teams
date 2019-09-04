param (
    <#
        Office 365 Credentials Name
        The name of the saved Office 365 administrative credentials.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$office365CredentialsName
)
@(
    "Get-SavedCredentials"
) | ForEach-Object -Process {
    . "$($PSScriptRoot)\..\..\functions\$($_).ps1"
}

# Retrieve the credentials object, username and password
$office365CredentialsObject = Get-SavedCredentials -CredentialsName $office365CredentialsName -UserCredentials
if ($null -eq $office365CredentialsObject) {
    Write-Warning "Failed to retrieve the Office 365 credentials '$($office365CredentialsName)'. Please check the name and try again."
    exit
}
$office365Username = $office365CredentialsObject.username
$office365Password = $office365CredentialsObject.password

$credentials = [PSCredential]::new(
    $office365Username,
    ($office365Password | ConvertTo-SecureString -AsPlainText -Force)
)

# Connecting
Write-Information "Connecting"
Connect-MicrosoftTeams -Credential $credentials

Connect-MicrosoftTeamsAdminAccount -endpoint $Office365AdministrativeCredentials
$token = Get-MicrosoftGraphAuthenticationToken -endpoint $MicroSoftGraphCredentials

Write-Information "Retrieving teams."
$allTeams = Get-Team

$numberOfPrivateTeams = 0
$numberOfPublicTeams = 0

$allChannels = @()

$allUsers = @()

# To check
$maxNumberOfTeams = 250

foreach ($team in $allTeams) {
    if ($team.Visibility -eq "Public") {
        $numberOfPublicTeams++
    }
    else {
        $numberOfPrivateTeams++
    }

    Write-Information "Retrieving file information for '$($team.DisplayName)'"
    $invokeWebRequestParams = @{
        Uri     = "https://graph.microsoft.com/v1.0/groups/$($team.GroupId)/drive/root/children"
        Method  = "GET"
        Headers = @{
            Authorization = "Bearer $($token)"
        }
    }
    $response = Invoke-RestMethod @invokeWebRequestParams
    $allFileInformation = $response.Value

    Write-Information "Retrieving channels in '$($team.DisplayName)'."
    $teamChannels = Get-TeamChannel -GroupId $team.GroupId
    foreach ($channel in $teamChannels) {
        $channel | Add-Member -NotePropertyName "Team" -NotePropertyValue $team.DisplayName
        $allChannels += $channel

        $teamFileInformation = $allFileInformation | Where-Object {$_.name -eq $channel.DisplayName}
        $channel | Add-Member -NotePropertyName "TotalFileSize" -NotePropertyValue $teamFileInformation.size
        $channel | Add-Member -NotePropertyName "FileLastModifiedTime" -NotePropertyValue $teamFileInformation.lastModifiedDateTime

        Write-Information "Retrieving chat information for channel $($channel.DisplayName)"
        $lastChatTime = (Get-Date).AddYears(-100)
        $invokeWebRequestParams = @{
            Uri     = "https://graph.microsoft.com/beta/teams/$($team.GroupId)/channels/$($channel.Id)/messages"
            Method  = "GET"
            Headers = @{
                Authorization = "Bearer $($adminToken)"
            }
        }

        try{
            $response = Invoke-RestMethod @invokeWebRequestParams
            $chatCount = $response.'@odata.count'
            $messages = $response.value
        }
        catch {
            # 404 means there isn't any chat message
        }

        foreach($message in $messages) {
            $messageTime = Get-Date -Date $message.createdDateTime
            if($messageTime -gt $lastChatTime) {
                $lastChatTime = $messageTime
            }
            $invokeWebRequestParams = @{
                Uri     = "https://graph.microsoft.com/beta/teams/$($team.GroupId)/channels/$($channel.Id)/messages/$($message.Id)/replies"
                Method  = "GET"
                Headers = @{
                    Authorization = "Bearer $($adminToken)"
                }
            }
            $response = Invoke-RestMethod @invokeWebRequestParams

            if($response.'@odata.count' -gt 0) {
                $replyTime = Get-Date -Date $response.value[0].createdDateTime
                if($replyTime -gt $lastChatTime) {
                    $lastChatTime = $replyTime
                }
            }
        }

        $channel | Add-Member -NotePropertyName "LastChatTime" -NotePropertyValue $lastChatTime -Force
    }
    #$team | Add-Member -NotePropertyName "Channels" -NotePropertyValue $teamChannels -Force
    $team | Add-Member -NotePropertyName "NumberOfChannels" -NotePropertyValue $teamChannels.Length -Force
    Write-Information "Retrieving users in '$($team.DisplayName)'."
    $teamUsers = Get-TeamUser -GroupId $team.GroupId
    foreach ($user in $teamUsers) {
        $user | Add-Member -NotePropertyName "Team" -NotePropertyValue $team.DisplayName
        $allUsers += $user
    }
    #$externalUsers = $teamUsers | Where-Object {$_.User -like "*#EXT#*"}
    #$team | Add-Member -NotePropertyName "ExternalUsers" -NotePropertyValue $externalUsers -Force
    #$team | Add-Member -NotePropertyName "NumberOfExternalUsers" -NotePropertyValue $externalUsers.Length -Force


}

# Confirmed
$maxNumberOfChannels = 201

### Report Data
$numberOfTeams = $allTeams.Length
$externalUsers = $allUsers.User | Select-Object -Unique | Where-Object { $_ -like "*#EXT#*" }
$numberOfExternalUsers = $externalUsers.Length

$1To50 = ($allTeams | Where-Object { $_.NumberOfChannels -le 50 }).Length
$51To100 = ($allTeams | Where-Object { $_.NumberOfChannels -le 100 -and $_.NumberOfChannels -ge 51 }).Length
$101To150 = ($allTeams | Where-Object { $_.NumberOfChannels -le 150 -and $_.NumberOfChannels -ge 101 }).Length
$151Plus = ($allTeams | Where-Object { $_.NumberOfChannels -ge 151 }).Length
$numberOfChannelsDistribution = @($1To50, $51To100, $101To150, $151Plus)

$templateFilePath = "$($PSScriptRoot)\TeamsReportTemplate.xlsx"
$reportFilePath = "$($PSScriptRoot)\TeamsReport.xlsx"

$templateFilePath = "C:\Users\yshen\Desktop\TeamsReportTemplate.xlsx"
$reportFilePath = "C:\Users\yshen\DesktopTeamsReport.xlsx"

Copy-Item -Path $templateFilePath -Destination $reportFilePath

$excelPackage = $allUsers | Export-Excel -Path $reportFilePath -WorksheetName "Users" -PassThru
$excelPackage = $allChannels | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Channels" -PassThru
$excelPackage = $allTeams | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Teams" -PassThru
$excelPackage = $numberOfChannelsDistribution | Export-Excel -ExcelPackage $excelPackage -WorksheetName "NumberOfChannels" -PassThru -StartRow 2 -StartColumn 2
$excelPackage = $numberOfTeams | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Dashboard" -PassThru -StartRow 5 -StartColumn 2
$excelPackage = $numberOfPublicTeams | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Dashboard" -PassThru -StartRow 5 -StartColumn 4
$excelPackage = $numberOfPrivateTeams | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Dashboard" -PassThru -StartRow 5 -StartColumn 6
$excelPackage = $numberOfExternalUsers | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Dashboard" -PassThru -StartRow 5 -StartColumn 8

Close-ExcelPackage -ExcelPackage $excelPackage -Show
Write-Information "Finished."

# Store the Excel workbook base64 encoded bytes in a file in the instance scope
$excelWorkbookContentsFileName = "$(New-Guid).txt"
$context.SaveInstanceText($excelWorkbookContentsFileName, (Convert-FileToBase64EncodedBytes -Path $reportFilePath))
$context.Outputs.teamsReportContentsFileName = $excelWorkbookContentsFileName