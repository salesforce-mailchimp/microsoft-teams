param (
    <#
        Office 365 Credentials Name
        The name of the saved Office 365 administrative credentials.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$office365CredentialsName,

    # Microsoft Graph Credentials Name
    # The name of the Microsoft Graph credentials.
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$microsoftGraphCredentialsName
)
@(
    "Get-SavedCredentials"
) | ForEach-Object -Process {
    . "$($PSScriptRoot)\..\..\functions\$($_).ps1"
}

function Get-MicrosoftGraphAuthenticationToken {
    param(
        $applicationId,
        $clientSecret,
        $username,
        $password
    )

    $invokeWebRequestParams = @{
        Uri     = "https://login.microsoftonline.com/common/oauth2/token"
        Method  = "POST"
        Headers = @{
            'Content-Type' = "application/x-www-form-urlencoded"
        }
        Body    = @{
            grant_type    = "password"
            resource      = "https://graph.microsoft.com"
            username      = $username
            password      = $password
            client_id     = $applicationId
            client_secret = $clientSecret
        }
    }
    $response = Invoke-RestMethod @invokeWebRequestParams

    return $response.access_token
}

function Update-Statistics {
    param(
        $Statistics,
        $Key,
        $Value
    )
    $Statistics.Add([PSCustomObject]@{
            Key   = $Key
            Value = $Value
        }) | Out-Null
}

function Get-MessagesJson {
    param(
        $Messages,
        $TeamId,
        $ChannelId,
        $AdminToken
    )
    $messagesJson = @()
    foreach ($message in $messages) {

        $repliesJson = @()
        $invokeWebRequestParams = @{
            Uri     = "https://graph.microsoft.com/beta/teams/$($team.GroupId)/channels/$($channel.Id)/messages/$($message.Id)/replies"
            Method  = "GET"
            Headers = @{
                Authorization = "Bearer $($adminToken)"
            }
        }
        $repliesResponse = Invoke-RestMethod @invokeWebRequestParams
        foreach ($reply in $repliesResponse.value) {
            $repliesJson += ($reply | ConvertTo-Json -Depth 10)
        }
        $message | Add-Member -NotePropertyName "Replies" -NotePropertyValue $repliesJson -Force
        $messagesJson += ($message | ConvertTo-Json -Depth 10)
    }
    return $messagesJson
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

$microsoftGraphCredentialsObject = Get-SavedCredentials -CredentialsName $microsoftGraphCredentialsName -ApplicationCredentials
if ($null -eq $microsoftGraphCredentialsObject) {
    Write-Error "Failed to retrieve the Microsoft Graph credentials '$($microsoftGraphCredentialsName)'. Please check the name and try again."
    exit
}
$applicationID = $microsoftGraphCredentialsObject.applicationId
$clientSecret = $microsoftGraphCredentialsObject.clientSecret
$domain = $microsoftGraphCredentialsObject.domain

# Connect to platforms
Write-Information "Connecting"
Connect-MicrosoftTeams -Credential $credentials
Connect-AzureADAdminAccount -Username $office365Username -Password ($office365Password | ConvertTo-SecureString -AsPlainText -Force)
$adminToken = Get-MicrosoftGraphAuthenticationToken -Username $office365Username -Password $office365Password -ApplicationId $applicationID -ClientSecret $clientSecret

Write-Information "Retrieving teams."
$allTeams = Get-Team

Write-Information "Retrieving Azure AD users."
$allAzureADUsers = Get-AzureADUser

# Initialize variables
$numberOfPrivateTeams = 0
$numberOfPublicTeams = 0
$allChannels = @()
$allUsers = @()

# To check
$maxNumberOfTeams = 250

$now = Get-Date
foreach ($team in $allTeams) {
    # Process error (To investigate the cause)
    if ([String]::IsNullOrWhiteSpace($team.GroupId)) {
        Write-Warning $team
        continue
    }

    # Update the count
    if ($team.Visibility -eq "Public") {
        $numberOfPublicTeams++
    }
    else {
        $numberOfPrivateTeams++
    }

    # Retrieve file information
    Write-Information "Retrieving file information for '$($team.DisplayName)'"
    $invokeWebRequestParams = @{
        Uri     = "https://graph.microsoft.com/v1.0/groups/$($team.GroupId)/drive/root/children"
        Method  = "GET"
        Headers = @{
            Authorization = "Bearer $($adminToken)"
        }
    }
    $filesResponse = Invoke-RestMethod @invokeWebRequestParams
    $allFileInformation = $filesResponse.Value

    # Retrieve the owners of the team
    $invokeWebRequestParams = @{
        Uri     = "https://graph.microsoft.com/v1.0/groups/$($team.GroupId)/owners"
        Method  = "GET"
        Headers = @{
            Authorization = "Bearer $($adminToken)"
        }
    }
    $ownerResponse = Invoke-RestMethod @invokeWebRequestParams
    $owners = $ownerResponse.value

    # Get an owner whose department is the majority
    foreach ($owner in $owners) {
        $user = $allAzureADUsers | Where-Object { $_.UserPrincipalName -eq $owner.userPrincipalName }
        $owner | Add-Member -NotePropertyName "Department" -NotePropertyValue $user.Department
    }
    $groupedOwners = $owners | Group-Object -Property "Department" | Sort-Object -Property Count -Descending
    $ownerRepresentative = $groupedOwners[0].Group[0]

    # Add the owner information to the team
    $team | Add-Member -NotePropertyName "Owner" -NotePropertyValue $ownerRepresentative.displayName -Force
    $team | Add-Member -NotePropertyName "OwnerDepartment" -NotePropertyValue $ownerRepresentative.Department -Force
    $team | Add-Member -NotePropertyName "OwnerJobTitle" -NotePropertyValue $ownerRepresentative.jobTitle -Force

    # Declare default values
    $earliestTime = $now.AddYears(-100)
    $minDaysSinceLastActivity = 99999999

    Write-Information "Retrieving channels in '$($team.DisplayName)'."
    $teamChannels = Get-TeamChannel -GroupId $team.GroupId
    foreach ($channel in $teamChannels) {
        $lastActivityTime = $earliestTime
        $channel | Add-Member -NotePropertyName "Team" -NotePropertyValue $team.DisplayName
        $channel | Add-Member -NotePropertyName "OwnerDepartment" -NotePropertyValue $ownerRepresentative.Department

        $teamFileInformation = $allFileInformation | Where-Object { $_.name -eq $channel.DisplayName }
        $channel | Add-Member -NotePropertyName "TotalFileSize" -NotePropertyValue $teamFileInformation.size
        $channel | Add-Member -NotePropertyName "FileLastModifiedTime" -NotePropertyValue $teamFileInformation.lastModifiedDateTime

        # Update last activity time
        if ($channel.FileLastModifiedTime -gt $lastActivityTime) {
            $lastActivityTime = $channel.FileLastModifiedTime
        }

        # Retrieve chat information
        Write-Information "Retrieving chat information for channel $($channel.DisplayName)"
        $lastChatTime = $earliestTime
        try {
            # Retrieve chat messages
            $invokeWebRequestParams = @{
                Uri     = "https://graph.microsoft.com/beta/teams/$($team.GroupId)/channels/$($channel.Id)/messages"
                Method  = "GET"
                Headers = @{
                    Authorization = "Bearer $($adminToken)"
                }
            }
            $response = Invoke-RestMethod @invokeWebRequestParams
            $chatCount = $response.'@odata.count'
            $channel | Add-Member -NotePropertyName "ChatCount" -NotePropertyValue $chatCount -Force
            $messages = $response.value

            # Get the latest message and update last chat time
            $latestMessage = $messages[0]
            $messageTime = Get-Date -Date $latestMessage.createdDateTime
            if ($messageTime -gt $lastChatTime) {
                $lastChatTime = $messageTime
            }

            # Retrieve the reply of the latest message
            $invokeWebRequestParams = @{
                Uri     = "https://graph.microsoft.com/beta/teams/$($team.GroupId)/channels/$($channel.Id)/messages/$($message.Id)/replies"
                Method  = "GET"
                Headers = @{
                    Authorization = "Bearer $($adminToken)"
                }
            }
            $response = Invoke-RestMethod @invokeWebRequestParams

            # Update last chat time
            if ($response.'@odata.count' -gt 0) {
                $latestReplyTime = Get-Date -Date $response.value[0].createdDateTime
                if ($latestReplyTime -gt $lastChatTime) {
                    $lastChatTime = $latestReplyTime
                }
            }
        }
        catch {
            # 404 means there isn't any chat message
        }
        finally {
            # Update last chat time and last activity time
            if ($lastChatTime -gt $lastActivityTime) {
                $lastActivityTime = $lastChatTime
            }
            if ($lastChatTime -eq $earliestTime) {
                $lastChatTime = ""
            }
            $channel | Add-Member -NotePropertyName "LastChatTime" -NotePropertyValue $lastChatTime -Force
        }

        # Calculate number of days since last activity
        $daysSinceLastActivity = ($now - (Get-Date -Date $lastActivityTime)).Days
        if ($daysSinceLastActivity -le $minDaysSinceLastActivity) {
            $minDaysSinceLastActivity = $daysSinceLastActivity
        }
        $channel | Add-Member -NotePropertyName "DaysSinceLastActivity" -NotePropertyValue $daysSinceLastActivity -Force
        $allChannels += $channel
    }
    $team | Add-Member -NotePropertyName "DaysSinceLastActivity" -NotePropertyValue $minDaysSinceLastActivity -Force
    $team | Add-Member -NotePropertyName "Channels" -NotePropertyValue $teamChannels -Force
    $team | Add-Member -NotePropertyName "ChannelNames" -NotePropertyValue ($teamChannels.DisplayName -join ", ") -Force
    $team | Add-Member -NotePropertyName "NumberOfChannels" -NotePropertyValue $teamChannels.Length -Force

    # Retrieve team users
    Write-Information "Retrieving users in '$($team.DisplayName)'."
    $teamUsers = Get-TeamUser -GroupId $team.GroupId
    foreach ($user in $teamUsers) {
        $user | Add-Member -NotePropertyName "Team" -NotePropertyValue $team.DisplayName
        $allUsers += $user
    }

    # Get external users
    $externalUsers = $teamUsers | Where-Object { $_.User -like "*#EXT#*" }
    $team | Add-Member -NotePropertyName "ExternalUsers" -NotePropertyValue ($externalUsers.User -join ", ") -Force
    $team | Add-Member -NotePropertyName "HasExternalUsers" -NotePropertyValue ($externalUsers.Length -gt 0) -Force
}

# This function calculates the similarity between two strings
function Get-StringSimilarity {
    param(
        $String1,
        $String2
    )
    return Get-Random -Maximum 100000
}

# This function returns similar string pairs within a collection of strings
function Get-SimilarStrings {
    param(
        $Strings,
        $StringName = "String",
        $Top,
        $Threshold
    )
    $string1Name = $StringName + "1"
    $string2Name = $StringName + "2"
    $comparisonResults =[System.Collections.ArrayList]@()

    # Calculate the similarity between each string pair, and append to the result
    for ($i=0; $i -lt $Strings.Length; $i++) {
        for ($j=$i+1; $j -lt $Strings.Length; $j++) {
            $similarity = Get-StringSimilarity -String1 $Strings[$i] -String2 $Strings[$j]
            $comparisonResults.Add([PSCustomObject]@{
                    $string1Name = $Strings[$i]
                    $string2Name = $Strings[$j]
                    Similarity   = $similarity
                }) | Out-Null
        }
    }

    # Return the top string pairs
    if ($Top) {
        return $comparisonResults | Sort-Object -Property Similarity -Descending | Select-Object -First $Top
    }

    # Return string pairs that have a higher similarity than the threshold
    elseif ($Threshold) {
        return $comparisonResults | Where-Object { $_.Similarity -ge $Threshold }
    }
}

$numTopResults = 5
$threshold = 8000
###
# Get similar names
####

# Get similar team names
Write-Information "Calculating name similarity for teams."
$similarTeamNames = Get-SimilarStrings -Strings $allTeams.DisplayName -StringName "TeamName" -Threshold 99990

# Get similar channel names for channels in the same team
$similarChannelNamesWithSameTeam = @()
foreach ($team in $allTeams) {
    Write-Information "Calculating name similarity for channels that are in $($team.DisplayName)."
    $similarNames = Get-SimilarStrings -Strings $team.Channels.DisplayName -StringName "ChannelName" -Threshold 99990
    foreach ($pair in $similarNames) {
        $pair | Add-Member -NotePropertyName "Team" -NotePropertyValue $team.DisplayName
    }
    $similarChannelNamesWithSameTeam += $similarNames
}

# Get similar channel names whose owners are in the same department
$channelsGroupedByOwner = $allChannels | Group-Object -Property "OwnerDepartment"
$similarChannelNamesWithSameOwner = @()
foreach ($owner in $channelsGroupedByOwner) {
    Write-Information "Calculating name similarity for channels whose owners are in $($owner.Name)."
    $similarNames = Get-SimilarStrings -Strings $owner.Group.DisplayName -StringName "ChannelName" -Threshold 99990
    foreach ($pair in $similarNames) {
        $pair | Add-Member -NotePropertyName "OwnerDepartment" -NotePropertyValue $owner.Name
    }
    $similarChannelNamesWithSameOwner += $similarNames
}

# Declare the ranges of the number of channels in a team (both sides inclusive)
$maxNumberOfChannels = 201
$ranges = @(
    @(1, 10),
    @(11, 50),
    @(51, 150),
    @(151, 190),
    @(190, $maxNumberOfChannels)
)

# Get the number in each range
$numberOfChannelsDistribution = @()
foreach ($range in $ranges) {
    $numberOfChannelsDistribution += ($allTeams | Where-Object { $_.NumberOfChannels -ge $range[0] -and $_.NumberOfChannels -le $range[1] }).Length
}

# Generate statistics
$externalUsers = $allUsers.User | Select-Object -Unique | Where-Object { $_ -like "*#EXT#*" }
$teamsWithExternalUsers = ($allTeams | Where-Object { $_.HasExternalUsers }).DisplayName -join ", "

# Create an array to store statistics
$stats = [System.Collections.ArrayList]@()
Update-Statistics -Statistics $stats -Key "numberOfTeams" -Value $allTeams.Length
Update-Statistics -Statistics $stats -Key "numberOfPublicTeams" -Value $numberOfPublicTeams
Update-Statistics -Statistics $stats -Key "numberOfPrivateTeams" -Value $numberOfPrivateTeams
Update-Statistics -Statistics $stats -Key "numberOfExternalUsers" -Value $externalUsers.Length
Update-Statistics -Statistics $stats -Key "teamsWithExternalUsers" -Value $teamsWithExternalUsers
Update-Statistics -Statistics $stats -Key "SimilarTeamNames" -Value $similarTeamNames.Length
Update-Statistics -Statistics $stats -Key "SimilarChannelNamesWithSameTeam" -Value $similarChannelNamesWithSameTeam.Length
Update-Statistics -Statistics $stats -Key "SimilarChannelNamesWithSameOwner" -Value $similarChannelNamesWithSameOwner.Length

$templateFilePath = "$($PSScriptRoot)\TeamsReportTemplate.xlsx"
$reportFilePath = "$($PSScriptRoot)\TeamsReport.xlsx"

#$templateFilePath = "C:\Users\yshen\Desktop\TeamsReportTemplate.xlsx"
#$reportFilePath = "C:\Users\yshen\Desktop\TeamsReport.xlsx"

# Save all the date to an Excel file
Copy-Item -Path $templateFilePath -Destination $reportFilePath
$excelPackage = $allUsers | Export-Excel -Path $reportFilePath -WorksheetName "Users" -PassThru
$excelPackage = $allChannels | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Channels" -PassThru
$excelPackage = $allTeams | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Teams" -PassThru
$excelPackage = $numberOfChannelsDistribution | Export-Excel -ExcelPackage $excelPackage -WorksheetName "NumberOfChannels" -PassThru -StartRow 2 -StartColumn 2
$excelPackage = $similarTeamNames | Export-Excel -ExcelPackage $excelPackage -WorksheetName "SimilarTeamNames" -PassThru
$excelPackage = $similarChannelNamesWithSameTeam | Export-Excel -ExcelPackage $excelPackage -WorksheetName "SimilarChannelNames1" -PassThru
$excelPackage = $similarChannelNamesWithSameOwner | Export-Excel -ExcelPackage $excelPackage -WorksheetName "SimilarChannelNames2" -PassThru
$excelPackage = $stats | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Stats" -PassThru
Close-ExcelPackage -ExcelPackage $excelPackage #-Show
Write-Information "Finished."

# Store the Excel workbook base64 encoded bytes in a file in the instance scope
$excelWorkbookContentsFileName = "$(New-Guid).txt"
$context.SaveInstanceText($excelWorkbookContentsFileName, (Convert-FileToBase64EncodedBytes -Path $reportFilePath))
$context.Outputs.teamsReportContentsFileName = $excelWorkbookContentsFileName