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
    "Get-SavedCredentials",
    "Connect-AzureADAdminAccount",
    "ConvertTo-Array"
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



$microsoftGraphCredentialsObject = Get-SavedCredentials -CredentialsName $microsoftGraphCredentialsName -ApplicationCredentials
if ($null -eq $microsoftGraphCredentialsObject) {
    Write-Error "Failed to retrieve the Microsoft Graph credentials '$($microsoftGraphCredentialsName)'. Please check the name and try again."
    exit
}
$applicationID = $microsoftGraphCredentialsObject.applicationId
$clientSecret = $microsoftGraphCredentialsObject.clientSecret
$domain = $microsoftGraphCredentialsObject.domain

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

$credentials = [PSCredential]::new(
    $office365Username,
    ($office365Password | ConvertTo-SecureString -AsPlainText -Force)
)

# Connect to platforms
Write-Information "Connecting to platforms."
Connect-MicrosoftTeams -Credential $credentials
Connect-AzureADAdminAccount -Username $office365Username -Password ($office365Password | ConvertTo-SecureString -AsPlainText -Force)
$adminToken = Get-MicrosoftGraphAuthenticationToken -Username $office365Username -Password $office365Password -ApplicationId $applicationID -ClientSecret $clientSecret

Write-Information "Retrieving teams."
$allTeams = Get-Team

Write-Information "Retrieving Azure AD users."
$allAzureADUsers = Get-AzureADUser -All:$true

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

    # Retrieve more information for each channel
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
        Write-Information "Retrieving chat information for channel '$($channel.DisplayName)'"
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
        $String2,
        $Threshold = 0.7
    )

    # WARNING:
    # This function only works if there are no similar words within the strings,
    # and the order of words doesn't matter

    # Split the strings into words
    $string1Words = $String1.Split(" ") | ForEach-Object { $_.ToLower() }
    $string2Words = $String2.Split(" ") | ForEach-Object { $_.ToLower() }

    # Count the number of matches
    $matchCount = 0
    foreach ($word1 in $string1Words) {
        foreach ($word2 in $string2Words) {
            if (Get-WordSimilarity -Word1 $word1 -Word2 $word2 -Threshold $Threshold) {
                $matchCount++
                break
            }
        }
    }

    # Return the similarity of two strings (a number between 0 and 1)
    return ($matchCount * $matchCount) / ($string1Words.Length * $string2Words.Length)
}

# This function calculates the similarity between two words
function Get-WordSimilarity {
    param(
        $Word1,
        $Word2,
        $Threshold = 0.7
    )

    # Get the normalized edit distance between two words
    $similarity = Get-LevenshteinDistance -String1 $Word1 -String2 $Word2 -Normalize

    # Decide the result based on the similarity and threshold
    if ($similarity -gt $Threshold) {
        return $true
    }
    return $false
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

function Get-LevenshteinDistance {
    <#
        .SYNOPSIS
            Get the Levenshtein distance between two strings.
        .DESCRIPTION
            The Levenshtein Distance is a way of quantifying how dissimilar two strings (e.g., words) are to one another by counting the minimum number of operations required to transform one string into the other.
        .EXAMPLE
            Get-LevenshteinDistance 'kitten' 'sitting'
        .LINK
            http://en.wikibooks.org/wiki/Algorithm_Implementation/Strings/Levenshtein_distance#C.23
            http://en.wikipedia.org/wiki/Edit_distance
            https://communary.wordpress.com/
            https://github.com/gravejester/Communary.PASM
        .NOTES
            Author: Øyvind Kallstad
            Date: 07.11.2014
            Version: 1.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$String1,

        [Parameter(Position = 1)]
        [string]$String2,

        # A normalized output will fall in the range 0 (perfect match) to 1 (no match).
        [Parameter()]
        [switch]$NormalizeOutput
    )

    $d = New-Object 'Int[,]' ($String1.Length + 1), ($String2.Length + 1)

    try {
        for ($i = 0; $i -le $d.GetUpperBound(0); $i++) {
            $d[$i, 0] = $i
        }

        for ($i = 0; $i -le $d.GetUpperBound(1); $i++) {
            $d[0, $i] = $i
        }

        for ($i = 1; $i -le $d.GetUpperBound(0); $i++) {
            for ($j = 1; $j -le $d.GetUpperBound(1); $j++) {
                $cost = [Convert]::ToInt32((-not($String1[$i-1] -eq $String2[$j-1])))
                $min1 = $d[($i-1), $j] + 1
                $min2 = $d[$i, ($j-1)] + 1
                $min3 = $d[($i-1), ($j-1)] + $cost
                $d[$i, $j] = [Math]::Min([Math]::Min($min1, $min2), $min3)
            }
        }

        $distance = ($d[$d.GetUpperBound(0), $d.GetUpperBound(1)])

        if ($NormalizeOutput) {
            return (1 - ($distance) / ([Math]::Max($String1.Length, $String2.Length)))
        }

        else {
            return $distance
        }
    }

    catch {
        Write-Warning $_.Exception.Message
    }
}

###
# Get similar names
####

# Get similar team names
Write-Information "Calculating name similarity for teams."
$similarTeamNames = Get-SimilarStrings -Strings (ConvertTo-Array $allTeams.DisplayName) -StringName "TeamName" -Threshold 0.4

# Get similar channel names for channels in the same team
$similarChannelNamesWithSameTeam = @()
foreach ($team in $allTeams) {
    Write-Information "Calculating name similarity for channels that are in '$($team.DisplayName)'."
    $similarNames = Get-SimilarStrings -Strings (ConvertTo-Array $team.Channels.DisplayName) -StringName "ChannelName" -Threshold 0.4
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
    $similarNames = Get-SimilarStrings -Strings (ConvertTo-Array $owner.Group.DisplayName) -StringName "ChannelName" -Threshold 0.4
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

# Declare the paths of the report template and the report file
$templateFilePath = "$($PSScriptRoot)\TeamsReportTemplate.xlsx"
$reportFilePath = "$($PSScriptRoot)\TeamsReport.xlsx"

#$templateFilePath = "C:\Users\yshen\Desktop\TeamsReportTemplate.xlsx"
#$reportFilePath = "C:\Users\yshen\Desktop\TeamsReport.xlsx"

# Save all the data to an Excel file
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