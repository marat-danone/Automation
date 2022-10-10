<#
$CurrentDate = Get-Date
[int]$CurrentDay = $CurrentDate.Day
$9Hour = Get-Date -Date "9:00"
$18Hour = Get-Date -Date "18:00"
$Holidays = Get-Content "C:\Scripts\Work Folder\HolidaysList.txt"
$CurrentDayStatus = $Holidays[$CurrentDay - 1]

if ($CurrentDate -ge $9Hour -and $CurrentDate -le $18Hour -and $CurrentDayStatus -ne 1) {
#>
$Date = Get-Date -Format "dd.MM.yyyy HH_mm"
$LogFile = "C:\Scripts\Logs\ServiceNow-Jira-Sync\$Date.txt"
Start-Transcript $LogFile

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('http://dn-ruibs-wss1.nead.danet:8089')
[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true

$JiraSnowUsers = Import-Excel 'C:\Scripts\Work Folder\SnowJiraUsers.xlsx' -DataOnly
$ExcludedIssues = Get-Content 'C:\Scripts\Work Folder\ExcludedIssues.txt'

function Send-Notification($subject, $body, $RecipientList, [array]$Attachment) {
    $username = 'notirobo@danone.com'
    $password = Get-Content "C:\Scripts\pwd\notirobo-sa.txt" | ConvertTo-SecureString
    $creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $password
    $smtpserver = 'smtp.office365.com'
    $port = '587'
    #Send-MailMessage -From $username -To $RecipientList -Subject $subject -Body $body -BodyAsHtml -Attachments $Attachment -SmtpServer $smtpserver -UseSsl -Credential $creds -Port $port
    Send-MailMessage -From $username -To $RecipientList -Subject $subject -Body $body -Attachments $Attachment -SmtpServer $smtpserver -UseSsl -Credential $creds -Port $port -BodyAsHtml
}

function Try-Expression($expression, $message, [boolean]$showsuccessmessage = $true) {
    if ($message) {
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) $message"
    }
    try {
        $error.Clear()
        $command = ''
        $command = Invoke-Expression $expression -ErrorAction Continue
    }
    catch {
    }
    if (!$error[0].exception) {
        if ($showsuccessmessage -eq $true) {
            Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Success"
        }
    }
    else {
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) $error[0]"
        $Body = @"
<p style="margin: 0cm; font-size: 15px; font-family: Calibri, sans-serif; line-height: 1;"><strong>The error is:</strong></p>
<p style="line-height: 1;"><em><span style="font-family: Calibri, sans-serif;">$( $error[0].exception.message );.</span></em><strong> </strong></p>
<p style="margin: 0cm; font-size: 15px; font-family: Calibri, sans-serif; line-height: 1;">Please check the logs on server wrumosrui032.</p>
"@
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Send email notification about error..."
        Stop-Transcript
        Send-Notification -subject "An error occured while processing `"ServiceNow-Jira-Sync`" script" -body $Body -RecipientList marat.aliyev@external.danone.com -Attachment $LogFile
        Exit
    }
    return $command
}

$TransitionTable = Import-Excel "C:\Scripts\Work Folder\TransitionTable.xlsx" -DataOnly

function Get-SnowIssue {
    $Params = @{
        Headers = $GlobalHeaders
        Method = "GET"
        Uri = "$ServiceNowInstance/api/now/table/$( $SnowTable )?sysparm_query=number=$( $JiraIssue.Number )&sysparm_display_value=All"
    }

    $SnowIssue = ""
    $SnowIssue = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Getting `"$( $JiraIssue.Number )`" parameters in Snow..."
    $SnowIssue = $SnowIssue.Result
    return $SnowIssue
}

function Set-JiraIssueTransition($TransitionID, $IssueKey) {
    $params = @{
        Body = @{
            "transition" = @{
                "id" = "$TransitionID"
            }
        } | ConvertTo-Json -Depth 100
        uri = "$JiraInstance/issue/$IssueKey/transitions"
        headers = @{ "Content-Type" = "application/json"; "Authorization" = "Basic $Base64" }
        method = "Post"
    }

    Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Setting jira issue transition status as `"$TransitionName`"..."
}

function Set-SnowIssueTransition($TransitionID, $IssueID, $SnowTable) {
    switch ($TransitionID) {
        "3" {
            $Params = @{
                Headers = $GlobalHeaders
                Method = "PATCH"
                Uri = "$ServiceNowInstance/api/now/table/$SnowTable/$IssueID"
                Body = @{
                    active = "false"
                    state = "$TransitionID"
                } | ConvertTo-Json -Depth 100
            }
        }
        Default {
            $Params = @{
                Headers = $GlobalHeaders
                Method = "PATCH"
                Uri = "$ServiceNowInstance/api/now/table/$SnowTable/$IssueID"
                Body = @{
                    state = "$TransitionID"
                } | ConvertTo-Json -Depth 100
            }
        }
    }

    Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Setting snow issue transition status as `"$SnowTransitionName`"..."
}

function Create-JiraStory($SnowIssue, $IssueType, $TransitionColumn) {

    $IssueAssignee = ""
    $IssueAssignee = ($JiraSnowUsers | Where { $_.UserName -eq $( $SnowIssue.assigned_to.display_value ) }).JiraID
    # Create new issue in Jira based on change parameters

    $Summary = "❄" + " " + $SnowIssue.short_description.display_value + " " + "[$( $SnowIssue.number.value )]"

    switch ($IssueType) {
        "CHANGE" {
            $Description = @"
*[Description]*
$( $SnowIssue.description.display_value )

*[Justification]*
$( $SnowIssue.justification.display_value )

*[Snow ticket]*
[$( $SnowIssue.number.value )|https://danone.service-now.com/nav_to.do?uri=task.do?sys_id=$( $SnowIssue.sys_id.value )]
"@
        }
        "RITM" {
            $Description = @"
*[Description]*
$( $SnowIssue.description.display_value )

*[Snow ticket]*
[$( $SnowIssue.number.value )|https://danone.service-now.com/nav_to.do?uri=task.do?sys_id=$( $SnowIssue.sys_id.value )]
"@
        }
    }


    $params = @{
        Body = @{
            "fields" = @{
                "project" = @{
                    "key" = $ProjectKey
                }
                "assignee" = @{
                    "accountId" = $IssueAssignee
                }
                "summary" = $Summary
                "description" = $Description
                "issuetype" = @{
                    "name" = "Story"
                }
                "parent" = @{
                    "key" = $( $EpicList."Server Administration" )
                }
            }
        } | ConvertTo-Json -Depth 100
        uri = "$JiraInstanceApi2/issue"
        headers = @{ "Content-Type" = "application/json"; "Authorization" = "Basic $Base64" }
        method = "Post"
    }

    $NewIssue = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Creating new Jira story with summary `"$Summary`"..."
    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Story $( $NewIssue.key ) created"
    #Start-Sleep -s 5

    $TransitionName = ""
    $TransitionID = ""
    $JiraEntry = ($TransitionTable | Where { $_.$TransitionColumn -match $SnowIssue.state.display_value })
    $TransitionName = $JiraEntry.JiraState
    $TransitionID = $JiraEntry.JiraTransitionID

    Set-JiraIssueTransition -TransitionID $TransitionID -IssueKey $( $NewIssue.key )
    #Start-Sleep -s 5

    # Add snow comments from change if exists
    $Comments = ""
    $Comments = $SnowIssue.comments_and_work_notes.display_value
    if ($Comments) {
        $SnowComments = Convert-SnowComments -Comments $Comments
        Add-SnowCommentsToJira -IssueKey $( $NewIssue.key ) -Comments $SnowComments.Comment
    }

    $Attachment = Find-SnowAttachment -SnowTable "change_request" -SnowIssue $( $SnowIssue.sys_id.value )

    if ($Attachment) {
        # Download attachment from snow, then upload it to Jira issue
        foreach ($File in $Attachment) {
            $FileName = $( $File.file_name ) -replace ":", "_"
            $FilePath = "C:\Scripts\Work Folder\JiraSnow Attachments\$FileName"
            $Params = @{
                OutFile = $FilePath
                Headers = $GlobalHeaders
                Method = "GET"
                Uri = "$( $File.download_link )"
            }
            Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Downloading attachment with name `"$FileName`"..."

            Add-JiraIssueAttachment -FileName $FileName -FilePath $FilePath -JiraIssue $( $NewIssue.key )
            #Start-Sleep -s 5
        }
    }

    # Get all subtasks in snow issue

    switch ($IssueType) {
        "CHANGE" {
            $Params = @{
                Headers = $GlobalHeaders
                Method = "GET"
                Uri = "$ServiceNowInstance/api/now/table/change_task?sysparm_query=change_request=$( $SnowIssue.sys_id.value )&sysparm_display_value=All"
            }
            $SubTaskTransitionColumn = "CtaskState"
            $SubTaskSnowTable = "change_task"
        }
        "RITM" {
            $Params = @{
                Headers = $GlobalHeaders
                Method = "GET"
                Uri = "$ServiceNowInstance/api/now/table/sc_task?sysparm_query=request_item.sys_id=$( $SnowIssue.sys_id.value )&sysparm_display_value=All"
            }
            $SubTaskTransitionColumn = "SctaskState"
            $SubTaskSnowTable = "sc_task"
        }
    }


    $SubTasks = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Getting all subtasks related to $IssueType `"$( $SnowIssue.number.value )`"..."
    $SubTasks = $SubTasks.result
    $Assignee = ($JiraSnowUsers | Where { $_.UserName -eq ($SubTasks | Where { $_.short_description.display_value -eq "Final full backup of server (for archive)" -or $_.short_description.display_value -eq "ITCC RUCIS SA - RUCIS - Configuration check" }).assigned_to.display_value }).JiraID
    if (!$Assignee) {
        $Assignee = ($JiraSnowUsers | Where { $_.UserName -eq $SnowIssue.assigned_to.display_value }).JiraID
    }

    # Add subtasks to Jira

    foreach ($Task in $SubTasks) {
        Create-JiraSubTask -ParentIssue $( $NewIssue.key ) -SnowIssue $Task -TransitionColumn $SubTaskTransitionColumn -SnowTable $SubTaskSnowTable -Assignee $Assignee
        #Start-Sleep -s 5
    }

    if ($IssueType -eq "RITM") {
        $params = @{
            Body = @{
                "fields" = @{
                    "assignee" = @{
                        "accountId" = $Assignee
                    }
                }
            } | ConvertTo-Json -Depth 100
            uri = "$JiraInstance/issue/$( $NewIssue.key )"
            headers = @{ "Content-Type" = "application/json"; "Authorization" = "Basic $Base64" }
            method = "Put"
        }

        Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Setting jira assignee for `"$( $NewIssue.key )`" as `"$( $MainSctask.assigned_to.display_value )`"..."
    }
}

function Create-JiraSubTask($ParentIssue, $SnowIssue, $TransitionColumn, $SnowTable, $Assignee) {
    $IssueAssignee = ($JiraSnowUsers | Where { $_.UserName -eq $( $SnowIssue.assigned_to.display_value ) }).JiraID
    if (!$IssueAssignee) {
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) User `"$( $SnowIssue.assigned_to.display_value )`" not found in JiraSnowUsers table..."
        $IssueAssignee = $Assignee
    }

    $TransitionName = ""
    $TransitionID = ""
    $JiraEntry = ($TransitionTable | Where { $_.$TransitionColumn -match $SnowIssue.state.display_value })
    $TransitionName = $JiraEntry.JiraState
    $TransitionID = $JiraEntry.JiraTransitionID

    $IssueSummary = "❄" + " " + $SnowIssue.short_description.value + " " + "[$( $SnowIssue.number.value )]"
    $IssueDescription = @"
*[Description]*
$( $SnowIssue.description.value )

*[Snow ticket]*
[$( $SnowIssue.number.value )|https://danone.service-now.com/nav_to.do?uri=task.do?sys_id=$( $SnowIssue.sys_id.value )]
"@

    $params = @{
        Body = @{
            "fields" = @{
                "project" = @{
                    "key" = $ProjectKey
                }
                "assignee" = @{
                    "accountId" = $IssueAssignee
                }
                "summary" = $IssueSummary
                "description" = $IssueDescription
                "issuetype" = @{
                    "name" = "Sub-task"
                }
                "parent" = @{
                    "key" = $ParentIssue
                }
            }
        } | ConvertTo-Json -Depth 100
        uri = "$JiraInstanceApi2/issue"
        headers = @{ "Content-Type" = "application/json"; "Authorization" = "Basic $Base64" }
        method = "Post"
    }

    $NewSubtask = ""
    $NewSubtask = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Creating new subtask with name `"$IssueSummary`" under $ParentIssue..."
    #Start-Sleep -s 1

    Set-JiraIssueTransition -TransitionID $TransitionID -IssueKey $( $NewSubtask.key )

    # Get comments
    $Comments = ""
    $Comments = $SnowIssue.comments_and_work_notes.display_value
    $SnowComments = Convert-SnowComments -Comments $Comments
    if ($SnowComments) {
        Add-SnowCommentsToJira -IssueKey $( $NewSubtask.key ) -Comments $SnowComments.Comment
    }

    $SnowAttachments = Find-SnowAttachment -SnowTable $SnowTable -SnowIssue $( $SnowIssue.sys_id.value )

    foreach ($File in $SnowAttachments) {
        if ($File.file_name -notin $JiraAttachments.FileName) {
            $FileName = $File.file_name -replace ":", "_"
            $FilePath = "C:\Scripts\Work Folder\JiraSnow Attachments\$FileName"
            $Params = @{
                OutFile = $FilePath
                Headers = $GlobalHeaders
                Method = "GET"
                Uri = "$( $File.download_link )"
            }
            Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Downloading attachment with name `"$( $File.file_name )`"..."

            Add-JiraIssueAttachment -FileName $( $File.file_name ) -FilePath $FilePath -JiraIssue $( $NewSubtask.key )
            #Start-Sleep -s 5
        }
    }
}

function Convert-SnowComments($Comments) {
    $MultiString = ""
    $Comments = $Comments.Split("`n")
    foreach ($str in $Comments) {
        if (!$str) {
            $MultiString += @"
--LineBreak--`r`n
"@
        }
        else {
            $MultiString += @"
$str`r`n
"@
        }
    }

    $CommentsArray = ($MultiString -split "--LineBreak--").Trim() | Where { $_ }
    $CommentsArray = $CommentsArray -replace "`r", ""
    # Reverse array
    [array]::Reverse($CommentsArray)

    $SnowComments = @()
    foreach ($str in $CommentsArray) {
        if ($str -notmatch [Regex]::Escape("[code]<i>Approved")) {
            $SnowComments += New-Object -TypeName PSObject -Property ([ordered]@{
                Name = $str.Split("`n")[0].Trim()
                Comment = $str
            })
        }
    }
    return $SnowComments
}

function Convert-JiraComments($Comments) {
    $JiraComments = @()
    foreach ($str in $Comments) {
        $JiraComments += New-Object -TypeName PSObject -Property ([ordered]@{
            Name = $str.Comments.Split("`n")[0].Trim()
            Comment = $str.Comments
            AuthorID = $str.Author.AccountId
            AuthorName = $str.Author.DisplayName
            Self = $str.self
        })
    }
    return $JiraComments
}

function Add-SnowCommentsToJira($IssueKey, $Comments) {
    # Add comments sequentally to Jira issue
    foreach ($Comment in $Comments) {
        if ($Comment -notmatch [Regex]::Escape("[code]<i>Approved")) {
            $Comment = $Comment -replace [Regex]::Escape($Comment.Split("`n")[0]), ("*" + $Comment.Split("`n")[0] + "*")
            $params = @{
                Body = @{
                    "update" = @{
                        "comment" = @(@{
                            "add" = @{
                                "body" = $Comment
                            }
                        })
                    }
                } | ConvertTo-Json -Depth 100
                uri = "$JiraInstanceApi2/issue/$IssueKey"
                headers = @{ "Content-Type" = "application/json"; "Authorization" = "Basic $Base64" }
                method = "Put"
            }

            Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Adding comment to Jira issue $IssueKey..."
            #Start-Sleep -s 5
        }
    }
}

function Find-SnowAttachment($SnowTable, $SnowIssue) {
    $Params = @{
        Headers = $GlobalHeaders
        Method = "GET"
        Uri = "$ServiceNowInstance/api/now/attachment?sysparm_query=table_name=$SnowTable&table_sys_id=$SnowIssue"
    }

    $Attachment = ""
    $Attachment = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Getting attachments in Snow..."
    $Attachment = $Attachment.Result
    return $Attachment
}

function Add-JiraIssueAttachment($FileName, $FilePath, $JiraIssue) {
    $FilePath = "C:\Scripts\Work Folder\JiraSnow Attachments\$FileName"

    $FileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $FileEnc = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString($FileBytes)
    $Boundary = [System.Guid]::NewGuid().ToString()


    $params = @{
        body = @'
--{0}
Content-Disposition: form-data; name="file"; filename="{1}"
Content-Type: application/octet-stream

{2}
--{0}--
'@ -f $Boundary, $FileName, $FileEnc
        uri = "$JiraInstance/issue/$JiraIssue/attachments"
        headers = @{ "X-Atlassian-Token" = "nocheck"; "Authorization" = "Basic $Base64" }
        method = "Post"
    }

    Try-Expression -expression 'Invoke-RestMethod @params -ContentType "multipart/form-data; boundary=`"$Boundary`"" -UseBasicParsing' -message "Adding attachment `"$FilePath`" to jira issue `"$JiraIssue`"..."
}

function Add-SnowIssueAttachment($FileName, $FilePath, $TableName, $SnowIssue) {
    $Params = @{
        Infile = "$FilePath"
        Headers = $GlobalHeaders
        Method = "Post"
        Uri = "$ServiceNowInstance/api/now/attachment/file?table_name=$TableName&table_sys_id=$( $SnowIssue.sys_id.value )&file_name=$FileName"
        ContentType = "multipart/form-data"
    }

    Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Adding attachment `"$FilePath`" to snow issue `"$( $SnowIssue.number.value )`"..."
}

function Add-JiraCommentToSnow($Table, $IssueID, $WorkNotes) {
    $Params = @{
        Headers = $GlobalHeaders
        Method = "PATCH"
        Uri = "$ServiceNowInstance/api/now/table/$Table/$IssueID"
        Body = @{
            work_notes = "$WorkNotes"
        } | ConvertTo-Json -Depth 100
    }

    Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Adding new comment in Service Now..."
}

################## Snow auth ##################################
# Set the credentials
$User = "service.powershell"
$pass = Get-Content "C:\Scripts\pwd\default\service.powershell.txt" | ConvertTo-SecureString
$pass = [System.Net.NetworkCredential]::new('', $pass).password

# Set headers
$GlobalHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$GlobalHeaders.Add('Accept', 'application/json')
$GlobalHeaders.Add('Content-Type', 'application/json')

# Build & set authentication header
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $User, $Pass)))
$GlobalHeaders.Add('Authorization', ('Basic {0}' -f $base64AuthInfo))

# Load variable with ServiceNow instance URL
$ServiceNowInstance = 'https://danone.service-now.com'


################## Jira auth ##################################
$JiraApiToken = Get-Content C:\Scripts\pwd\default\JiraApiToken.txt | ConvertTo-SecureString
$JiraApiToken = [System.Net.NetworkCredential]::new('', $JiraApiToken).password

$JiraInstance = "https://speedboats-gti.atlassian.net/rest/api/3"
$JiraInstanceApi2 = "https://speedboats-gti.atlassian.net/rest/api/2"
$Pair = "notirobo@danone.com:$JiraApiToken"
$Bytes = [System.Text.Encoding]::ASCII.GetBytes($Pair)
$Base64 = [System.Convert]::ToBase64String($Bytes)
$BasicAuthValue = "Basic $Base64"
$Headers = @{ Authorization = $BasicAuthValue }

$Projects = Invoke-RestMethod -Uri "$JiraInstance/project" -Method Get -Headers $Headers

$ProjectName = "Managed Cloud Platform CIS"
$Params = @{
    uri = "$JiraInstance/project/search?query=$ProjectName"
    headers = @{ "Authorization" = "Basic $Base64" }
    method = "Get"
}

$Project = Try-Expression -expression 'Invoke-RestMethod @Params -UseBasicParsing' -message "Getting project `"$ProjectName`" info..."
$Project = $Project.values
$ProjectKey = $Project.key
$ProjectID = $Project.id

# Get issues

# Additional condition to add below: and status!=Cancelled and and status!=Done

function Get-JiraIssues($StartAt) {
    $Params = @{
        uri = "$JiraInstance/search?jql=project=$ProjectKey and issuetype!=Epic and status!=Done and status!=Cancelled&fields=key,summary,assignee,status,subtasks,comment,attachment&startAt=$StartAt&maxResults=100"
        headers = @{ "Authorization" = "Basic $Base64" }
        method = "Get"
    }

    $Issues = Try-Expression -expression 'Invoke-RestMethod @Params -UseBasicParsing' -message "Getting all issues in project with statuses not Done and not Cancelled starting at $StartAt..."
    return $Issues
}

$StartAt = 0
$Issues = Get-JiraIssues -StartAt $StartAt
$TotalIssues = @()
$TotalCount = $Issues.Total
$TotalIssues += $Issues.issues

While ($TotalIssues.Count -lt $TotalCount) {
    $StartAt = $TotalIssues.Count
    $Issues = Get-JiraIssues -StartAt $StartAt
    $TotalIssues += $Issues.issues
}

$SnowIssuesInJira = $TotalIssues | Where { $_.fields.summary -like "❄*" }
$SnowIssuesInJiraSummary = $SnowIssuesInJira.fields.summary

$JiraIssuesList = @()
foreach ($Item in $SnowIssuesInJira) {
    $IssueNumber = $Item.fields.summary.Split(" ")[($Item.fields.summary.Split(" ").count - 1)] -replace '[^\w\d]', ''
    switch ($IssueNumber) {
        { $_ -like "CHG*" } {
            $TransitionColumn = "ChangeState"
        }
        { $_ -like "CTASK*" } {
            $TransitionColumn = "CtaskState"
        }
        { $_ -like "SCTASK*" } {
            $TransitionColumn = "SctaskState"
        }
        { $_ -like "RITM*" } {
            $TransitionColumn = "RitmState"
        }
    }
    $IssueState = ""
    $IssueState = ($TransitionTable | Where { $_.JiraState -eq $Item.fields.status.name }).$TransitionColumn
    if ($IssueState | Select-String ",") {
        $IssueState = $IssueState.Split(",").Trim()
    }
    $CommentsList = @()
    foreach ($Cmt in $Item.fields.comment.comments) {
        $CommentsList += New-Object -TypeName PSObject -Property ([ordered]@{
            Comments = ($Cmt.body.content.content.text | Out-String).Trim() -replace "`r", ""
            Author = $Cmt.author
            Self = $Cmt.self
        })
    }
    $JiraIssuesList += New-Object -TypeName PSObject -Property ([ordered]@{
        Key = $Item.key
        Summary = $Item.fields.summary
        Number = $Item.fields.summary.Split(" ")[($Item.fields.summary.Split(" ").count - 1)] -replace '[^\w\d]', ''
        Assignee = $Item.fields.assignee.displayname
        State = $IssueState
        Subtasks = $Item.fields.subtasks
        Comments = $CommentsList
        Attachments = $Item.fields.attachment
    })
}

###################### Snow ##########################

# Get list of Epics

$Params = @{
    uri = "$JiraInstance/search?jql=project=$ProjectKey and issuetype=Epic&maxResults=100&fields=summary,key"
    headers = @{ "Authorization" = "Basic $Base64" }
    method = "Get"
}

$Epics = Try-Expression -expression 'Invoke-RestMethod @Params -UseBasicParsing' -message "Getting all Epics in project..."
$Epics = $Epics.issues

$EpicList = @()
foreach ($Epic in $Epics) {
    $EpicList += [ordered]@{
        $Epic.fields.summary = $Epic.key
    }
}

# Searching for existing jira issues in Service Now

Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Processing existing snow issues in Jira..."

foreach ($JiraIssue in $JiraIssuesList) {
    switch ($JiraIssue.Number) {
        { $_ -like "CHG*" } {
            $SnowTable = "change_request"
            $TransitionColumn = "ChangeState"
        }
        { $_ -like "CTASK*" } {
            $SnowTable = "change_task"
            $TransitionColumn = "CtaskState"
        }
        { $_ -like "SCTASK*" } {
            $SnowTable = "sc_task"
            $TransitionColumn = "SctaskState"
        }
        { $_ -like "RITM*" } {
            $SnowTable = "sc_req_item"
            $TransitionColumn = "RitmState"
        }
    }

    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Processing issue `"$( $JiraIssue.Number )`" with summary `"$( $JiraIssue.Summary )`" and snow ticket `"$( $JiraIssue.number )`"..."

    # Get issue parameters in snow

    $SnowIssue = Get-SnowIssue
    $JiraIssueStateString = ""
    if ($JiraIssue.State.Count -gt 1) {
        foreach ($Issue in $JiraIssue.State) {
            $JiraIssueStateString += $Issue + ","
        }
        $JiraIssueState = $JiraIssueStateString.Substring(0, $JiraIssueStateString.Length - 1)
    }
    else {
        $JiraIssueState = $JiraIssue.State
    }

    $SnowIssueState = ($( $TransitionTable.$TransitionColumn ) | Where { $_ -match $( $SnowIssue.state.display_value ) })

    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Jira issue state: $JiraIssueState"
    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Snow issue state: $SnowIssueState"

    if ($SnowIssueState) {
        $SnowIssueStateLevel = [Array]::IndexOf($( $TransitionTable.$TransitionColumn ), $SnowIssueState)
        $JiraIssueStateLevel = [Array]::IndexOf($( $TransitionTable.$TransitionColumn ), $JiraIssueState)

        if ($SnowIssueStateLevel -gt $JiraIssueStateLevel) {
            $StateToChange = ($TransitionTable | Where { $_.$TransitionColumn -match $JiraIssueState }).$TransitionColumn
            if (!$StateToChange) {
                Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) State to change is empty. Manual investigation needed!"
                $Body = @"
<p style="margin: 0cm; font-size: 15px; font-family: Calibri, sans-serif; line-height: 1;"><strong>State to change is empty. Check the log file in attachment.</strong></p>
"@
                Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Send email notification about error..."
                Stop-Transcript
                Send-Notification -subject "Jira-ServiceNow Sync" -body $Body -RecipientList marat.aliyev@external.danone.com -Attachment $LogFile
                Exit
            }
            Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Cannot change issue state in Service Now to `"$StateToChange`". Rolling back state is not allowed!"
            $TransitionName = ""
            $TransitionID = ""
            $JiraEntry = ($TransitionTable | Where { $_.$TransitionColumn -match $SnowIssue.state.display_value })
            $TransitionName = $JiraEntry.JiraState
            $TransitionID = $JiraEntry.JiraTransitionID
            Set-JiraIssueTransition -TransitionID $TransitionID -IssueKey $JiraIssue.key
        }
        elseif ($SnowIssueStateLevel -lt $JiraIssueStateLevel) {
            $TransitionName = ""
            $TransitionID = ""
            $JiraEntry = ($TransitionTable | Where { $_.$TransitionColumn -match $JiraIssue.state })
            $SnowTransitionName = $JiraEntry.$TransitionColumn
            $SnowIssueTransitionID = $JiraEntry.(($TransitionColumn -replace "State", "") + "TransitionID")
            Set-SnowIssueTransition -TransitionID $SnowIssueTransitionID -IssueID $SnowIssue.sys_id.value -SnowTable $SnowTable
        }
    }
    else {
        $JiraEntry = ($TransitionTable | Where { $_.$TransitionColumn -match $SnowIssue.state.display_value })
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Cannot find the equivalent of Jira issue state to Snow in transition table. Issue state will be set as `"$( $JiraEntry.JiraState )`"."
        $TransitionName = ""
        $TransitionID = ""
        $JiraEntry = ($TransitionTable | Where { $_.$TransitionColumn -match $SnowIssue.state.display_value })
        $TransitionName = $JiraEntry.JiraState
        $TransitionID = $JiraEntry.JiraTransitionID
        Set-JiraIssueTransition -TransitionID $TransitionID -IssueKey $JiraIssue.key
    }

    if ($JiraIssue.Number -like "RITM*") {
        $Params = @{
            Headers = $GlobalHeaders
            Method = "GET"
            Uri = "$ServiceNowInstance/api/now/table/sc_task?sysparm_query=request_item.number=$( $JiraIssue.Number )"
        }

        $Sctasks = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Getting all STASKS related to the same RITM `"$( $JiraIssue.Number )`"..."
        $Sctasks = $Sctasks.result
        $SnowAssignee = ($JiraSnowUsers | Where { $_.SnowID -eq ($Sctasks | Where { $_.short_description -eq "ITCC RUCIS SA - RUCIS - Configuration check" -or $_.short_description -eq "Final full backup of server (for archive)" -or $_.short_description -eq "ITCC SA - Restore a file/DB/DB object" }).assigned_to.value }).UserName
    }
    else {
        $SnowAssignee = $SnowIssue.assigned_to.display_value
    }
    if ($SnowAssignee -eq "Service Powershell (GTI ITCC CIS - Servers and Infra)") {
        $SnowAssignee = "JiraRobot"
    }
    $JiraAssignee = $JiraIssue.Assignee -replace [regex]::Escape("."), " "
    $JiraAssigneeReverse = $JiraAssignee.Split(" ")[1] + " " + $JiraAssignee.Split(" ")[0]

    if ($SnowAssignee -and $SnowAssignee -ne $JiraAssignee -and $SnowAssignee -ne $JiraAssigneeReverse) {
        $AccountID = ($JiraSnowUsers | Where { $_.UserName -eq $SnowAssignee }).JiraID
        if ($AccountID) {
            $params = @{
                Body = @{
                    "fields" = @{
                        "assignee" = @{
                            "accountId" = ($JiraSnowUsers | Where { $_.UserName -eq $SnowAssignee }).JiraID
                        }
                    }
                } | ConvertTo-Json -Depth 100
                uri = "$JiraInstance/issue/$( $JiraIssue.key )"
                headers = @{ "Content-Type" = "application/json"; "Authorization" = "Basic $Base64" }
                method = "Put"
            }

            Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Change issue assignee in Jira to `"$SnowAssignee`"..."
        }
        else {
            Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Cannot find the assignee `"$SnowAssignee`" in JiraSnowUsers table..."
        }
    }

    ## Sync comments
    if ($JiraIssue.Number -notin $ExcludedIssues) {
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Searching for new comments in Snow and Jira..."
        $Comments = ""
        $Comments = $SnowIssue.comments_and_work_notes.display_value
        $JiraComments = Convert-JiraComments -Comments $JiraIssue.comments

        $SnowComments = Convert-SnowComments -Comments $Comments
        $JiraNewComments = $JiraComments | Where { $_.name -notin $SnowComments.Name }

        if ($JiraNewComments) {
            foreach ($NewComment in $JiraNewComments) {
                $JiraCommentAuthor = $NewComment.AuthorID
                $JiraCommentAuthor = ($JiraSnowUsers | Where { $_.JiraID -eq $JiraCommentAuthor }).UserName
                #$NewComment = $NewComment -replace [regex]::Escape("\!"), "!"
                $SnowComment = @"
[$JiraCommentAuthor]
$( $NewComment.Comment )
"@
                Add-JiraCommentToSnow -Table $SnowTable -IssueID $SnowIssue.sys_id.value -WorkNotes $SnowComment
                #Start-Sleep -s 5

                # Remove jira comment from Jira

                $params = @{
                    uri = "$( $NewComment.self )"
                    headers = @{ "Content-Type" = "application/json"; "Authorization" = "Basic $Base64" }
                    method = "Delete"
                }

                Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Removing comment `"$( $NewComment.self )`" from Jira..."
            }

            # Getting snow issue comments again
            $SnowIssue = Get-SnowIssue
        }

        $Comments = ""
        $Comments = $SnowIssue.comments_and_work_notes.display_value
        $SnowComments = Convert-SnowComments -Comments $Comments
        $SnowNewComments = $SnowComments | Where { $_.Name -notin $JiraComments.Name }

        if ($SnowNewComments) {
            Add-SnowCommentsToJira -IssueKey $( $JiraIssue.key ) -Comments $SnowNewComments.Comment
        }
        #Start-Sleep -s 5
    }
    ##

    # Sync attachments
    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Searching for new attachments in Snow and Jira..."

    $SnowAttachments = Find-SnowAttachment -SnowTable $SnowTable -SnowIssue $( $SnowIssue.sys_id.value )
    $JiraAttachments = $JiraIssue.attachments

    foreach ($File in $SnowAttachments) {
        $JiraAttachmentsFileNames = $JiraAttachments.FileName -replace ":", "_"
        $FileName = $File.file_name -replace ":", "_"
        if ($FileName -notin $JiraAttachmentsFileNames) {
            $FilePath = "C:\Scripts\Work Folder\JiraSnow Attachments\$FileName"
            $Params = @{
                OutFile = $FilePath
                Headers = $GlobalHeaders
                Method = "GET"
                Uri = "$( $File.download_link )"
            }
            Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing -ContentType "application/json;charset=utf-8"' -message "Downloading attachment with name `"$FileName`"..."

            Add-JiraIssueAttachment -FileName $FileName -FilePath $FilePath -JiraIssue $( $JiraIssue.key )
            #Start-Sleep -s 5
        }
    }

    foreach ($File in $JiraAttachments) {
        $SnowAttachmentsFileNames = $SnowAttachments.file_name -replace ":", "_"
        $FileName = $File.FileName -replace ":", "_"
        if ($FileName -notin $SnowAttachmentsFileNames) {
            $FilePath = "C:\Scripts\Work Folder\JiraSnow Attachments\$FileName"
            $Params = @{
                OutFile = $FilePath
                Method = "GET"
                Uri = "$( $File.Content )"
                headers = @{ "ContentType" = "image/png"; "Authorization" = "Basic $Base64" }
            }

            Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Downloading attachment with name `"$FileName`"..."

            Add-SnowIssueAttachment -FileName $FileName -FilePath $FilePath -TableName $SnowTable -SnowIssue $SnowIssue
            #Start-Sleep -s 5
        }
    }

    # Add new subtasks and comments if exists

    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Searching for new subtasks in Snow..."

    if ($SnowTable -eq "change_request" -or $SnowTable -eq "sc_req_item") {
        switch ($SnowTable) {
            "change_request" {
                $Params = @{
                    Headers = $GlobalHeaders
                    Method = "GET"
                    Uri = "$ServiceNowInstance/api/now/table/change_task?sysparm_query=change_request.number=$( $JiraIssue.Number )&sysparm_display_value=All"
                }
                $SubTaskSnowTable = "change_task"
                $TransitionColumn = "CtaskState"
            }
            "sc_req_item" {
                $Params = @{
                    Headers = $GlobalHeaders
                    Method = "GET"
                    Uri = "$ServiceNowInstance/api/now/table/sc_task?sysparm_query=request_item.number=$( $JiraIssue.Number )&sysparm_display_value=All"
                }
                $SubTaskSnowTable = "sc_task"
                $TransitionColumn = "SctaskState"
            }
        }

        $SubTasks = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Getting all subtasks related to snow issue `"$( $JiraIssue.Number )`"..."
        $SubTasks = $SubTasks.result

        switch ($SnowTable) {
            "change_request" {
                $Assignee = ($JiraSnowUsers | Where { $_.UserName -eq $SnowIssue.assigned_to.display_value }).JiraID
            }
            "sc_req_item" {
                $Assignee = ($JiraSnowUsers | Where { $_.UserName -eq ($SubTasks | Where { $_.short_description.display_value -eq "Final full backup of server (for archive)" -or $_.short_description.display_value -eq "ITCC RUCIS SA - RUCIS - Configuration check" }).assigned_to.display_value }).JiraID
            }
        }

        $JiraSubTaskList = @()
        foreach ($String in $JiraIssue.SubTasks.fields.summary) {
            $JiraSubTaskList += $String.Split(" ")[($String.Split(" ").count - 1)] -replace '[^\w\d]', ''
        }

        foreach ($SubTask in $SubTasks) {
            if ($SubTask.number.value -notin $JiraSubTaskList) {
                Create-JiraSubTask -ParentIssue $JiraIssue.key -SnowIssue $SubTask -TransitionColumn $TransitionColumn -SnowTable $SubTaskSnowTable -Assignee $Assignee
                #Start-Sleep -s 5
            }
        }
    }

    #Start-Sleep -s 1
}

# Searching for open CHANGE or RITM in Service Now assigned to someone

#$IssueType = "Story"
# Set the credentials
$User = "service.powershell"
$pass = Get-Content "C:\Scripts\pwd\default\service.powershell.txt" | ConvertTo-SecureString
$pass = [System.Net.NetworkCredential]::new('', $pass).password

# Set headers
$GlobalHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$GlobalHeaders.Add('Accept', 'application/json')
$GlobalHeaders.Add('Content-Type', 'application/json')

# Build & set authentication header
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $User, $Pass)))
$GlobalHeaders.Add('Authorization', ('Basic {0}' -f $base64AuthInfo))

# Load variable with ServiceNow instance URL
$ServiceNowInstance = 'https://danone.service-now.com'

[string]$MonthAgo = Get-Date (Get-Date).AddDays(-30) -Format "yyyy-MM-dd"

Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Processing new change,ctask,ritm,sctask..."

$Params = @{
    Headers = $GlobalHeaders
    Method = "GET"
    Uri = "$ServiceNowInstance/api/now/table/change_request?sysparm_query=assignment_group=dc1031a81b89249081d3c84b1d4bcb63^active=true^stateNOT IN-5,-4,3,4,0^opened_at>javascript:gs.dateGenerate(`'$MonthAgo`','00:00:00')&sysparm_display_value=All"
}

$Changes = ""
$Changes = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Getting changes opened for the last 30 days assigned to group `"GTI - Managed Cloud Platform CIS`"..."
$Changes = $Changes.Result

if (!$Changes) {
    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Nothing found..."
}
else {
    $ChangesListString = ""
    foreach ($ch in $Changes) {
        $ChangesListString += $ch.number.value + ","
    }
    $ChangesListString = $ChangesListString.Substring(0, $ChangesListString.Length - 1)
    Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Found CHANGE's: $ChangesListString..."
}

foreach ($Change in $Changes) {
    if ($Change.number.value -notin $JiraIssuesList.number) {
        Create-JiraStory -SnowIssue $Change -IssueType "CHANGE" -TransitionColumn "ChangeState"
    }
    else {
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) CHANGE `"$( $Change.number.value )`" is already existing in Jira..."
    }
}

# Searching for open RITM where templates are: REQ - ITCC SI - SA: New Virtual Server - ITCC CIS, REQ - ITCC SI - SA: Server decommission - ITCC CIS

$RitmTemplatesList = Import-Excel "C:\Scripts\Work Folder\RitmTemplatesList.xlsx" -DataOnly

foreach ($Item in $RitmTemplatesList) {
    $Params = @{
        Headers = $GlobalHeaders
        Method = "GET"
        Uri = "$ServiceNowInstance/api/now/table/sc_req_item?sysparm_query=cat_item=$( $Item.Id )^active=true^state!=3^opened_at>javascript:gs.dateGenerate(`'$MonthAgo`','00:00:00')&sysparm_display_value=All"
    }

    $RITMList = ""
    $RITMList = Try-Expression -expression 'Invoke-RestMethod @params -UseBasicParsing' -message "Searching for open RITM with ci name `"$( $Item.Name )`"..."
    $RITMList = $RITMList.result

    if (!$RITMList) {
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Nothing found..."
    }
    else {
        $RITMListString = ""
        foreach ($r in $RITMList) {
            $RITMListString += $r.number.value + ","
        }
        $RITMListString = $RITMListString.Substring(0, $RITMListString.Length - 1)
        Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) Found RITM's: $RITMListString..."
    }

    foreach ($RITM in $RITMList) {
        if ($RITM.number.value -notin $JiraIssuesList.number) {
            Create-JiraStory -SnowIssue $RITM -IssueType "RITM" -TransitionColumn "RitmState"
        }
        else {
            Write-Host "$( Get-Date -Format "[dd.MM.yyyy HH:mm:ss]" ) RITM `"$( $RITM.number.value )`" is already existing in Jira..."
        }
    }
}

Stop-Transcript
#}