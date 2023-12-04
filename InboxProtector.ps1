#Import-Module Microsoft.Graph
#Import-Module Microsoft.Graph.Users.Actions

Set-Location $PSScriptRoot

$userId = $env:HOTMAIL_USER_ID
$inboxId = $env:HOTMAIL_INBOX_ID
$junkMailId = $env:HOTMAIL_JUNK_ID

$url = "https://graph.microsoft.com/v1.0/me/messages?`$filter=(isread eq false and (parentFolderId eq '$inboxId' or parentFolderId eq '$junkMailId'))&`$select=id,isread,sender,from,subject"

$params = @{
    destinationId = "deleteditems"
}

$scooes = "Mail.ReadBasic,Mail.ReadWrite"

$urlseantd = "https://raw.githubusercontent.com/mensaco/HotmailManagerBlacklists/main/sender.emailAddress.name.to.delete?v=$([System.DateTime]::Now.Millisecond)"
$urlseactd = "https://raw.githubusercontent.com/mensaco/HotmailManagerBlacklists/main/sender.emailAddress.address.contains.to.delete?v=$([System.DateTime]::Now.Millisecond)"
$urlsctd = "https://raw.githubusercontent.com/mensaco/HotmailManagerBlacklists/main/subject.contains.to.delete?v=$([System.DateTime]::Now.Millisecond)"

$response = Invoke-WebRequest -Uri $urlseantd -UseDefaultCredentials 
$sendersToDelete = $response.Content.Split("`n", [StringSplitOptions]::RemoveEmptyEntries)

$response = Invoke-WebRequest -Uri $urlseactd -UseDefaultCredentials 
$sendersContainsToDelete = $response.Content.Split("`n", [StringSplitOptions]::RemoveEmptyEntries)

$response = Invoke-WebRequest -Uri $urlsctd -UseDefaultCredentials 
$subjectContainsToDelete = $response.Content.Split("`n", [StringSplitOptions]::RemoveEmptyEntries)

Connect-MgGraph -Scopes $scooes

$response = Invoke-MgGraphRequest -Method GET $url

$idstodelete = New-Object "System.Collections.Generic.HashSet[string]"

$response.value | ForEach-Object {

    $deleting = $false
    
    $deletingInfo = "sender.emailAddress.name:`t[$($_.sender.emailAddress.name)]`nsender.emailAddress.address:`t[$($_.sender.emailAddress.address)]`nsubject:`t[$($_.subject)]`n"
    
    if($sendersToDelete.ToLower().Contains($_.sender.emailAddress.name.ToLower())) {
        $deleting = $true
        $idstodelete.Add($_.id)
    }

    foreach($contains in $sendersContainsToDelete)  {
        if($null -ne $_.from.emailAddress.address -and $_.sender.emailAddress.address.ToLower().Contains($contains.toLower())){
            $deleting = $true
            $idstodelete.Add($_.id)
        }
        if($null -ne $_.from.emailAddress.name -and $_.sender.emailAddress.name.ToLower().Contains($contains.toLower())){
            $deleting = $true
            $idstodelete.Add($_.id)
        }

    }


    foreach ($contains in $subjectContainsToDelete) {
        if($null -ne $_.subject -and $_.subject.ToLower().Contains($contains.toLower())){
            $deleting = $true
            $idstodelete.Add($_.id)
        }
    }

    if($deleting){
        write-host $deletingInfo -ForegroundColor Red
    }
    else {
        write-host $deletingInfo -ForegroundColor Gray
    }

}


foreach ($messageId in $idstodelete) {
    Move-MgUserMessage  -UserId $userId -MessageId $messageId -BodyParameter $params
}
