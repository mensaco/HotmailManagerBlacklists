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

$urlseantd = "https://raw.githubusercontent.com/mensaco/HotmailManagerBlacklists/main/sender.emailAddress.name.to.delete"
$urlseactd = "https://raw.githubusercontent.com/mensaco/HotmailManagerBlacklists/main/sender.emailAddress.address.contains.to.delete"
$urlsctd = "https://raw.githubusercontent.com/mensaco/HotmailManagerBlacklists/main/subject.contains.to.delete"

$response = Invoke-WebRequest -Uri $urlseantd -UseDefaultCredentials 
$sendersToDelete = $response.Content.Split("`n")

$response = Invoke-WebRequest -Uri $urlseactd -UseDefaultCredentials 
$sendersContainsToDelete = $response.Content.Split("`n")

$response = Invoke-WebRequest -Uri $urlsctd -UseDefaultCredentials 
$subjectContainsToDelete = $response.Content.Split("`n")

Connect-MgGraph -Scopes $scooes

$response = Invoke-MgGraphRequest -Method GET $url

$idstodelete = @()

$response.value | ForEach-Object {

    $deleting = $false
    
    $deletingInfo = "sender.emailAddress.name:`t[$($_.sender.emailAddress.name)]`nsender.emailAddress.address:`t[$($_.sender.emailAddress.address)]`nsubject:`t[$($_.subject)]`n"
    
    if($sendersToDelete.ToLower().Contains($_.sender.emailAddress.name.ToLower())) {
        $deleting = $true
        $idstodelete += $_.id
    }

    foreach($contains in $sendersContainsToDelete)  {
        if($null -ne $_.from.emailAddress.address -and $_.sender.emailAddress.address.ToLower().Contains($contains.toLower())){
            $deleting = $true
            $idstodelete += $_.id
        }
        if($null -ne $_.from.emailAddress.name -and $_.sender.emailAddress.name.ToLower().Contains($contains.toLower())){
            $deleting = $true
            $idstodelete += $_.id
        }

    }


    foreach ($contains in $subjectContainsToDelete) {
        if($null -ne $_.subject -and $_.subject.ToLower().Contains($contains.toLower())){
            $deleting = $true
            $idstodelete += $_.id
        }
    }

    if($deleting){
        write-host $deletingInfo -ForegroundColor Red
    }
    else {
        write-host $deletingInfo -ForegroundColor Gray
    }

}




$idstodelete | ForEach-Object {
    $messageId = $_
    #$messageId
    # A UPN can also be used as -UserId.
    Move-MgUserMessage  -UserId $userId -MessageId $messageId -BodyParameter $params
}



# $dummy = 0

# Write-Host "OK"