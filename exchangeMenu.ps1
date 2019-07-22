 function senderSearch {
    $Global:senderAddress = Read-Host "Sender address" 
    $backlog = Read-Host "Checking only the last 48 hours. Do you want to go back further? [y/n]"
    if ($backlog -eq 'n'){
        Get-MessageTrace -SenderAddress $senderAddress | Format-Table Received, RecipientAddress, SenderAddress, Subject, Status, FromIP
        }
    elseif ($backlog -eq 'y'){
        Write-Host "Searching the past 7 days." -ForegroundColor Red
        Get-MessageTrace -SenderAddress $senderAddress -StartDate ((Get-Date).AddDays(-7).ToString('MM/dd/yyyy')) -EndDate (Get-Date -Uformat "%m/%d/%Y") | Format-Table Received, RecipientAddress, SenderAddress, Subject, Status, FromIP
        }
    }


function emailPull {
    $targetMailbox = Read-Host "Enter name of user [Firstname Lastname]"
    $newSenderCheck = Read-Host "Keep the same sender of '$senderAddress'? [y/n]"
    if ($newSenderCheck -eq 'y'){
        Search-Mailbox -Identity "$targetMailbox" -SearchQuery '(from:$senderAddress)' -DeleteContent
        }
    elseif ($newSenderCheck -eq 'n'){
        $newSender = Read-Host "Sender address"
        Search-Mailbox -Identity "$targetMailbox" -SearchQuery '(from:$newSender)' -DeleteContent
        }
    $continue = Read-Host "Do you have more messages to pull? [y/n]"
    if ($continue -eq 'y'){
        emailPull
        }
    elseif ($continue -eq 'n'){
        Return
        }
    }


function Main {
    Write-Host "Select a module:"
    Write-Host "1: Sender address search"
    Write-Host "2: Delete messages from inboxes"
    Write-Host "q: Quit"
    $Global:selection = Read-Host "Selection"
    if ($selection -eq '1'){
        senderSearch
        }
    elseif ($selection -eq '2'){
        emailPull
        }
    }


do {
    Main
    }

until ($selection -eq 'q')
