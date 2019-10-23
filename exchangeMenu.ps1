 function Format-Color([hashtable] $Colors = @{}, [switch] $SimpleMatch) {
	$lines = ($input | Out-String) -replace "`r", "" -split "`n"
	foreach($line in $lines) {
		$color = ''
		foreach($pattern in $Colors.Keys){
			if(!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
			elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
		}
		if($color) {
			Write-Host -ForegroundColor $color $line
		} else {
			Write-Host $line
		}
	}
}
 

 function senderSearch {
    $Global:senderAddress = Read-Host "Sender address" 
    $backlog = Read-Host "Checking only the last 48 hours. Do you want to go back further? [y/n]"
    if ($backlog -eq 'n'){
        Get-MessageTrace -SenderAddress $senderAddress | Format-Table Received, RecipientAddress, Status, SenderAddress, Subject, FromIP | Format-Color @{'Delivered' = 'Green'; 'FilteredAsSpam' = 'Red'}
        }
    elseif ($backlog -eq 'y'){
        Write-Host "Searching the past 7 days." -ForegroundColor Red
        Get-MessageTrace -SenderAddress $senderAddress -StartDate ((Get-Date).AddDays(-7).ToString('MM/dd/yyyy')) -EndDate (Get-Date -Uformat "%m/%d/%Y") | Format-Table Received, RecipientAddress, Status, SenderAddress, Subject, FromIP | Format-Color @{'Delivered' = 'Green'; 'FilteredAsSpam' = 'Red'}
        }
    }


function ipSearch {
    $ipaddresss = Read-Host "IP address" 
    $backlog = Read-Host "Checking only the last 48 hours. Do you want to go back further? [y/n]"
    if ($backlog -eq 'n'){
        Get-MessageTrace -FromIP $ipaddress | Format-Table Received, RecipientAddress, Status, SenderAddress, Subject, FromIP | Format-Color @{'Delivered' = 'Green'; 'FilteredAsSpam' = 'Red'}
        }
    elseif ($backlog -eq 'y'){
        Write-Host "Searching the past 7 days." -ForegroundColor Red
        Get-MessageTrace -FromIP $senderAddress -StartDate ((Get-Date).AddDays(-7).ToString('MM/dd/yyyy')) -EndDate (Get-Date -Uformat "%m/%d/%Y") | Format-Table Received, RecipientAddress, Status, SenderAddress, Subject,  FromIP | Format-Color @{'Delivered' = 'Green'; 'FilteredAsSpam' = 'Red'}
        }
    }


function emailPull {
    $targetMailbox = Read-Host "Enter name of user [Firstname Lastname]"
    $newSenderCheck = Read-Host "Keep the same sender of '$senderAddress'? [y/n]"
    if ($newSenderCheck -eq 'y'){
        Search-Mailbox -Identity "$targetMailbox" -SearchQuery '(from:$senderAddress)' -DeleteContent
        Write-Host "Message from $sendAddress has been pulled from $targetMailbox" -ForegroundColor Red
        }
    elseif ($newSenderCheck -eq 'n'){
        $newSender = Read-Host "Sender address"
        Search-Mailbox -Identity "$targetMailbox" -SearchQuery '(from:$newSender)' -DeleteContent
        Write-Host "Message from $newSender has been pulled from $targetMailbox" -ForegroundColor Red
        }
    $continue = Read-Host "Do you have more messages to pull? [y/n]"
    if ($continue -eq 'y'){
        emailPull
        }
    elseif ($continue -eq 'n'){
        Return
        }
    }


function blockSender {
    $blockSenderCheck = Read-Host "Keep the same sender of '$senderAddress'? [y/n]"
    if ($blockSenderCheck -eq 'y'){
        Set-HostedContentFilterPolicy -Identity "ULX-Spam-Policy" -BlockedSenders @{Add="$senderAddress"}
        Write-Host "$senderAddress has been blocked" -ForegroundColor Red
        }
    elseif ($blockSenderCheck -eq 'n'){
        $newBlockSender = Read-Host "Sender address to block"
        Set-HostedContentFilterPolicy -Identity "ULX-Spam-Policy" -BlockedSenders @{Add="$newBlockSender"}
        Write-Host "$newBlockSender has been blocked" -ForegroundColor Red
        } 
    }


function checkRules {
    $userMailbox = Read-Host "Enter user to check in form of email address"
    Get-InboxRule -Mailbox $userMailbox | Select Name, Description, ForwardTo, ForwardAsAttachmentTo, RedirectTo | FL
    }


function Main {
    Write-Host "----------------"
    Write-Host "Select a module:"
    Write-Host "----------------"
    Write-Host "1: Sender address search"
    Write-Host "2: Sender IP address search"
    Write-Host "3: Delete messages from inboxes"
    Write-Host "4: Block sender address"
    Write-Host "5: Check applied mailbox rules for user"
    Write-Host "q: Quit"
    $Global:selection = Read-Host "Selection"
    if ($selection -eq '1'){
        senderSearch
        }
    elseif ($selection -eq '2'){
        ipSearch
        }
    elseif ($selection -eq '3'){
        emailPull
        }
    elseif ($selection -eq '4'){
        blockSender
        }
    elseif ($selection -eq '5'){
        checkRules
        }
    }


do {
    Main
    }


until ($selection -eq 'q')
