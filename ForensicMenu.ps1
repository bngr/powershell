$systemName = Read-Host "Enter the system name or IP address"
"`n" 
Write-Host "Checking to see if the host is alive. If ping fails, I probably won't work."
ping -n 4 -a $systemName | Out-Default
"`n" 

$workingDirectory = '[insert working directory]'

function Show-Menu
{
    param (
        [string]$Title = $systemName
    )
    Write-Host "=f-a=n-c=y-=-=-=-=-= $Title =-=-=-=-=-a=s-c=i-i="
    Write-Host "All information will be written to $workingDirectory."
    Write-Host "1: Installed Software"
    Write-Host "2: Accounts logged on"
    Write-Host "3: Running Processes"
    Write-Host "4: Startup Commands"
    Write-Host "5: Get C:\Users\"
    Write-Host "6: Chrome Extensions for ALL Users on Machine"
    Write-Host "Q: Quit"
}

"`n" 

do
 {
    Show-Menu –Title $systemName
    $selection = Read-Host "Select Number: "
    switch ($selection)
    {
     '1' {
         Get-WmiObject -ComputerName $systemName -Query “select * from Win32_Product” | Select-Object * | Sort-Object Name | Out-File -FilePath $workingDirectory\$systemName-applications.txt
     } '2' {
         psloggedon \\$systemName |  Out-File -FilePath $workingDirectory\$systemName-accounts.txt
     } '3' {
         Get-WmiObject win32_process -ComputerName $systemName | select processname,@{NAME='CreationDate';EXPRESSION={$_.ConvertToDateTime($_.CreationDate)}},ProcessId,CommandLine | sort CreationDate -desc | format-table –auto -wrap  | Out-File -FilePath $workingDirectory\$systemName-processes.txt
     } '4' {  
         Get-WmiObject -ComputerName $systemName -Query “select * from Win32_StartupCommand” | Select-Object * | Out-File -FilePath $workingdirectory\$systemName-startupcommands.txt
     } '5' {
         Get-ChildItem -path "\\$systemName\C$\Users" | Sort-Object LastWriteTime -descending | select * |  Out-File -FilePath $workingdirectory\$systemName-c_users.txt
     } '6' { 
         $directoryPath = "\\$systemName\c$\Users\" 
         foreach ($file in get-ChildItem $directoryPath*) 
           {
                Write-Host "Currently checking $file on $systemName"
                Get-ChildItem "$directoryPath\$($file.name)\AppData\Local\Google\Chrome\User Data\Default\Extensions\ " -ErrorAction SilentlyContinue | Out-File -FilePath $workingDirectory\$systemName-allUsersChrome.txt
           }
     } 'q' {
         return
        }
    }
    pause
}
until ($selection -eq 'q')