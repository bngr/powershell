param (
    [parameter(Position=0)]
    [String]
    $File=$(throw "Usage: './MassServerLogoff.ps1 -File [filename]'")
)

$Global:serverlist = Get-Content $File

function loop {
    foreach ($Global:server in $serverlist){
        Main
    }
}

function qwServer {
    Write-Host "Current sessions for $server" -ForegroundColor white -BackgroundColor red
    qwinsta /SERVER:$server
}

function qwLogoff {
    $session = Read-Host "Session"
        if ($session -like "*x*"){
            Continue
        }
    logoff $session /SERVER:$server /V
}

function Main {
     qwServer
     qwLogoff
}

loop
##
