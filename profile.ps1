try {
    $script:Config = Import-Clixml -Path "$PSScriptRoot\config.xml"
}
catch {
    $script:Config = [ordered]@{
        Credentials  = [ordered]@{
            Admin     = [ordered]@{ }
            O365      = [ordered]@{ }
            NonDomain = [ordered]@{ }
            Standard  = [ordered]@{ }
        }
        Repositories = [ordered]@{ }
    }
    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
}
New-Variable -Name ProfileCredential -Value $script:Config.Credentials -Scope Global
Clear-Host
$(Get-Process -Id $PID).Refresh()

function Prompt {

    $Admin = ''
    $Identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object -TypeName System.Security.Principal.WindowsPrincipal($Identity)
    if ($Principal.IsInRole('Administrators')) {
        $Admin = '[ADMIN]'
    }
    Write-Host $Admin -ForegroundColor Red -NoNewLine
    Write-Host "[${env:USERDOMAIN}]" -ForegroundColor Magenta -NoNewline
    Write-Host "[${env:USERNAME}]" -ForegroundColor Green -NoNewline
    Write-Host "$(Get-Location)>" -ForegroundColor White -NoNewline
    return " "
}