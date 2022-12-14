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

function Start-PowerShellConfiguration {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential
    )

    if (!($Credential)) {
        if (!($ProfileCredential.Standard.$($env:USERDOMAIN))) {
            $Credential = $Host.UI.PromptForCredential('PowerShell Configuration', 'Enter password', "${env:USERDOMAIN}\${env:USERNAME}", '')
            Add-ProfileCredential -Type Standard -NewName $env:USERDOMAIN -NewCredential $Credential
        }
        else {
            $Credential = $ProfileCredential.Standard.$($env:USERDOMAIN)
        }
    }
    Update-ProfileRegistryConfiguration -UpdateType All
    Update-PackageManagement
    Update-ProfileRepository
    Update-ProfileModules
}

function Update-ProfileRegistryConfiguration {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateSet('FontSize', 'PackageManagement', 'Repository', 'Module', 'All')]
        [System.String]
        $UpdateType,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $ProcessID
    )

    $windowVisible = if ($(Get-Process -Id $PID | Select-Object -ExpandProperty MainWindowHandle) -eq 0) { $false } else { $true }

    if (!($UpdateType)) {
        if ($windowVisible) {
            $choiceList = New-Object -TypeName System.Collections.ArrayList
            $choiceCount = 0
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
            ForEach ($key in @('FontSize', 'PackageManagement', 'Repository', 'Module', 'All')) {
                $choiceCount++
                [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
            }
            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
            $choice = $Host.UI.PromptForChoice('Select Update Type', 'Update Type:', $choices, 0)
            if ($choice -eq 0) {
                return
            }
            else {
                $UpdateType = $choiceList[$choice].HelpMessage
            }
        }
        else {
            $UpdateType = 'All'
        }
    }

    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object -TypeName System.Security.Principal.WindowsPrincipal($identity)
    if (!($principal.IsInRole('Administrators'))) {
        if ($windowVisible) { Write-Host 'Updating Registry...' }
        Start-Process PowerShell -Verb RunAs -ArgumentList $MyInvocation.MyCommand.Name, $UpdateType -WindowStyle Hidden -Wait
        if ($windowVisible) { Write-Host 'Registry Update: ' -NoNewline; Write-Host 'Complete' -ForegroundColor Green }
    }
    else {
        $fontPath = "REGISTRY::HKCU\Console"
        $basePath = "REGISTRY::HKLM\SOFTWARE"
        $modulePath = "REGISTRY::HKLM\SOFTWARE\ITAdmin\WindowsPowerShell\ProfileManagement\Modules"
        $repositoryPath = "REGISTRY::HKLM\SOFTWARE\ITAdmin\WindowsPowerShell\ProfileManagement\Repositories"
        $pkgProviderPath = "REGISTRY::HKLM\SOFTWARE\ITAdmin\WindowsPowerShell\ProfileManagement\PackageProviders"
        if ($ProcessID) { if (Get-Process -Id $ProcessID -ErrorAction SilentlyContinue) { Stop-Process -Id $ProcessID -Force } }
        if ($UpdateType -eq 'FontSize' -or $UpdateType -eq 'All') {
            if ($windowVisible) { Write-Progress -Activity $MyInvocation.MyCommand.Name -Status "UpdateType: $UpdateType" -Id 1 -CurrentOperation 'Setting Console FontSize' -PercentComplete $((0 / 1) * 100) }
            $fontPath = "REGISTRY::HKCU\Console"
            if (Get-ItemProperty -Path $(Join-Path -Path $fontPath -ChildPath "%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe") -Name FontSize -ErrorAction SilentlyContinue) {
                if ($(Get-ItemProperty -Path $(Join-Path -Path $fontPath -ChildPath "%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe") -Name FontSize).FontSize -ne 917504) {
                    Push-Location -Path $fontPath -StackName PowerShell
                    Set-ItemProperty -Path $(Join-Path -Path $fontPath -ChildPath "%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe") -Name FontSize -Value 917504 -Type DWord
                    Pop-Location -StackName PowerShell
                }
            }
            else {
                Push-Location -Path $fontPath -StackName PowerShell
                $null = New-ItemProperty -Path $(Join-Path -Path $fontPath -ChildPath "%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe") -Name FontSize -Value 917504 -Type DWord
                Pop-Location -StackName PowerShell
            }
            if ($windowVisible) { Write-Progress -Activity $MyInvocation.MyCommand.Name -Id 1 -Completed }
        }

        if ($UpdateType -eq 'Module' -or $UpdateType -eq 'All') {
            if ($windowVisible) { Write-Progress -Activity $MyInvocation.MyCommand.Name -Status "UpdateType: $UpdateType" -Id 1 -CurrentOperation 'Performing Setup' -PercentComplete $((0 / 2) * 100) }
            'ITAdmin', 'WindowsPowerShell', 'ProfileManagement', 'Modules' | ForEach-Object {
                if (!(Test-Path -Path $(Join-Path -Path $basePath -ChildPath $_) -ErrorAction SilentlyContinue)) {
                    $null = New-Item -Path $(Join-Path -Path $basePath -ChildPath $_) -ItemType Directory -Force
                }
                $basePath = Join-Path -Path $basePath -ChildPath $_
            }
            $basePath = Split-Path -Path $(Split-Path -Path $(Split-Path -Path $(Split-Path -Path $basePath -Parent) -Parent) -Parent) -Parent
            'Program Files (x86)', 'Program Files', 'Users', 'WINDOWS' | ForEach-Object {
                if (!(Test-Path -Path $(Join-Path -Path $modulePath -ChildPath $_))) {
                    $null = New-Item -Path $(Join-Path -Path $modulePath -ChildPath $_) -ItemType Directory -Force
                }
            }
            if ($windowVisible) { Write-Progress -Activity $MyInvocation.MyCommand.Name -Status "UpdateType: $UpdateType" -Id 1 -CurrentOperation 'Processing Modules' -PercentComplete $((1 / 2) * 100) }
            $modCount = 0
            $totalMods = Get-Module -ListAvailable
            Get-Module -ListAvailable | Sort-Object -Property Name, Version | ForEach-Object {
                $mod = $_
                if ($windowVisible) { Write-Progress -Activity 'Processing Modules' -Status "Completed: ${modCount} / $($totalMods.Count)" -Id 2 -ParentId 1 -CurrentOperation "Processing $($mod.Name)v$($mod.Version)" -PercentComplete $(($modCount / $totalMods.Count) * 100) }
                $installPath = Join-Path -Path $modulePath -ChildPath $($mod.ModuleBase.Split('\') | Select-Object -First 2 | Select-Object -Last 1)
                if (!(Test-Path -Path $(Join-Path -Path $installPath -ChildPath $mod.Name) -ErrorAction SilentlyContinue)) {
                    $null = New-Item -Path $(Join-Path -Path $installPath -ChildPath $mod.Name) -ItemType Directory -Force
                    $installPath = Join-Path -Path $installPath -ChildPath $mod.Name
                    if (!(Test-Path -Path $(Join-Path -Path $installPath -ChildPath -$mod.Version) -ErrorAction SilentlyContinue)) {
                        $null = New-Item -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -ItemType Directory -Force
                    }
                    'Guid', 'HelpInfoUri', 'ModuleBase', 'Author', 'CompanyName', 'Copyright' | ForEach-Object {
                        if (!(Get-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_ -ErrorAction SilentlyContinue)) {
                            $null = New-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_ -Value $mod.$($_)
                        }
                    }
                }
                else {
                    $installPath = Join-Path -Path $installPath -ChildPath $mod.Name
                    if (!(Test-Path -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -ErrorAction SilentlyContinue)) {
                        $null = New-Item -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -ItemType Directory -Force
                        'Guid', 'HelpInfoUri', 'ModuleBase', 'Author', 'CompanyName', 'Copyright' | ForEach-Object {
                            if (!(Get-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_ -ErrorAction SilentlyContinue)) {
                                $null = New-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_ -Value $mod.$($_)
                            }
                        }
                    }
                    else {
                        'Guid', 'HelpInfoUri', 'ModuleBase', 'Author', 'CompanyName', 'Copyright' | ForEach-Object {
                            if (!(Get-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_ -ErrorAction SilentlyContinue)) {
                                $null = New-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_ -Value $mod.$($_)
                            }
                            else {
                                if ($(Get-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_).$($_) -ne $mod.$($_)) {
                                    Set-ItemProperty -Path $(Join-Path -Path $installPath -ChildPath $mod.Version) -Name $_ -Value $mod.$($_)
                                }
                            }
                        }
                    }
                }
                $modCount++
            }
            if ($windowVisible) {
                Write-Progress -Activity 'Processing Modules' -Id 2 -Completed
                Write-Progress -Activity $MyInvocation.MyCommand.Name -Id 1 -Completed
            }
        }

        if ($UpdateType -eq 'Repository' -or $UpdateType -eq 'All') {
            if ($windowVisible) { Write-Progress -Activity $MyInvocation.MyCommand.Name -Status "UpdateType: $UpdateType" -Id 1 -CurrentOperation 'Processing Repositories' -PercentComplete $((0 / 2) * 100) }
            'ITAdmin', 'WindowsPowerShell', 'ProfileManagement', 'Repositories' | ForEach-Object {
                if (!(Test-Path -Path $(Join-Path -Path $basePath -ChildPath $_) -ErrorAction SilentlyContinue)) {
                    $null = New-Item -Path $(Join-Path -Path $PWD.Path -ChildPath $_) -ItemType Directory -Force
                }
                $basePath = Join-Path -Path $basePath -ChildPath $_
            }
            $basePath = Split-Path -Path $(Split-Path -Path $(Split-Path -Path $(Split-Path -Path $basePath -Parent) -Parent) -Parent) -Parent
            $repoCount = 0
            $repoTotal = @(Get-PSRepository).Count
            Get-PSRepository | Sort-Object -Property Name | ForEach-Object {
                $repo = $_
                if ($windowVisible) { Write-Progress -Activity 'Processing Repositories' -Status 'Processing PSRepositories' -Id 2 -ParentId 1 -CurrentOperation "Processing $($repo.Name) PSRepository" -PercentComplete $(($repoCount / $repoTotal) * 100) }
                if (!(Test-Path -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -ErrorAction SilentlyContinue)) {
                    $null = New-Item -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -ItemType Directory -Force
                    'SourceLocation', 'PublishLocation', 'ScriptSourceLocation', 'ScriptPublishLocation', 'InstallationPolicy' | ForEach-Object {
                        if (!(Get-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -Name $_ -ErrorAction SilentlyContinue)) {
                            $null = New-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -Name $_ -Value $repo.$($_)
                        }
                    }
                }
                else {
                    'SourceLocation', 'PublishLocation', 'ScriptSourceLocation', 'ScriptPublishLocation', 'InstallationPolicy' | ForEach-Object {
                        if (!(Get-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -Name $_ -ErrorAction SilentlyContinue)) {
                            $null = New-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -Name $_ -Value $repo.$($_)
                        }
                        else {
                            if ($(Get-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -Name $_).$($_) -ne $repo.$($_)) {
                                Set-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $repo.Name) -Name $_ -Value $repo.$($_)
                            }
                        }
                    }
                }
                $repoCount++
            }
            if ($windowVisible) {
                Write-Progress -Activity 'Processing Repositories' -Id 2 -Completed
                Write-Progress -Activity $MyInvocation.MyCommand.Name -Status "UpdateType: $UpdateType" -Id 1 -CurrentOperation 'Processing Resource Repositories' -PercentComplete $((1 / 2) * 100)
            }
            try {
                $resourceRepoCount = 0
                $resourceRepoTotal = @(Get-PSResourceRepository).Count
                Get-PSResourceRepository | Sort-Object -Property Name | ForEach-Object {
                    $resourceRepoCount++
                    $resourceRepo = $_
                    if ($windowVisible) { Write-Progress -Activity 'Processing Resource Repositories' -Status 'Processing PSResourceRepositories' -Id 2 -ParentId 1 -CurrentOperation "Processing $($resourceRepo.Name) PSResourceRepository" -PercentComplete $(($resourceRepoCount / $resourceRepoTotal) * 100) }
                    if (!(Test-Path -Path $(Join-Path -Path $repositoryPath -ChildPath $resourceRepo.Name) -ErrorAction SilentlyContinue)) {
                        $null = New-Item -Path $(Join-Path -Path $repositoryPath -ChildPath $resourceRepo.Name) -ItemType Directory -Force
                        'Url', 'Trusted' | ForEach-Object {
                            if (!(Get-ItemProperty -Path $(Join-Path $repositoryPath -ChildPath $resourceRepo.Name) -Name $_ -ErrorAction SilentlyContinue)) {
                                $null = New-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $resourceRepo.Name) -Name $_ -Value $resourceRepo.$($_)
                            }
                        }
                    }
                    else {
                        'Url', 'Trusted' | ForEach-Object {
                            if (!(Get-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $resourceRepo.Name) -Name $_ -ErrorAction SilentlyContinue)) {
                                $null = New-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $resourceRepo.Name) -Name $_ -Value $resourceRepo.$($_)
                            }
                            else {
                                if ($(Get-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $resourceRepo.Name) -Name $_).$($_) -ne $resourceRepo.$($_)) {
                                    Set-ItemProperty -Path $(Join-Path -Path $repositoryPath -ChildPath $resourceRepo.Name) -Name $_ -Value $resourceRepo.$($_)
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            if ($windowVisible) {
                Write-Progress -Activity 'Processing Repositories' -Id 2 -Completed
                Write-Progress -Activity $MyInvocation.MyCommand.Name -Id 1 -Completed
            }
        }

        if ($UpdateType -eq 'PackageManagement' -or $UpdateType -eq 'All') {
            if ($windowVisible) { Write-Progress -Activity $MyInvocation.MyCommand.Name -Status "UpdateType: $UpdateType" -Id 1 -CurrentOperation 'Setting Package Providers' -PercentComplete $((0 / 4) * 100) }
            'ITAdmin', 'WindowsPowerShell', 'ProfileManagement', 'PackageProviders' | ForEach-Object {
                if (!(Test-Path -Path $(Join-Path -Path $basePath -ChildPath $_) -ErrorAction SilentlyContinue)) {
                    $null = New-Item -Path $(Join-Path -Path $basePath -ChildPath $_) -ItemType Directory -Force
                }
                $basePath = Join-Path -Path $basePath -ChildPath $_
            }
            $basePath = Split-Path -Path $(Split-Path -Path $(Split-Path -Path $(Split-Path -Path $basePath -Parent) -Parent) -Parent) -Parent
            $pkgProviders = Get-PackageProvider | Sort-Object -Property Name
            Get-ChildItem -Path $pkgProviderPath | Where-Object { $_.PSChildName -notin $pkgProviders.Name } | Remove-Item -Recurse -Force
            $completeCount = 0
            ForEach ($pkg in $pkgProviders) {
                if ($windowVisible) { Write-Progress -Activity 'Setting Package Providers' -Status "UpdateType: $UpdateType" -Id 2 -ParentId 1 -CurrentOperation "Processing $($pkg.Name)" -PercentComplete $(($completeCount / $pkgProviders.Count) * 100) }
                if (!(Test-Path -Path $(Join-Path -Path $pkgProviderPath -ChildPath $pkg.Name) -ErrorAction SilentlyContinue)) {
                    $null = New-Item -Path $(Join-Path -Path $pkgProviderPath -ChildPath $pkg.Name) -ItemType Directory -Force
                    $null = New-ItemProperty -Path $(Join-Path -Path $pkgProviderPath -ChildPath $pkg.Name) -Name Version -Value $pkg.Version
                }
                else {
                    if (!(Get-ItemProperty -Path $(Join-Path -Path $pkgProviderPath -ChildPath $pkg.Name) -Name Version -ErrorAction SilentlyContinue)) {
                        $null = New-ItemProperty -Path $(Join-Path -Path $pkgProviderPath -ChildPath $pkg.Name) -Name Version -Value $pkg.Version
                    }
                    else {
                        if ($(Get-ItemProperty -Path $(Join-Path -Path $pkgProviderPath -ChildPath $pkg.Name) -Name Version).Version -ne $pkg.Version) {
                            Set-ItemProperty -Path $(Join-Path -Path $pkgProviderPath -ChildPath $pkg.Name) -Name Version -Value $pkg.Version
                        }
                    }
                }
                $completeCount++
            }
            if ($windowVisible) {
                Write-Progress -Activity 'Setting Package Providers' -Id 2 -Completed
                Write-Progress -Activity $MyInvocation.MyCommand.Name -Id 1 -Completed
            }
        }
    }
}

function Update-PackageManagement {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Repositories.Keys })]
        [System.String]
        $Repository,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $ProcessID
    )

    $windowVisible = if ($(Get-Process -Id $PID | Select-Object -ExpandProperty MainWindowHandle) -eq 0) { $false } else { $true }

    if (!($Repository)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Repositories.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        Switch ($choices.Count) {
            1 {
                Write-Host 'Run ' -NoNewline
                Write-Host 'Add-ProfileRepositoryConfig' -ForegroundColor Yellow -NoNewline
                Write-Host ' to add repository for Package Management'
                return
            }
            2 {
                $Repository = $choiceList[1].HelpMessage
            }
            Default {
                $choice = $Host.UI.PromptForChoice('Select PS Repository', 'PS Repositories:', $choices, 0)
                if ($choice -eq 0) {
                    return
                }
                else {
                    $Repository = $choiceList[$choice].HelpMessage
                }
            }
        }
    }

    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object -TypeName System.Security.Principal.WindowsPrincipal($identity)
    if (!($principal.IsInRole('Administrators'))) {
        if ($windowVisible) { Write-Host 'Updating Package Management...' }
        Start-Process PowerShell -Verb RunAs -ArgumentList $MyInvocation.MyCommand.Name, $Repository -WindowStyle Hidden -Wait
        if ($windowVisible) { Write-Host 'Package Management Update: ' -NoNewline; Write-Host 'Complete' -ForegroundColor Green }
    }
    else {
        if ($ProcessID) { if (Get-Process -Id $ProcessID -ErrorAction SilentlyContinue) { Stop-Process -Id $ProcessID -Force } }
        $repo = $script:Config.Repositories.$Repository
        $regPath = "REGISTRY::HKLM\SOFTWARE\ITAdmin\WindowsPowerShell\ProfileManagement"
        $repoPath = Join-Path -Path $regPath -ChildPath Repositories
        if ($repo.PSRepository -notin $(Get-ChildItem -Path $repoPath).PSChildName) {
            Update-ProfileRepository -Repository $Repository
        }
        $modulePath = Join-Path -Path $regPath -ChildPath Modules
        $programFilesPath = Join-Path -Path $modulePath -ChildPath 'Program Files'
        'PackageManagement', 'PowerShellGet' | ForEach-Object {
            $path = Join-Path -Path $programFilesPath -ChildPath $_
            if ($_ -eq 'PackageManagement') {
                if ([version]'1.4.7' -notin $(Get-ChildItem -Path $path).PSChildName) {
                    try {
                        Save-Module -Name $_ -Repository $repo.PSRepository -RequiredVersion 1.4.7 -Path 'C:\Program Files\WindowsPowerShell\Modules' -Force
                    }
                    catch { }
                }
            }
            if ($_ -eq 'PowerShellGet') {
                if ([version]'2.2.5' -notin $(Get-ChildItem -Path $path).PSChildName) {
                    try {
                        Save-Module -Name $_ -Repository $repo.PSRepository -RequiredVersion 2.2.5 -Path 'C:\Program Files\WindowsPowerShell\Modules' -Force
                    }
                    catch { }
                }
                if ([version]'3.0.12' -notin $(Get-ChildItem -Path $path).PSChildName) {
                    try {
                        Save-Module -Name $_ -Repository $repo.PSRepository -RequiredVersion 3.0.12 -Path 'C:\Program Files\WindowsPowerShell\Modules' -Force
                    }
                    catch { }
                }
            }
        }
        Update-ProfileRegistryConfiguration -UpdateType All
    }
}

function Update-ProfileRepository {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Repositories.Keys })]
        [System.String]
        $Repository,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $ProcessID
    )

    $windowVisible = if ($(Get-Process -Id $PID | Select-Object -ExpandProperty MainWindowHandle) -eq 0) { $false } else { $true }

    if (!($Repository)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Repositories.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        Switch ($choices.Count) {
            1 {
                Write-Host 'Run ' -NoNewline
                Write-Host 'Add-ProfileRepositoryConfig' -ForegroundColor Yellow -NoNewline
                Write-Host ' to add repository for Package Management'
                return
            }
            2 {
                $Repository = $choiceList[1].HelpMessage
            }
            Default {
                $choice = $Host.UI.PromptForChoice('Select PS Repository', 'PS Repositories:', $choices, 0)
                if ($choice -eq 0) {
                    return
                }
                else {
                    $Repository = $choiceList[$choice].HelpMessage
                }
            }
        }
    }

    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object -TypeName System.Security.Principal.WindowsPrincipal($identity)
    if (!($principal.IsInRole('Administrators'))) {
        if ($windowVisible) { Write-Host 'Updating Repository...' }
        Start-Process PowerShell -Verb RunAs -ArgumentList $MyInvocation.MyCommand.Name, $Repository -WindowStyle Hidden -Wait
        if ($windowVisible) { Write-Host 'Repository Update: ' -NoNewline; Write-Host 'Complete' -ForegroundColor Green }
    }
    else {
        if ($ProcessID) { if (Get-Process -Id $ProcessID -ErrorAction SilentlyContinue) { Stop-Process -Id $ProcessID -Force } }
        $repo = $script:Config.Repositories.$Repository
        $regPath = "REGISTRY::HKLM\SOFTWARE\ITAdmin\WindowsPowerShell\ProfileManagement"
        $repoPath = Join-Path -Path $regPath -ChildPath Repositories
        if (!(Test-Path -Path $(Join-Path -Path $repoPath -ChildPath $repo.PSRepository) -ErrorAction SilentlyContinue)) {
            $repoParams = @{
                Name                  = $repo.PSRepository
                SourceLocation        = $repo.SourceLocation
                PublishLocation       = $repo.PublishLocation
                ScriptSourceLocation  = $repo.ScriptSourceLocation
                ScriptPublishLocation = $repo.ScriptPublishLocation
                InstallationPolicy    = $repo.InstallationPolicy
            }
            Register-PSRepository @repoParams
        }
        else {
            $repoParams = New-Object -TypeName System.Collections.Hashtable
            'SourceLocation', 'PublishLocation', 'ScriptSourceLocation', 'ScriptPublishLocation', 'InstallationPolicy' | ForEach-Object {
                if (!(Get-ItemProperty -Path $(Join-Path -Path $repoPath -ChildPath $repo.PSRepository) -Name $_ -ErrorAction SilentlyContinue)) {
                    $repoParams.Add($_, $repo.$($_))
                }
                else {
                    if ($(Get-ItemProperty -Path $(Join-Path -Path $repoPath -ChildPath $repo.PSRepository) -Name $_).$($_) -ne $repo.$($_)) {
                        $repoParams.Add($_, $repo.$($_))
                    }
                }
            }
            if ($repoParams.Count -ne 0) {
                $repoParams.Add('Name', $repo.PSRepository)
                Set-PSRepository @repoParams
            }
        }
        if (!(Test-Path -Path $(Join-Path -Path $repoPath -ChildPath $repo.PSResourceRepository) -ErrorAction SilentlyContinue)) {
            try {
                $repoParams = @(@{
                        Name    = $repo.PSResourceRepository
                        Url     = $repo.Url
                        Trusted = $repo.Trusted
                    })
                Register-PSResourceRepository -Repositories $repoParams
            }
            catch { }
        }
        else {
            $repoParams = New-Object -TypeName System.Collections.Hashtable
            'Url', 'Trusted' | ForEach-Object {
                if (!(Get-ItemProperty -Path $(Join-Path -Path $repoPath -ChildPath $repo.PSResourceRepository) -Name $_ -ErrorAction SilentlyContinue)) {
                    $repoParams.Add($_, $repo.$($_))
                }
                else {
                    if ($(Get-ItemProperty -Path $(Join-Path -Path $repoPath -ChildPath $repo.PSResourceRepository) -Name $_).$($_) -ne $repo.$($_)) {
                        $repoParams.Add($_, $repo.$($_))
                    }
                }
            }
            if ($repoParams.Count -ne 0) {
                try {
                    Set-PSResourceRepository -Repositories @($repoParams)
                }
                catch { }
            }
        }
        Update-ProfileRegistryConfiguration -UpdateType All
    }
}

function Update-ProfileModules {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Repositories.Keys })]
        [System.String]
        $Repository
    )

    if (!($Repository)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Repositories.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        Switch ($choices.Count) {
            1 {
                Write-Host 'Run ' -NoNewline
                Write-Host 'Add-ProfileRepositoryConfig' -ForegroundColor Yellow -NoNewline
                Write-Host ' to add repository for Package Management'
                return
            }
            2 {
                $Repository = $choiceList[1].HelpMessage
            }
            Default {
                $choice = $Host.UI.PromptForChoice('Select PS Repository', 'PS Repositories:', $choices, 0)
                if ($choice -eq 0) {
                    return
                }
                else {
                    $Repository = $choiceList[$choice].HelpMessage
                }
            }
        }
    }

    $repo = $script:Config.Repositories.$Repository
    $regPath = "REGISTRY::HKLM\SOFTWARE\ITAdmin\WindowsPowerShell\ProfileManagement"
    $moduleRegPath = Join-Path -Path $(Join-Path -Path $regPath -ChildPath Modules) -ChildPath Users
    Get-ChildItem -Path $moduleRegPath | Select-Object -ExpandProperty PSChildName | ForEach-Object {
        $moduleRegPath = Join-Path -Path $moduleRegPath -ChildPath $_
        if ($_ -notin $repo.Modules) {
            try {
                Uninstall-PSResource -Name $_
            }
            catch { }
        }
        else {
            $currentVersion = Get-ChildItem -Path $moduleRegPath | Sort-Object -Property PSChildName -Descending | Select-Object -First 1 -ExpandProperty PSChildName
            if ($(Find-PSResource -Name $_ -Repository $repo.PSResourceRepository | Sort-Object -Property Version -Descending | Select-Object -First 1).Version -gt $currentVersion) {
                try {
                    Update-PSResource -Name $_ -Repository $repo.PSResourceRepository -Scope CurrentUser -WarningAction Stop -ErrorAction Stop
                }
                catch {
                    Install-PSResource -Name $_ -Repository $repo.PSResourceRepository -Scope CurrentUser
                }
            }
        }
        $moduleRegPath = Split-Path -Path $moduleRegPath -Parent
    }
    $moduleRegPath = Split-Path -Path $moduleRegPath -Parent
    $repo.Modules | Where-Object { $_ -notin $(Get-ChildItem -Path $moduleRegPath).PSChildName } | ForEach-Object {
        if (!(Test-Path -Path $(Join-Path -Path $env:ONEDRIVE\Documents\WindowsPowerShell\Modules -ChildPath $_) -ErrorAction SilentlyContinue)) {
            try {
                Install-PSResource -Name $_ -Repository $repo.PSResourceRepository -Scope CurrentUser -WarningAction Stop -ErrorAction Stop
            }
            catch { }
        }
        else {
            $currentVersion = Get-ChildItem -Path $(Join-Path -Path $env:ONEDRIVE\Documents\WindowsPowerShell\Modules -ChildPath $_) | Sort-Object -Property Name -Descending | Select-Object -First 1 -ExpandProperty Name
            if ($(Find-PSResource -Name $_ -Repository $repo.PSResourceRepository | Sort-Object -Property Version -Descending | Select-Object -First 1).Version -gt $currentVersion) {
                try {
                    Update-PSResource -Name $_ -Repository $repo.PSResourceRepository -Scope CurrentUser -WarningAction Stop -ErrorAction Stop
                }
                catch { }
            }
        }
    }
    Update-ProfileRegistryConfiguration -UpdateType All
}

function Add-ProfileCategory {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript({ $_ -notin $script:Config.Keys })]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Object]
        $Value
    )

    $oldConfig = $script:Config
    $newKeys = [System.Collections.ArrayList]::New($oldConfig.Keys)
    [void]$newKeys.Add($Name)
    $newKeys = $newKeys | Sort-Object -Unique
    $newHash = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
    ForEach ($nk in $newKeys) {
        if ($nk -ne $Name) {
            $newHash.Add($nk, $oldConfig.$nk)
        }
        else {
            $newHash.Add($nk, $Value)
        }
    }
    $script:Config = $newHash
    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Name: ' -NoNewLine
    Write-Host $Name -ForegroundColor Yellow -NoNewline
    Write-Host ' - Value: ' -NoNewline
    Write-Host $Value -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Update-ProfileCategory {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Keys })]
        [System.String]
        $Name,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Object]
        $NewValue
    )

    if (!($Name)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Profile Category', 'Profile Categories:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Name = $choiceList[$choice].HelpMessage
        }
    }
    $script:Config.$($Name) = $NewValue
    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Name: ' -NoNewLine
    Write-Host $Name -ForegroundColor Yellow -NoNewline
    Write-Host ' - NewValue: ' -NoNewline
    Write-Host $NewValue -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Remove-ProfileCategory {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Keys })]
        [System.String]
        $Name,

        [Parameter(Mandatory = $false)]
        [Switch]
        $Clear
    )

    if (!($Name)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Profile Category', 'Profile Categories:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Name = $choiceList[$choice].HelpMessage
        }
    }

    if ($Clear) {
        $script:Config.$($Name).Clear()
    }
    else {
        $script:Config.Remove($Name)
    }

    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Name: ' -NoNewLine
    Write-Host $Name -ForegroundColor Yellow -NoNewline
    Write-Host ' - Clear: ' -NoNewline
    Write-Host $Clear.ToString() -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Add-ProfileCredentialType {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript({ $_ -notin $script:Config.Credentials.Keys })]
        [System.String]
        $NewType
    )

    $oldTypes = $script:Config.Credentials
    $newKeys = New-Object -TypeName System.Collections.ArrayList
    $oldTypes.Keys | ForEach-Object { [void]$newKeys.Add($_.ToString()) }
    [void]$newKeys.Add($NewType)
    $newKeys = $newKeys | Sort-Object -Unique
    $newHash = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
    ForEach ($nk in $newKeys) {
        if ($nk -ne $NewType) {
            $newHash.Add($nk, $oldTypes.$nk)
        }
        else {
            $newHash.Add($nk, $(New-Object -TypeName System.Collections.Specialized.OrderedDictionary))
        }
    }
    $script:Config.Credentials = $newHash
    $Credential = $script:Config.Credentials
    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  NewType: ' -NoNewLine
    Write-Host $NewType -ForegroundColor Yellow -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Update-ProfileCredentialType {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Credentials.Keys })]
        [System.String]
        $Type,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateScript({ $_ -notin $script:Config.Credentials.Keys })]
        [System.String]
        $NewType
    )

    if (!($Type)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Credentials.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Credential Type', 'Credential Types:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Type = $choiceList[$choice].HelpMessage
        }
    }

    $oldTypes = $script:Config.Credentials
    $newKeys = New-Object -TypeName System.Collections.ArrayList
    $oldTypes.Keys | Where-Object { $_ -ne $Type } ForEach-Object { [void]$newKeys.Add($_.ToString()) }
    [void]$newKeys.Add($NewType)
    $newKeys = $newKeys | Sort-Object -Unique
    $newHash = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
    ForEach ($nk in $newKeys) {
        if ($nk -ne $NewType) {
            $newHash.Add($nk, $oldTypes.$nk)
        }
        else {
            $newHash.Add($nk, $oldTypes.$Type)
        }
    }
    $script:Config.Credentials = $newHash
    $Credential = $script:Config.Credentials
    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Type: ' -NoNewLine
    Write-Host $Type -ForegroundColor Yellow -NoNewline
    Write-Host ' - NewType: ' -NoNewline
    Write-Host $NewType -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Remove-ProfileCredentialType {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Credentials.Keys })]
        [System.String]
        $Type,

        [Parameter(Mandatory = $false)]
        [Switch]
        $Clear
    )

    if (!($Type)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Credentials.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Credential Type', 'Credential Types:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Type = $choiceList[$choice].HelpMessage
        }
    }

    if ($Clear) {
        $script:Config.Credentials.$($Type) = $null
        $Credential = $script:Config.Credentials
        $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    }
    else {
        $script:Config.Credentials.Remove($Type)
        $Credential = $script:Config.Credentials
        $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    }
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Type: ' -NoNewLine
    Write-Host $Type -ForegroundColor Yellow -NoNewline
    Write-Host ' - Clear: ' -NoNewline
    Write-Host $Clear.ToString() -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Add-ProfileCredential {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Credentials.Keys })]
        [System.String]
        $Type,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $NewName,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $NewCredential
    )

    if (!($Type)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Credentials.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Credential Type', 'Credential Types:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Type = $choiceList[$choice].HelpMessage
        }
    }

    if (!($NewCredential)) {
        $NewCredential = $Host.UI.PromptForCredential($MyInvocation.MyCommand.Name, 'Enter new credentials', '', '')
    }

    $oldGroup = $script:Config.Credentials.$($Type)
    $newKeys = New-Object -TypeName System.Collections.ArrayList
    $oldGroup.Keys | ForEach-Object { [void]$newKeys.Add($_.ToString()) }
    [void]$newKeys.Add($NewName)
    $newKeys = $newKeys | Sort-Object -Unique
    $newHash = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
    ForEach ($nk in $newKeys) {
        if ($nk -ne $NewName) {
            $newHash.Add($nk, $oldGroup.$nk)
        }
        else {
            $newHash.Add($nk, $NewCredential)
        }
    }
    $script:Config.Credentials.$($Type) = $newHash
    $Credential = $script:Config.Credentials
    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Type: ' -NoNewLine
    Write-Host $Type -ForegroundColor Yellow -NoNewline
    Write-Host ' - NewName: ' -NoNewline
    Write-Host $NewName -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Update-ProfileCredential {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Credentials.Keys })]
        [System.String]
        $Type,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateScript({ $_ -in $script:Config.Credentials.Values.Keys })]
        [System.String]
        $Name,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $NewCredential
    )

    if (!($Type)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Credentials.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Credential Type', 'Credential Types:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Type = $choiceList[$choice].HelpMessage
        }
    }

    if (!($Name)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Credentials.$($Type).Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Credential Name', 'Credential Names:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Name = $choiceList[$choice].HelpMessage
        }
    }
    else {
        if ($Name -notin $script:Config.Credentials.$($Type).Keys) {
            Write-Error -Message "$Name not valid for $Type"
            $choiceList = New-Object -TypeName System.Collections.ArrayList
            $choiceCount = 0
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
            ForEach ($key in $script:Config.Credentials.$($Type).Keys) {
                $choiceCount++
                [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
            }
            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
            $choice = $Host.UI.PromptForChoice('Select Credential Name', 'Credential Names:', $choices, 0)
            if ($choice -eq 0) {
                return
            }
            else {
                $Name = $choiceList[$choice].HelpMessage
            }
        }
    }

    $oldCredential = $script:Config.Credentials.$($Type).$($Name)
    if (!($NewCredential)) {
        $NewCredential = $Host.UI.PromptForCredential($MyInvocation.MyCommand.Name, 'Enter new credentials', $oldCredential.UserName, '')
    }
    $script:Config.Credentials.$($Type).$($Name) = $NewCredential
    $Credential = $script:Config.Credentials
    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Type: ' -NoNewLine
    Write-Host $Type -ForegroundColor Yellow -NoNewline
    Write-Host ' - Name: ' -NoNewline
    Write-Host $Name -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Remove-ProfileCredential {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Credentials.Keys })]
        [System.String]
        $Type,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateScript({ $_ -in $script:Config.Credentials.Values.Keys })]
        [System.String]
        $Name,

        [Parameter(Mandatory = $false)]
        [Switch]
        $Clear
    )

    if (!($Type)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Credentials.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Credential Type', 'Credential Types:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Type = $choiceList[$choice].HelpMessage
        }
    }

    if (!($Name)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Credentials.$($Type).Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Credential Name', 'Credential Names:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $Name = $choiceList[$choice].HelpMessage
        }
    }
    else {
        if ($Name -notin $script:Config.Credentials.$($Type).Keys) {
            Write-Error -Message "$Name not valid for $Type"
            $choiceList = New-Object -TypeName System.Collections.ArrayList
            $choiceCount = 0
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
            ForEach ($key in $script:Config.Credentials.$($Type).Keys) {
                $choiceCount++
                [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
            }
            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
            $choice = $Host.UI.PromptForChoice('Select Credential Name', 'Credential Names:', $choices, 0)
            if ($choice -eq 0) {
                return
            }
            else {
                $Name = $choiceList[$choice].HelpMessage
            }
        }
    }

    if ($Clear) {
        $oldCredential = $script:Config.Credentials.$($Type).$($Name)
        $clearedCredential = [PSCredential]::New($oldCredential.UserName, $(ConvertTo-SecureString -String 'Clear' -AsPlainText -Force))
        $script:Config.Credentials.$($Type).$($Name) = $clearedCredential
        $Credential = $script:Config.Credentials
        $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    }
    else {
        $script:Config.Credentials.$($Type).Remove($Name)
        $Credential = $script:Config.Credentials
        $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    }
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  Type: ' -NoNewLine
    Write-Host $Type -ForegroundColor Yellow -NoNewline
    Write-Host ' - Name: ' -NoNewline
    Write-Host $Name -ForegroundColor Magenta -NoNewline
    Write-Host ' - Clear: ' -NoNewline
    Write-Host $Clear.ToString() -ForegroundColor Cyan -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Add-ProfileRepositoryConfig {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript({ $_ -notin $script:Config.Repositories.Keys })]
        [System.String]
        $BaseName,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Url,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Modules,

        [Parameter(Mandatory = $false)]
        [Switch]
        $Trusted
    )

    $oldRepositories = $script:Config.Repositories
    $newKeys = New-Object -TypeName System.Collections.ArrayList
    $oldRepositories.Keys | ForEach-Object { [void]$newKeys.Add($_.ToString()) }
    [void]$newKeys.Add($BaseName)
    $newKeys = $newKeys | Sort-Object -Unique
    $newHash = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
    ForEach ($nk in $newKeys) {
        if ($nk -ne $BaseName) {
            $newHash.Add($nk, $oldRepositories.$nk)
        }
        else {
            $newHash.Add($nk, $(New-Object -TypeName System.Collections.Specialized.OrderedDictionary))
        }
    }
    $newHash.$BaseName.Add('PSRepository', "${BaseName}_Gallery")
    $newHash.$BaseName.Add('SourceLocation', $Url)
    $newHash.$BaseName.Add('PublishLocation', "${Url}\Packages")
    $newHash.$BaseName.Add('ScriptSourceLocation', $Url)
    $newHash.$BaseName.Add('ScriptPublishLocation', "${Url}\ScriptPackages")
    if ($Trusted) {
        $newHash.$BaseName.Add('InstallationPolicy', 'Trusted')
        $newHash.$BaseName.Add('PSResourceRepository', "${BaseName}_Resources")
        $newHash.$BaseName.Add('Url', $Url)
        $newHash.$BaseName.Add('Trusted', $true)
    }
    else {
        $newHash.$BaseName.Add('InstallationPolicy', 'Untrusted')
        $newHash.$BaseName.Add('PSResourceRepository', "${BaseName}_Resources")
        $newHash.$BaseName.Add('Url', $Url)
        $newHash.$BaseName.Add('Trusted', $false)
    }
    if ($Modules) {
        $newHash.$BaseName.Add('Modules', $Modules)
    }
    else {
        $newHash.$BaseName.Add('Modules', $(New-Object -TypeName System.Collections.ArrayList))
    }
    $script:Config.Repositories = $newHash
    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  BaseName: ' -NoNewLine
    Write-Host $BaseName -ForegroundColor Yellow -NoNewline
    Write-Host ' - Trusted: ' -NoNewline
    Write-Host $Trusted.ToString() -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Update-ProfileRepositoryConfig {

    [CmdletBinding(DefaultParameterSetName = 'Basic')]

    Param (

        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Basic')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Name')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Url')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Modules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Trusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Untrusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUrl')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameTrusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUntrusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUrlModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUrlTrusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUrlUntrusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameTrustedModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUntrustedModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUrlTrustedModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'NameUrlUntrustedModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'UrlModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'UrlTrusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'UrlUntrusted')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'UrlTrustedModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'UrlUntrustedModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'TrustedModules')]
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'UntrustedModules')]
        [ValidateScript({ $_ -in $script:Config.Repositories.Keys })]
        [System.String]
        $BaseName,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'Name')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUrl')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameModules')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameTrusted')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUntrusted')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUrlModules')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUrlTrusted')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUrlUntrusted')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameTrustedModules')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUntrustedModules')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUrlTrustedModules')]
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'NameUrlUntrustedModules')]
        [ValidateScript({ $_ -notin $script:Config.Repositories.Keys })]
        [System.String]
        $NewBaseName,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'Url')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'NameUrl')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'NameUrlModules')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'NameUrlTrusted')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'NameUrlUntrusted')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'NameUrlTrustedModules')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'NameUrlUntrustedModules')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'UrlModules')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'UrlTrusted')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'UrlUntrusted')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'UrlTrustedModules')]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'UrlUntrustedModules')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $NewUrl,

        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'Modules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'NameModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'UrlModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'TrustedModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'UntrustedModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'NameUrlModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'NameTrustedModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'NameUntrustedModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'NameUrlTrustedModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'NameUrlUntrustedModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'UrlTrustedModules')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'UrlUntrustedModules')]
        [ValidateNotNull()]
        [AllowEmptyCollection()]
        [System.String[]]
        $Modules,

        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'Modules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'NameModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'UrlModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'TrustedModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'UntrustedModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'NameUrlModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'NameTrustedModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'NameUntrustedModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'NameUrlTrustedModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'NameUrlUntrustedModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'UrlTrustedModules')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'UrlUntrustedModules')]
        [ValidateSet('Add', 'Overwrite', 'Remove')]
        [System.String]
        $ModuleUpdateType = 'Add',

        [Parameter(Mandatory = $true, ParameterSetName = 'Trusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameTrusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'UrlTrusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'TrustedModules')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameUrlTrusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameTrustedModules')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameUrlTrustedModules')]
        [Parameter(Mandatory = $true, ParameterSetName = 'UrlTrustedModules')]
        [Switch]
        $Trusted,

        [Parameter(Mandatory = $true, ParameterSetName = 'Untrusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameUntrusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'UrlUntrusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'UntrustedModules')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameUrlUntrusted')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameUntrustedModules')]
        [Parameter(Mandatory = $true, ParameterSetName = 'NameUrlUntrustedModules')]
        [Parameter(Mandatory = $true, ParameterSetName = 'UrlUntrustedModules')]
        [Switch]
        $Untrusted
    )

    if (!($BaseName)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Repositories.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Profile Repository', 'Profile Repositories:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $BaseName = $choiceList[$choice].HelpMessage
        }
    }

    if ($NewBaseName) {
        if ($BaseName -ne $NewBaseName) {
            $oldRepositories = $script:Config.Repositories
            $newKeys = New-Object -TypeName System.Collections.ArrayList
            $oldRepositories.Keys | Where-Object { $_ -ne $BaseName } | ForEach-Object { [void]$newKeys.Add($_.ToString()) }
            [void]$newKeys.Add($NewBaseName)
            $newKeys = $newKeys | Sort-Object -Unique
            $newHash = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
            ForEach ($nk in $newKeys) {
                if ($nk -ne $NewBaseName) {
                    $newHash.Add($nk, $oldRepositories.$nk)
                }
                else {
                    $newHash.Add($nk, $oldRepositories.$BaseName)
                }
            }
            $script:Config.Repositories = $newHash
        }
        $currentRepository = $script:Config.Repositories.$($NewBaseName)
        $currentRepository.PSRepository = "${NewBaseName}_Gallery"
        $currentRepository.PSResourceRepository = "${NewBaseName}_Resources"
    }
    else {
        $currentRepository = $script:Config.Repositories.$($BaseName)
    }

    if ($NewUrl) {
        if ($currentRepository.Url -ne $NewUrl) {
            $currentRepository.SourceLocation = $NewUrl
            $currentRepository.PublishLocation = "${NewUrl}\Packages"
            $currentRepository.ScriptSourceLocation = $NewUrl
            $currentRepository.ScriptPublishLocation = "${NewUrl}\ScriptPackages"
            $currentRepository.Url = $NewUrl
        }
    }

    if ($Modules) {
        Switch ($ModuleUpdateType) {
            'Add' {
                $newModules = [System.Collections.ArrayList]::New($currentRepository.Modules)
                ForEach ($m in $Modules) {
                    [void]$newModules.Add($m)
                }
                $newModules = $newModules | Sort-Object -Unique
                $currentRepository.Modules = $newModules
            }
            'Overwrite' {
                $currentRepository.Modules = $Modules | Sort-Object -Unique
            }
            'Remove' {
                $currentRepository.Modules = $currentRepository.Modules | Where-Object { $_ -notin $Modules } | Sort-Object -Unique
            }
        }
    }

    if ($Trusted) {
        $currentRepository.InstallationPolicy = 'Trusted'
        $currentRepository.Trusted = $true
    }

    if ($Untrusted) {
        $currentRepository.InstallationPolicy = 'Untrusted'
        $currentRepository.Trusted = $false
    }

    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  BaseName: ' -NoNewLine
    Write-Host $BaseName -ForegroundColor Yellow -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}

function Remove-ProfileRepositoryConfig {

    [CmdletBinding()]

    Param (

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateScript({ $_ -in $script:Config.Repositories.Keys })]
        [System.String]
        $BaseName,

        [Parameter(Mandatory = $false)]
        [Switch]
        $Clear
    )

    if (!($BaseName)) {
        $choiceList = New-Object -TypeName System.Collections.ArrayList
        $choiceCount = 0
        [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} Cancel", 'Cancel'))
        ForEach ($key in $script:Config.Repositories.Keys) {
            $choiceCount++
            [void]$choiceList.Add([System.Management.Automation.Host.ChoiceDescription]::New("&${choiceCount} $key", $key))
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList)
        $choice = $Host.UI.PromptForChoice('Select Profile Repository', 'Profile Repositories:', $choices, 0)
        if ($choice -eq 0) {
            return
        }
        else {
            $BaseName = $choiceList[$choice].HelpMessage
        }
    }

    if ($Clear) {
        $clearRepository = $script:Config.Repositories.$($BaseName)
        ForEach ($key in $clearRepository.Keys) {
            if ($key -ne 'Modules') {
                $clearRepository.$($key) = $null
            }
            else {
                $clearRepository.$($key) = $(New-Object -TypeName System.Collections.ArrayList)
            }
        }
    }
    else {
        $script:Config.Repositories.Remove($BaseName)
    }

    $script:Config | Export-Clixml -Path $PSScriptRoot\config.xml -Depth 100
    Write-Host $MyInvocation.MyCommand.Name -NoNewline
    Write-Host '  ||  BaseName: ' -NoNewLine
    Write-Host $BaseName -ForegroundColor Yellow -NoNewline
    Write-Host ' - Clear: ' -NoNewline
    Write-Host $Clear.ToString() -ForegroundColor Magenta -NoNewline
    Write-Host ' - Status: ' -NoNewline
    Write-Host 'Complete' -ForegroundColor Green
    pause
}
