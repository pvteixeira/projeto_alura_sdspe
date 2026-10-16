<#
.SYNOPSIS
    USM-Install (Unattended Software Manager) - Automação de Instalação Padronizada de Softwares.

.DESCRIPTION
    Script corporativo em PowerShell para provisionamento automatizado de softwares via Chocolatey.
    
    Suporta:
    - Exibição de acentuação correta no console (UTF-8 com BOM).
    - Detecção visual de softwares já instalados no sistema [✓ Já Instalado].
    - Opção de baixar novamente (reinstalar/atualizar) ou manter sem baixar de novo.
    - Loop interativo: permite instalar outros programas sem precisar reabrir o script.
    - Menu com números ou janela gráfica com caixas de seleção (Out-GridView).
    - Auto-elevação de privilégios (UAC) para execução local.
    - Suporte a execução remota via WinRM e Jobs assíncronos.
#>

[CmdletBinding(DefaultParameterSetName = 'Direct')]
param(
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName = @('localhost'),

    [Parameter()]
    [string[]]$Packages,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [switch]$AsJob,

    [Parameter()]
    [string]$JobName = "USM-Install-Job",

    [Parameter()]
    [switch]$Upgrade,

    [Parameter()]
    [switch]$NoPause
)

begin {
    # Garante acentuação perfeita no console do Windows PowerShell
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $OutputEncoding = [System.Text.Encoding]::UTF8
    } catch {}

    # Catálogo corporativo oficial de softwares homologados
    $DefaultCatalog = @(
        # Runtimes, SDKs & Compiladores
        [PSCustomObject]@{ Id = 'temurin21';                  Name = 'Eclipse Temurin JDK 21 (Java)';  Category = 'Runtimes & SDKs' }
        [PSCustomObject]@{ Id = 'vcredist-all';               Name = 'Visual C++ All-in-One (2005-22)';Category = 'Runtimes & SDKs' }
        [PSCustomObject]@{ Id = 'python3';                    Name = 'Python 3';                       Category = 'Runtimes & SDKs' }

        # Ferramentas Essenciais de Infraestrutura & Diagnóstico
        [PSCustomObject]@{ Id = 'sysinternals';               Name = 'Sysinternals Suite';             Category = 'Infra & Diagnóstico' }
        [PSCustomObject]@{ Id = 'wireshark';                  Name = 'Wireshark';                      Category = 'Redes & Diagnóstico' }
        [PSCustomObject]@{ Id = 'advanced-ip-scanner';        Name = 'Advanced IP Scanner';            Category = 'Redes & Diagnóstico' }
        [PSCustomObject]@{ Id = 'putty.install';              Name = 'PuTTY (SSH/Telnet)';             Category = 'Acesso Remoto' }
        [PSCustomObject]@{ Id = 'winscp.install';             Name = 'WinSCP (SFTP/FTP)';              Category = 'Acesso Remoto' }

        # Terminais Modernos & Automação
        [PSCustomObject]@{ Id = 'microsoft-windows-terminal'; Name = 'Windows Terminal';               Category = 'Terminais & Shells' }
        [PSCustomObject]@{ Id = 'powershell-core';            Name = 'PowerShell 7 (pwsh)';            Category = 'Terminais & Shells' }

        # Desenvolvimento, IDEs & Containers
        [PSCustomObject]@{ Id = 'vscode';                     Name = 'Visual Studio Code';             Category = 'Desenvolvimento' }
        [PSCustomObject]@{ Id = 'visualstudio2022community';  Name = 'Visual Studio 2022 Community';   Category = 'Desenvolvimento' }
        [PSCustomObject]@{ Id = 'git.install';                Name = 'Git';                            Category = 'Desenvolvimento' }
        [PSCustomObject]@{ Id = 'docker-desktop';             Name = 'Docker Desktop';                 Category = 'Containers & Infra' }
        [PSCustomObject]@{ Id = 'dbeaver';                    Name = 'DBeaver Community (SQL)';       Category = 'Banco de Dados' }

        # Navegadores
        [PSCustomObject]@{ Id = 'googlechrome';               Name = 'Google Chrome';                  Category = 'Navegadores' }
        [PSCustomObject]@{ Id = 'firefox';                    Name = 'Mozilla Firefox';                Category = 'Navegadores' }
        [PSCustomObject]@{ Id = 'brave';                      Name = 'Brave Browser';                  Category = 'Navegadores' }

        # Produtividade & Utilitários
        [PSCustomObject]@{ Id = '7zip.install';               Name = '7-Zip';                          Category = 'Utilitários' }
        [PSCustomObject]@{ Id = 'notepadplusplus.install';    Name = 'Notepad++';                      Category = 'Editores' }
        [PSCustomObject]@{ Id = 'powertoys';                  Name = 'Microsoft PowerToys';            Category = 'Utilitários' }
        [PSCustomObject]@{ Id = 'adobereader';                Name = 'Adobe Acrobat Reader DC';        Category = 'Produtividade' }
        [PSCustomObject]@{ Id = 'vlc';                        Name = 'VLC Media Player';               Category = 'Multimídia' }
        [PSCustomObject]@{ Id = 'spotify';                    Name = 'Spotify';                        Category = 'Multimídia' }
        [PSCustomObject]@{ Id = 'discord';                    Name = 'Discord';                        Category = 'Comunicação' }
    )

    #region Funções de Log e Validação Local
    function Write-USMLog {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter()]
            [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
            [string]$Level = 'INFO',

            [Parameter()]
            [string]$Target = $env:COMPUTERNAME
        )

        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $color = switch ($Level) {
            'INFO'    { 'Cyan' }
            'SUCCESS' { 'Green' }
            'WARNING' { 'Yellow' }
            'ERROR'   { 'Red' }
        }

        $formatted = "[$timestamp] [$Level] [$Target] $Message"
        Write-Host $formatted -ForegroundColor $color
    }

    function Test-IsAdministrator {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]$identity
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    function Invoke-ElevatedSelf {
        param([string[]]$ScriptArgs)
        Write-USMLog -Message "Solicitando elevação de privilégios de Administrador (UAC)..." -Level 'WARNING'
        
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) {
            $scriptPath = $MyInvocation.MyCommand.Definition
        }

        $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"")
        if ($ScriptArgs) {
            $argList += $ScriptArgs
        }

        try {
            Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Verb RunAs
            Write-USMLog -Message "Janela elevada como Administrador aberta com sucesso." -Level 'SUCCESS'
        }
        catch {
            Write-USMLog -Message "A elevação de privilégios foi cancelada ou falhou: $_" -Level 'ERROR'
        }
    }

    # Obtém o mapa de softwares já instalados no Chocolatey
    function Get-InstalledPackagesMap {
        $installed = @{}
        $chocoBin = "$env:ProgramData\chocolatey\bin\choco.exe"
        if (Test-Path $chocoBin) {
            $raw = & $chocoBin list --local-only --limit-output 2>$null
            foreach ($line in $raw) {
                if ($line -match '^([^|]+)\|([^|]+)') {
                    $installed[$matches[1].ToLower()] = $matches[2]
                }
            }
        }
        return $installed
    }

    function Show-PackageMenu {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [array]$Catalog,

            [Parameter()]
            [hashtable]$InstalledMap = @{}
        )

        Write-Host ""
        Write-Host "==========================================================================================" -ForegroundColor Cyan
        Write-Host "                       MENU DE SELEÇÃO DE SOFTWARES - USM-Install                         " -ForegroundColor White
        Write-Host "==========================================================================================" -ForegroundColor Cyan
        Write-Host " Escolha quais softwares deseja instalar no computador:                                   " -ForegroundColor Gray
        Write-Host "------------------------------------------------------------------------------------------" -ForegroundColor DarkCyan

        for ($i = 0; $i -lt $Catalog.Count; $i++) {
            $item = $Catalog[$i]
            $indexStr = ($i + 1).ToString().PadLeft(2, ' ')
            $nameStr = $item.Name.PadRight(35, ' ')
            $catStr = "[$($item.Category)]".PadRight(25, ' ')
            
            $isInstalled = $InstalledMap.ContainsKey($item.Id.ToLower())
            if ($isInstalled) {
                $statusStr = "[Já Instalado v$($InstalledMap[$item.Id.ToLower()])]"
                Write-Host "  [$indexStr] $nameStr $catStr " -NoNewline -ForegroundColor Yellow
                Write-Host $statusStr -ForegroundColor Green
            } else {
                $statusStr = "[Disponível]"
                Write-Host "  [$indexStr] $nameStr $catStr " -NoNewline -ForegroundColor Yellow
                Write-Host $statusStr -ForegroundColor DarkGray
            }
        }

        Write-Host "------------------------------------------------------------------------------------------" -ForegroundColor DarkCyan
        Write-Host "  [ T] Instalar TODOS os softwares da lista" -ForegroundColor Green
        Write-Host "  [ G] Abrir janela com CHECKBOXES (Seleção Gráfica com mouse)" -ForegroundColor Magenta
        Write-Host "  [ S] Sair / Cancelar" -ForegroundColor Red
        Write-Host "==========================================================================================" -ForegroundColor Cyan
        Write-Host ""

        while ($true) {
            $choice = (Read-Host " Digite sua opção (ex: 1, 4, 7 ou G ou T)").Trim()

            if ([string]::IsNullOrWhiteSpace($choice)) {
                Write-Host " Por favor, digite os números desejados ou 'S' para sair." -ForegroundColor Yellow
                continue
            }

            if ($choice -match '^(s|sair|cancelar|exit|q)$') {
                Write-Host " Operação cancelada pelo usuário." -ForegroundColor Yellow
                return @()
            }

            if ($choice -match '^(t|todos|all)$') {
                Write-Host " Selecionados: TODOS os softwares do catálogo." -ForegroundColor Green
                return @($Catalog | ForEach-Object { $_.Id })
            }

            if ($choice -match '^(g|gui|grid)$') {
                try {
                    $gridItems = @()
                    foreach ($it in $Catalog) {
                        $already = $InstalledMap.ContainsKey($it.Id.ToLower())
                        $gridItems += [PSCustomObject]@{
                            Id        = $it.Id
                            Nome      = $it.Name
                            Categoria = $it.Category
                            Status    = if ($already) { "Já Instalado (v$($InstalledMap[$it.Id.ToLower()]))" } else { "Disponível para Instalar" }
                        }
                    }

                    $selectedFromGrid = $gridItems | Out-GridView -Title "USM-Install: Selecione os Softwares desejados e clique em OK (Segure Ctrl para selecionar múltiplos)" -OutputMode Multiple
                    if ($selectedFromGrid) {
                        return @($selectedFromGrid | ForEach-Object { $_.Id })
                    } else {
                        Write-Host " Nenhum software selecionado na janela gráfica. Tente novamente ou digite 'S'." -ForegroundColor Yellow
                        continue
                    }
                }
                catch {
                    Write-Host " Interface gráfica indisponível neste terminal. Digite os números da lista." -ForegroundColor Yellow
                    continue
                }
            }

            # Parse de múltiplos números separados por vírgula, espaço ou ponto-e-vírgula
            $tokens = $choice -split '[,;\s]+' | Where-Object { $_ -ne '' }
            $selectedIds = @()
            $valid = $true

            foreach ($token in $tokens) {
                if ($token -match '^\d+$') {
                    $num = [int]$token
                    if ($num -ge 1 -and $num -le $Catalog.Count) {
                        $selectedItem = $Catalog[$num - 1]
                        if ($selectedItem.Id -notin $selectedIds) {
                            $selectedIds += $selectedItem.Id
                        }
                    } else {
                        Write-Host " Número '$num' fora do intervalo permitido (1 a $($Catalog.Count))." -ForegroundColor Red
                        $valid = $false
                    }
                } else {
                    Write-Host " Opção inválida: '$token'." -ForegroundColor Red
                    $valid = $false
                }
            }

            if ($valid -and $selectedIds.Count -gt 0) {
                return $selectedIds
            }
        }
    }
    #endregion

    # Bloco encapsulado de execução para o alvo (compatível com sessões locais e remotas)
    $WorkerScriptBlock = {
        param(
            [array]$SoftwareTasks # Array de PSCustomObject com Id e ForceReinstall
        )

        function Write-InternalLog {
            param([string]$Message, [string]$Level = 'INFO')
            $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $color = switch ($Level) {
                'INFO'    { 'Cyan' }
                'SUCCESS' { 'Green' }
                'WARNING' { 'Yellow' }
                'ERROR'   { 'Red' }
            }
            Write-Host "[$timestamp] [$Level] [$env:COMPUTERNAME] $Message" -ForegroundColor $color
        }

        Write-InternalLog -Message "Iniciando provisionamento USM-Install..." -Level 'INFO'

        # 1. Validação de Privilégios de Administrador
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            Write-InternalLog -Message "ERRO CRÍTICO: O script requer privilégios de Administrador para instalar softwares." -Level 'ERROR'
            throw "Privilégios insuficientes na máquina $env:COMPUTERNAME."
        }

        # 2. Bootstrap do Chocolatey
        $chocoInstalled = (Get-Command choco.exe -ErrorAction SilentlyContinue) -ne $null
        $chocoBinPath = "$env:ProgramData\chocolatey\bin"

        if (-not $chocoInstalled -and (Test-Path "$chocoBinPath\choco.exe")) {
            $env:Path = "$chocoBinPath;$env:Path"
            $chocoInstalled = $true
        }

        if (-not $chocoInstalled) {
            Write-InternalLog -Message "Chocolatey não detectado. Iniciando bootstrap oficial..." -Level 'WARNING'
            try {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
                Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

                $installScript = ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
                Invoke-Expression $installScript

                $env:Path = "$chocoBinPath;$env:Path"
                
                if (Test-Path "$chocoBinPath\choco.exe") {
                    Write-InternalLog -Message "Chocolatey instalado com sucesso!" -Level 'SUCCESS'
                } else {
                    throw "Falha ao verificar binário do Chocolatey após download."
                }
            }
            catch {
                Write-InternalLog -Message "Erro durante a instalação do Chocolatey: $_" -Level 'ERROR'
                throw $_
            }
        } else {
            $chocoVersion = & choco -v 2>&1
            Write-InternalLog -Message "Chocolatey pronto para uso (Versão: $chocoVersion)." -Level 'INFO'
        }

        # 3. Gerenciamento e Instalação dos Pacotes Homologados
        Write-InternalLog -Message "Total de pacotes a processar: $($SoftwareTasks.Count)" -Level 'INFO'

        foreach ($task in $SoftwareTasks) {
            $pkg = if ($task -is [hashtable] -or $task -is [PSCustomObject]) { $task.Id } else { [string]$task }
            $force = if ($task -is [hashtable] -or $task -is [PSCustomObject]) { [bool]$task.ForceReinstall } else { $false }

            if ([string]::IsNullOrWhiteSpace($pkg)) { continue }

            Write-InternalLog -Message "Processando pacote: '$pkg'..." -Level 'INFO'

            try {
                if ($force) {
                    Write-InternalLog -Message "Baixando novamente / Forçando reinstalação de '$pkg'..." -Level 'INFO'
                    $process = Start-Process -FilePath "choco.exe" -ArgumentList "upgrade", $pkg, "-y", "--force", "--no-progress" -Wait -PassThru -NoNewWindow
                } else {
                    Write-InternalLog -Message "Instalando pacote '$pkg' em modo silencioso..." -Level 'INFO'
                    $process = Start-Process -FilePath "choco.exe" -ArgumentList "install", $pkg, "-y", "--no-progress" -Wait -PassThru -NoNewWindow
                }

                switch ($process.ExitCode) {
                    0 {
                        Write-InternalLog -Message "Pacote '$pkg' concluído com SUCESSO!" -Level 'SUCCESS'
                    }
                    3010 {
                        Write-InternalLog -Message "Pacote '$pkg' concluído com SUCESSO (Requer reinicialização do sistema)." -Level 'WARNING'
                    }
                    default {
                        Write-InternalLog -Message "Falha na instalação de '$pkg'. Código de saída: $($process.ExitCode)" -Level 'ERROR'
                    }
                }
            }
            catch {
                Write-InternalLog -Message "Exceção ao tentar instalar '$pkg': $_" -Level 'ERROR'
            }
        }

        Write-InternalLog -Message "Rotina de provisionamento concluída para $env:COMPUTERNAME." -Level 'SUCCESS'
    }
}

process {
    Write-Host @"
========================================================================
     USM-Install | Unattended Software Manager via PowerShell & Choco   
========================================================================
"@ -ForegroundColor Cyan

    # 1. Separar destinos: Local vs Remotos
    $localTargets = @()
    $remoteTargets = @()

    foreach ($comp in $ComputerName) {
        if ($comp -in @('localhost', '127.0.0.1', '.', $env:COMPUTERNAME)) {
            $localTargets += $comp
        } else {
            $remoteTargets += $comp
        }
    }

    # 2. Execução Local
    if ($localTargets.Count -gt 0) {
        Write-USMLog -Message "Alvo local identificado: $env:COMPUTERNAME" -Level 'INFO'

        # Auto-elevação UAC caso não seja Administrador
        if (-not (Test-IsAdministrator)) {
            Write-USMLog -Message "O script requer privilégios de Administrador." -Level 'WARNING'
            Write-USMLog -Message "Disparando elevação UAC automática em nova janela..." -Level 'INFO'
            
            $reconstructedArgs = @()
            if ($PSBoundParameters.ContainsKey('Packages') -and $Packages) {
                $reconstructedArgs += "-Packages"
                $reconstructedArgs += ($Packages -join ',')
            }
            if ($Upgrade) { $reconstructedArgs += "-Upgrade" }
            if ($AsJob) { $reconstructedArgs += "-AsJob" }

            Invoke-ElevatedSelf -ScriptArgs $reconstructedArgs
            return
        }

        # Garante que o Chocolatey já esteja no PATH da sessão elevada
        $chocoBinPath = "$env:ProgramData\chocolatey\bin"
        if (Test-Path "$chocoBinPath\choco.exe") {
            $env:Path = "$chocoBinPath;$env:Path"
        }

        # Loop interativo principal
        $keepRunning = $true
        while ($keepRunning) {
            $installedMap = Get-InstalledPackagesMap
            $taskList = @()

            # Caso pacotes tenham sido informados diretamente na linha de comando
            if ($PSBoundParameters.ContainsKey('Packages') -and $Packages -and $Packages.Count -gt 0) {
                foreach ($p in $Packages) {
                    $taskList += [PSCustomObject]@{ Id = $p; ForceReinstall = [bool]$Upgrade }
                }
                $keepRunning = $false # Execução direta via parâmetro não faz loop
            } else {
                # Exibe Menu com status de cada programa
                $chosen = Show-PackageMenu -Catalog $DefaultCatalog -InstalledMap $installedMap
                if (-not $chosen -or $chosen.Count -eq 0) {
                    Write-USMLog -Message "Nenhum software selecionado. Operação encerrada." -Level 'WARNING'
                    break
                }

                Write-Host ""
                Write-Host "Verificando status dos softwares escolhidos..." -ForegroundColor Cyan

                foreach ($pkgId in $chosen) {
                    $item = $DefaultCatalog | Where-Object { $_.Id -eq $pkgId }
                    $name = if ($item) { $item.Name } else { $pkgId }
                    $alreadyInstalled = $installedMap.ContainsKey($pkgId.ToLower())

                    if ($alreadyInstalled) {
                        $ver = $installedMap[$pkgId.ToLower()]
                        Write-Host ""
                        Write-Host "[!] O software '$name' já está instalado no computador (Versão: $ver)." -ForegroundColor Yellow
                        $reinstallAnswer = Read-Host "    Deseja baixar novamente / reinstalar? (S/N) [Padrão: N]"
                        
                        if ($reinstallAnswer -match '^(s|sim|y|yes)$') {
                            Write-Host "    -> Configurado para BAIXAR NOVAMENTE / REINSTALAR." -ForegroundColor Green
                            $taskList += [PSCustomObject]@{ Id = $pkgId; ForceReinstall = $true }
                        } else {
                            Write-Host "    -> Mantendo versão instalada (NÃO será baixado novamente)." -ForegroundColor DarkGray
                        }
                    } else {
                        Write-Host "  [+] '$name' pronto para primeiro download e instalação." -ForegroundColor Green
                        $taskList += [PSCustomObject]@{ Id = $pkgId; ForceReinstall = $false }
                    }
                }

                if ($taskList.Count -eq 0) {
                    Write-Host ""
                    Write-USMLog -Message "Nenhum pacote selecionado para download/instalação." -Level 'INFO'
                } else {
                    Write-Host ""
                    Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
                    Write-Host "Resumo de softwares para instalar/baixar:" -ForegroundColor White
                    foreach ($t in $taskList) {
                        $it = $DefaultCatalog | Where-Object { $_.Id -eq $t.Id }
                        $label = if ($it) { $it.Name } else { $t.Id }
                        $mode = if ($t.ForceReinstall) { "[Baixar Novamente / Forçar]" } else { "[Nova Instalação]" }
                        Write-Host "  • $label $mode" -ForegroundColor Green
                    }
                    Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan

                    $confirm = Read-Host "Confirma a execução agora? (S/N) [Padrão: S]"
                    if ($confirm -match '^(n|nao|não|no)$') {
                        Write-USMLog -Message "Operação cancelada pelo operador." -Level 'WARNING'
                    } else {
                        if ($AsJob) {
                            Write-USMLog -Message "Disparando instalação em segundo plano (PowerShell Job)..." -Level 'INFO'
                            $currentJobName = "$JobName-Local-$((Get-Date).ToString('HHmmss'))"
                            $job = Start-Job -Name $currentJobName -ScriptBlock $WorkerScriptBlock -ArgumentList (,$taskList)
                            Write-USMLog -Message "Job local criado! [ID: $($job.Id) | Nome: $($job.Name)]" -Level 'SUCCESS'
                            Write-USMLog -Message "Acompanhe com: Get-Job -Id $($job.Id) | Receive-Job" -Level 'INFO'
                        } else {
                            Write-USMLog -Message "Iniciando processo de instalação..." -Level 'INFO'
                            & $WorkerScriptBlock -SoftwareTasks $taskList
                        }
                    }
                }

                # Pergunta se deseja baixar/escolher outro software sem fechar a janela
                Write-Host ""
                $again = Read-Host "Deseja escolher e instalar outros softwares agora? (S/N) [Padrão: N]"
                if ($again -notmatch '^(s|sim|y|yes)$') {
                    $keepRunning = $false
                }
            }
        }
    }

    # 3. Execução Remota (WinRM / Invoke-Command)
    if ($remoteTargets.Count -gt 0) {
        $remoteTaskList = @()
        if (-not $PSBoundParameters.ContainsKey('Packages') -or -not $Packages -or $Packages.Count -eq 0) {
            Write-USMLog -Message "Nenhum pacote especificado para o deploy remoto. Selecione os pacotes:" -Level 'INFO'
            $chosen = Show-PackageMenu -Catalog $DefaultCatalog
            if (-not $chosen -or $chosen.Count -eq 0) {
                Write-USMLog -Message "Nenhum software selecionado. Cancelando deploy remoto." -Level 'WARNING'
                return
            }
            foreach ($p in $chosen) {
                $remoteTaskList += [PSCustomObject]@{ Id = $p; ForceReinstall = [bool]$Upgrade }
            }
        } else {
            foreach ($p in $Packages) {
                $remoteTaskList += [PSCustomObject]@{ Id = $p; ForceReinstall = [bool]$Upgrade }
            }
        }

        Write-USMLog -Message "Processando $($remoteTargets.Count) alvo(s) remoto(s): $($remoteTargets -join ', ')" -Level 'INFO'

        $invokeParams = @{
            ComputerName = $remoteTargets
            ScriptBlock  = $WorkerScriptBlock
            ArgumentList = (,$remoteTaskList)
            ErrorAction  = 'Continue'
        }

        if ($Credential) {
            $invokeParams['Credential'] = $Credential
        }

        foreach ($target in $remoteTargets) {
            try {
                $wsman = Test-WSMan -ComputerName $target -ErrorAction Stop
                Write-USMLog -Message "WinRM acessível em '$target' (OS: $($wsman.ProductVersion))." -Level 'SUCCESS' -Target $target
            } catch {
                Write-USMLog -Message "Não foi possível conectar via WinRM em '$target'. Verifique o serviço WinRM e o Firewall." -Level 'WARNING' -Target $target
            }
        }

        if ($AsJob) {
            $currentJobName = "$JobName-Remote-$((Get-Date).ToString('HHmmss'))"
            $invokeParams['AsJob'] = $true
            $invokeParams['JobName'] = $currentJobName

            Write-USMLog -Message "Disparando tarefas remotas em segundo plano (Invoke-Command -AsJob)..." -Level 'INFO'
            $remoteJob = Invoke-Command @invokeParams
            Write-USMLog -Message "Job remoto criado! [ID: $($remoteJob.Id) | Nome: $($remoteJob.Name)]" -Level 'SUCCESS'
            Write-Host @"
------------------------------------------------------------------------
Comandos úteis para monitorar este Job:
  • Ver status:    Get-Job -Id $($remoteJob.Id)
  • Aguardar fim:  Wait-Job -Id $($remoteJob.Id)
  • Obter logs:    Receive-Job -Id $($remoteJob.Id) -Keep
------------------------------------------------------------------------
"@ -ForegroundColor Yellow
        } else {
            Write-USMLog -Message "Executando tarefas remotas de forma síncrona..." -Level 'INFO'
            Invoke-Command @invokeParams
        }
    }
}

end {
    Write-USMLog -Message "Execução do comando USM-Install finalizada." -Level 'INFO'

    # Pausa final para não fechar a tela
    if (-not $NoPause -and -not $AsJob -and ($localTargets.Count -gt 0) -and (Test-IsAdministrator)) {
        Write-Host ""
        Write-Host "========================================================================" -ForegroundColor DarkGray
        Write-Host " Processo finalizado! Pressione ENTER para fechar esta janela...       " -ForegroundColor Yellow
        Write-Host "========================================================================" -ForegroundColor DarkGray
        $null = Read-Host
    }
}
