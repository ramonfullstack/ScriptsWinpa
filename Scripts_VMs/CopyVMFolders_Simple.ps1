# Script para copiar pastas Panther e OEM das VMs Azure para maquina local

# Configuracoes globais
$ErrorActionPreference = "Continue"

# Lista de Resource Groups para processar
$ResourceGroups = @(
    "windows-2016-noncvm-kakothari-1001113343",
    "windows-2016-datacenter-gen2-noncvm-kakothari-1001113436",
    "windows-2016-datacenter-core-noncvm-kakothari-1001113524",
    "windows-2016-datacenter-core-g2-noncvm-kakothari-1001113614",
    "windows-2019-datacenter-gen2-noncvm-kakothari-1001113705",
    "windows-2019-datacenter-core-noncvm-kakothari-1001113802",
    "windows-2019-datacenter-core-g2-noncvm-kakothari-1001113859",
    "windows-2022-datacenter-gen1-noncvm-kakothari-1001113944",
    "windows-2022-datacenter-gen2-noncvm-kakothari-1001114051",
    "windows-2022-datacenter-azure-core-g2-noncvm-kakothari-1001114150",
    "windows-2022-datacenter-azure-gen2-noncvm-kakothari-1001114247",
    "windows-2022-datacenter-servercore-gen1-noncvm-kakothari-1001114344",
    "windows-2022-datacenter-servercore-gen2-noncvm-kakothari-1001114433",
    "windows-2025-datacenter-azure-edition-core-gen2-noncvm-kakothari-1001114519",
    "windows-2025-datacenter-azure-edition-gen2-noncvm-kakothari-1001114616",
    "windows-2025-datacenter-servercore-gen1-noncvm-kakothari-1001114720",
    "windows-2025-datacenter-servercore-gen2-noncvm-kakothari-1001114817",
    "windows-2025-datacenter-gen1-noncvm-kakothari-1001114909",
    "windows-2025-datacenter-gen2-noncvm-kakothari-1001115015",
    "windows-10-enterprise-22h2-gen1-noncvm-kakothari-1001115116",
    "windows-10-enterprise-22h2-gen2-noncvm-kakothari-1001115346",
    "windows-11-enterprise-22h2-gen2-noncvm-kakothari-1001115603",
    "windows-11-enterprise-23h2-gen2-noncvm-kakothari-1001115853",
    "windows-11-enterprise-24h2-gen2-noncvm-kakothari-1001120147"
)

# Pastas para copiar das VMs
$FoldersToSync = @(
    "C:\Windows\Panther",
    "C:\Windows\OEM"
)

# Diretorio local de destino
$LocalDestination = "C:\temp"

# Funcao para verificar conexao Azure
function Test-AzureConnection {
    try {
        $context = Get-AzContext
        if ($null -eq $context) {
            Write-Host "NAO CONECTADO ao Azure. Execute Connect-AzAccount primeiro." -ForegroundColor Red
            return $false
        }
        Write-Host "CONECTADO ao Azure: $($context.Account.Id)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "ERRO ao verificar conexao: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Funcao para criar diretorio local
function New-LocalDirectory {
    param([string]$Path)
    
    if (-not (Test-Path -Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-Host "DIRETORIO CRIADO: $Path" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "ERRO ao criar diretorio $Path : $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# Funcao para obter VMs de um Resource Group
function Get-VMsFromResourceGroup {
    param([string]$ResourceGroupName)
    
    try {
        Write-Host "BUSCANDO VMs no Resource Group: $ResourceGroupName" -ForegroundColor Yellow
        
        $rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
        if ($null -eq $rg) {
            Write-Host "Resource Group '$ResourceGroupName' NAO ENCONTRADO." -ForegroundColor Yellow
            return @()
        }
        
        $vms = Get-AzVM -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
        
        if ($vms.Count -eq 0) {
            Write-Host "NENHUMA VM encontrada no Resource Group '$ResourceGroupName'." -ForegroundColor Yellow
            return @()
        }
        
        Write-Host "ENCONTRADAS $($vms.Count) VM(s)" -ForegroundColor Green
        return $vms
    }
    catch {
        Write-Host "ERRO ao obter VMs: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Funcao para copiar pasta da VM usando robocopy
function Copy-VMFolder {
    param(
        [string]$ResourceGroupName,
        [string]$VMName,
        [string]$RemoteFolder,
        [string]$LocalVMPath
    )
    
    try {
        Write-Host "  COPIANDO pasta: $RemoteFolder" -ForegroundColor Cyan
        
        # Nome da pasta local baseado no caminho remoto
        $folderName = ($RemoteFolder -replace ":", "") -replace "\\", "_"
        $localFolderPath = Join-Path -Path $LocalVMPath -ChildPath $folderName
        
        # Cria diretorio local se nao existir
        New-LocalDirectory -Path $localFolderPath | Out-Null
        
        # Script para usar robocopy diretamente na VM
        $copyScript = @"
if (Test-Path -Path '$RemoteFolder') {
    Write-Output "PASTA_EXISTE:$RemoteFolder"
    
    # Cria pasta temporaria
    `$tempShare = "C:\temp_share_$(Get-Random)"
    New-Item -ItemType Directory -Path `$tempShare -Force | Out-Null
    
    # Copia usando robocopy
    robocopy '$RemoteFolder' `$tempShare /E /COPY:DAT /R:1 /W:1 /XD `$RECYCLE.BIN "System Volume Information"
    
    # Lista arquivos copiados
    `$files = Get-ChildItem -Path `$tempShare -Recurse -File
    Write-Output "ARQUIVOS_COPIADOS:`$(`$files.Count)"
    
    # Compacta para facilitar transferencia
    `$zipPath = "`$tempShare.zip"
    if (`$files.Count -gt 0) {
        Compress-Archive -Path `$tempShare -DestinationPath `$zipPath -Force
        `$zipSize = (Get-Item `$zipPath).Length
        Write-Output "ZIP_CRIADO:`$zipSize"
        
        # Converte para Base64
        `$zipBytes = [System.IO.File]::ReadAllBytes(`$zipPath)
        `$base64 = [Convert]::ToBase64String(`$zipBytes)
        Write-Output "BASE64_START"
        Write-Output `$base64
        Write-Output "BASE64_END"
    }
    
    # Limpa arquivos temporarios
    Remove-Item -Path `$tempShare -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path `$zipPath -Force -ErrorAction SilentlyContinue
}
else {
    Write-Output "PASTA_NAO_ENCONTRADA:$RemoteFolder"
}
"@
        
        # Executa o script na VM
        Write-Host "  EXECUTANDO copia na VM..." -ForegroundColor Magenta
        $result = Invoke-AzVMRunCommand -ResourceGroupName $ResourceGroupName -VMName $VMName -CommandId 'RunPowerShellScript' -ScriptString $copyScript -ErrorAction Stop
        
        $output = $result.Value[0].Message
        $lines = $output -split "`n"
        
        # Processa o resultado
        $base64Started = $false
        $base64Content = @()
        
        foreach ($line in $lines) {
            $line = $line.Trim()
            
            if ($line.StartsWith("PASTA_EXISTE:")) {
                Write-Host "  PASTA ENCONTRADA na VM" -ForegroundColor Green
            }
            elseif ($line.StartsWith("PASTA_NAO_ENCONTRADA:")) {
                Write-Host "  PASTA NAO ENCONTRADA na VM: $RemoteFolder" -ForegroundColor Yellow
                return
            }
            elseif ($line.StartsWith("ARQUIVOS_COPIADOS:")) {
                $fileCount = $line.Substring(18)
                Write-Host "  ARQUIVOS ENCONTRADOS: $fileCount" -ForegroundColor Gray
            }
            elseif ($line.StartsWith("ZIP_CRIADO:")) {
                $zipSize = $line.Substring(11)
                Write-Host "  ZIP CRIADO: $zipSize bytes" -ForegroundColor Gray
            }
            elseif ($line -eq "BASE64_START") {
                $base64Started = $true
                Write-Host "  RECEBENDO dados..." -ForegroundColor Gray
            }
            elseif ($line -eq "BASE64_END") {
                if ($base64Content.Count -gt 0) {
                    try {
                        # Reconstroi o arquivo ZIP
                        $base64String = $base64Content -join ""
                        $zipBytes = [Convert]::FromBase64String($base64String)
                        
                        # Salva o ZIP temporariamente
                        $tempZipPath = Join-Path -Path $env:TEMP -ChildPath "vm_folder_$(Get-Random).zip"
                        [System.IO.File]::WriteAllBytes($tempZipPath, $zipBytes)
                        
                        # Extrai o ZIP para a pasta local
                        Expand-Archive -Path $tempZipPath -DestinationPath $localFolderPath -Force
                        
                        # Remove ZIP temporario
                        Remove-Item -Path $tempZipPath -Force
                        
                        Write-Host "  PASTA COPIADA com sucesso para: $localFolderPath" -ForegroundColor Green
                        
                        # Mostra estatisticas
                        $localFiles = Get-ChildItem -Path $localFolderPath -Recurse -File -ErrorAction SilentlyContinue
                        Write-Host "  TOTAL DE ARQUIVOS salvos: $($localFiles.Count)" -ForegroundColor Cyan
                    }
                    catch {
                        Write-Host "  ERRO ao extrair pasta: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                $base64Started = $false
                $base64Content = @()
            }
            elseif ($base64Started -and $line.Length -gt 0) {
                $base64Content += $line
            }
        }
    }
    catch {
        Write-Host "  ERRO ao copiar pasta $RemoteFolder : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Funcao para processar uma VM
function Copy-VMFolders {
    param(
        [string]$ResourceGroupName,
        [string]$VMName
    )
    
    Write-Host "PROCESSANDO VM: $VMName" -ForegroundColor Green
    Write-Host "----------------------------------------" -ForegroundColor Green
    
    # Verifica se a VM esta rodando
    $vmStatus = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -Status
    $powerState = ($vmStatus.Statuses | Where-Object { $_.Code -like "PowerState/*" }).DisplayStatus
    
    if ($powerState -ne "VM running") {
        Write-Host "VM NAO ESTA RODANDO (Status: $powerState). Pulando." -ForegroundColor Yellow
        Write-Host ""
        return
    }
    
    # Cria diretorio local para a VM
    $vmLocalPath = Join-Path -Path $LocalDestination -ChildPath "$ResourceGroupName\$VMName"
    if (-not (New-LocalDirectory -Path $vmLocalPath)) {
        return
    }
    
    # Copia cada pasta
    foreach ($folder in $FoldersToSync) {
        Copy-VMFolder -ResourceGroupName $ResourceGroupName -VMName $VMName -RemoteFolder $folder -LocalVMPath $vmLocalPath
    }
    
    Write-Host "VM PROCESSADA com sucesso!" -ForegroundColor Green
    Write-Host ""
}

# Funcao principal
function Main {
    Write-Host "=== COPIANDO PASTAS DAS VMs AZURE ===" -ForegroundColor Magenta
    Write-Host ""
    
    # Verifica conexao com Azure
    if (-not (Test-AzureConnection)) {
        return
    }
    
    # Cria diretorio de destino
    if (-not (New-LocalDirectory -Path $LocalDestination)) {
        Write-Host "NAO FOI POSSIVEL criar diretorio de destino. Encerrando." -ForegroundColor Red
        return
    }
    
    Write-Host "CONFIGURACAO:" -ForegroundColor Cyan
    Write-Host "- Resource Groups: $($ResourceGroups.Count)" -ForegroundColor White
    Write-Host "- Pastas para copiar: $($FoldersToSync -join ', ')" -ForegroundColor White
    Write-Host "- Destino: $LocalDestination" -ForegroundColor White
    Write-Host ""
    
    $totalProcessed = 0
    
    foreach ($rgName in $ResourceGroups) {
        Write-Host "RESOURCE GROUP: $rgName" -ForegroundColor Blue
        Write-Host "================================================" -ForegroundColor Blue
        
        $vms = Get-VMsFromResourceGroup -ResourceGroupName $rgName
        
        if ($vms.Count -eq 0) {
            Write-Host "SEM VMs para processar." -ForegroundColor Yellow
            Write-Host ""
            continue
        }
        
        foreach ($vm in $vms) {
            Copy-VMFolders -ResourceGroupName $rgName -VMName $vm.Name
            $totalProcessed++
        }
    }
    
    Write-Host "=== COPIA FINALIZADA ===" -ForegroundColor Magenta
    Write-Host "- VMs processadas: $totalProcessed" -ForegroundColor White
    Write-Host "- Pastas salvas em: $LocalDestination" -ForegroundColor White
    Write-Host ""
    Write-Host "OPERACAO CONCLUIDA!" -ForegroundColor Green
}

# Executa o script
try {
    Main
}
catch {
    Write-Host "ERRO CRITICO: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")