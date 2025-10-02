# Copy VM folders by fetching each file in a separate RunCommand call (avoids stdout size/truncation)
# Copies ALL files from C:\Windows\Panther and C:\Windows\OEM to local C:\temp preserving structure.
# Requirements: Connected to Azure (Connect-AzAccount). Az.Compute module installed.

param(
    [string[]]$ResourceGroups = @(
        'windows-2016-noncvm-kakothari-1001113343',
        'windows-2016-datacenter-gen2-noncvm-kakothari-1001113436',
        'windows-2016-datacenter-core-noncvm-kakothari-1001113524',
        'windows-2016-datacenter-core-g2-noncvm-kakothari-1001113614',
        'windows-2019-datacenter-gen2-noncvm-kakothari-1001113705',
        'windows-2019-datacenter-core-noncvm-kakothari-1001113802',
        'windows-2019-datacenter-core-g2-noncvm-kakothari-1001113859',
        'windows-2022-datacenter-gen1-noncvm-kakothari-1001113944',
        'windows-2022-datacenter-gen2-noncvm-kakothari-1001114051',
        'windows-2022-datacenter-azure-core-g2-noncvm-kakothari-1001114150',
        'windows-2022-datacenter-azure-gen2-noncvm-kakothari-1001114247',
        'windows-2022-datacenter-servercore-gen1-noncvm-kakothari-1001114344',
        'windows-2022-datacenter-servercore-gen2-noncvm-kakothari-1001114433',
        'windows-2025-datacenter-azure-edition-core-gen2-noncvm-kakothari-1001114519',
        'windows-2025-datacenter-azure-edition-gen2-noncvm-kakothari-1001114616',
        'windows-2025-datacenter-servercore-gen1-noncvm-kakothari-1001114720',
        'windows-2025-datacenter-servercore-gen2-noncvm-kakothari-1001114817',
        'windows-2025-datacenter-gen1-noncvm-kakothari-1001114909',
        'windows-2025-datacenter-gen2-noncvm-kakothari-1001115015',
        'windows-10-enterprise-22h2-gen1-noncvm-kakothari-1001115116',
        'windows-10-enterprise-22h2-gen2-noncvm-kakothari-1001115346',
        'windows-11-enterprise-22h2-gen2-noncvm-kakothari-1001115603',
        'windows-11-enterprise-23h2-gen2-noncvm-kakothari-1001115853',
        'windows-11-enterprise-24h2-gen2-noncvm-kakothari-1001120147'
    ),
    [string[]]$FoldersToCopy = @('C:\Windows\Panther','C:\Windows\OEM'),
    [string]$LocalRoot = 'C:\temp',
    [int]$MaxFileMB = 50
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

function New-LocalDir {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-VMFilesList {
    param(
        [string]$ResourceGroup,
        [string]$VMName,
        [string]$RemoteFolder
    )

                $script = @'
param([string]$folder)
$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
Write-Output ("__DBG__:FOLDER={0}" -f $folder)
Write-Output ("__DBG__:EXISTS={0}" -f (Test-Path $folder))
if (Test-Path $folder) {
    $list = @()
    try { $list = Get-ChildItem -Path $folder -Recurse -File -Force } catch { }
    if (-not $list -or $list.Count -eq 0) {
        try { $list = Get-ChildItem -Path $folder -Recurse -Force | Where-Object { -not $_.PSIsContainer } } catch { }
        Write-Output ("__DBG__:FALLBACK_NO_RECURSE_COUNT={0}" -f ($list | Measure-Object).Count)
    } else {
        Write-Output ("__DBG__:RECURSE_COUNT={0}" -f ($list | Measure-Object).Count)
    }
    $list | ForEach-Object {
        try { Write-Output ("__FILE__:{0}|{1}" -f $_.FullName, $_.Length) } catch { }
    }
} else {
    Write-Output ("__FOLDER_NOT_FOUND__|{0}" -f $folder)
}
'@

    try {
    $res = Invoke-AzVMRunCommand -ResourceGroupName $ResourceGroup -VMName $VMName -CommandId 'RunPowerShellScript' -ScriptString $script -Parameter @{"folder"=$RemoteFolder}
    $msg = $res.Value[0].Message
    $lines = ($msg -split "`n") | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 0 }
    foreach ($d in $lines) { if ($d -like '__DBG__:*') { Write-Host ("      [VMDBG] {0}" -f $d) -ForegroundColor DarkGray } }
        $files = @()
        foreach ($line in $lines) {
            if ($line -like '__FOLDER_NOT_FOUND__*') { continue }
            if (-not ($line -like '__FILE__:*')) { continue }
            $payload = $line.Substring(9)
            $parts = $payload -split '\|',2
            if ($parts.Count -eq 2) {
                $full = $parts[0]
                $len = [int64]($parts[1])
                $files += [pscustomobject]@{ FullPath=$full; Length=$len }
            }
        }
        return $files
    }
    catch {
        Write-Host "    ERRO ao listar: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

function Get-VMFileBytes {
    param(
        [string]$ResourceGroup,
        [string]$VMName,
        [string]$RemoteFile
    )

                $script = @'
param([string]$path)
$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
if (Test-Path $path) {
    try {
        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            Write-Output "__HEX_BEGIN__"
            $buf = New-Object byte[] 65536
            while (($read = $fs.Read($buf, 0, $buf.Length)) -gt 0) {
                if ($read -eq $buf.Length) {
                    $chunk = $buf
                } else {
                    $chunk = New-Object byte[] $read
                    [Array]::Copy($buf, $chunk, $read)
                }
                $hex = -join ( $chunk | ForEach-Object { $_.ToString('X2') } )
                Write-Output ("__HEX__:{0}" -f $hex)
            }
            Write-Output "__HEX_END__"
        } finally {
            $fs.Close()
            $fs.Dispose()
        }
    } catch {
        Write-Output ("__ERROR__:{0}|{1}" -f $path, $_.Exception.Message)
    }
} else {
    Write-Output ("__NOT_FOUND__:{0}" -f $path)
}
'@

        $res = Invoke-AzVMRunCommand -ResourceGroupName $ResourceGroup -VMName $VMName -CommandId 'RunPowerShellScript' -ScriptString $script -Parameter @{"path"=$RemoteFile}
        $allMsgs = ($res.Value | ForEach-Object { $_.Message }) -join "`n"
        $tryCount = 0
        while ($true) {
            $tryCount++
            $msg = $allMsgs
            $lines = ($msg -split "`n") | ForEach-Object { $_.Trim() }
            $collect = $false
            $hexChunks = @()
            foreach ($l in $lines) {
                if ($l -eq '__HEX_BEGIN__') { $collect = $true; continue }
                if ($l -eq '__HEX_END__')   { break }
                if ($collect -and $l.StartsWith('__HEX__:')) {
                    $hexChunks += $l.Substring(8)
                }
                if ($l -like '__ERROR__*') { throw "VM read error: $l" }
                if ($l -like '__NOT_FOUND__*') { throw "VM file not found: $RemoteFile" }
            }
            if ($hexChunks.Count -gt 0) {
                $hex = ($hexChunks -join '')
                # decode hex to bytes
                $len = $hex.Length / 2
                $bytes = New-Object byte[] $len
                for ($i=0; $i -lt $len; $i++) {
                    $bytes[$i] = [Convert]::ToByte($hex.Substring($i*2,2),16)
                }
                return $bytes
            }
            if ($tryCount -ge 2) { throw "Empty content for $RemoteFile" }
            Start-Sleep -Milliseconds 200
        }
}

# Main
$ctx = Get-AzContext
if (-not $ctx) { Write-Host '❌ Não conectado ao Azure. Use Connect-AzAccount.' -ForegroundColor Red; exit 1 }

New-LocalDir -Path $LocalRoot

$totalSaved = 0

foreach ($rg in $ResourceGroups) {
    Write-Host "=== $rg ===" -ForegroundColor Blue
    $vms = @( Get-AzVM -ResourceGroupName $rg -ErrorAction SilentlyContinue )
    if ($vms.Count -eq 0) { Write-Host '  (sem VMs)' -ForegroundColor Yellow; continue }

    foreach ($vm in $vms) {
        $vmName = $vm.Name
        Write-Host "  VM: $vmName" -ForegroundColor Cyan
        $status = Get-AzVM -ResourceGroupName $rg -Name $vmName -Status -ErrorAction SilentlyContinue
        $power = ($status.Statuses | Where-Object { $_.Code -like 'PowerState/*' }).DisplayStatus
        if ($power -ne 'VM running') { Write-Host "   ⚠️ VM não está rodando: $power" -ForegroundColor Yellow; continue }

    $vmRoot = Join-Path $LocalRoot "$rg\$vmName"
    New-LocalDir -Path $vmRoot

        foreach ($remoteFolder in $FoldersToCopy) {
            $folderLeaf = Split-Path $remoteFolder -Leaf
            $destFolder = Join-Path $vmRoot $folderLeaf
            # fresh folder
            if (Test-Path $destFolder) { Remove-Item $destFolder -Recurse -Force -ErrorAction SilentlyContinue }
            New-LocalDir -Path $destFolder

            Write-Host "    Pasta: $remoteFolder" -ForegroundColor White
            $files = Get-VMFilesList -ResourceGroup $rg -VMName $vmName -RemoteFolder $remoteFolder
            if ($files.Count -eq 0) { Write-Host '      (nenhum arquivo encontrado)' -ForegroundColor Yellow; continue }
            Write-Host "      Encontrados: $($files.Count) arquivo(s)" -ForegroundColor Green

            foreach ($f in $files) {
                try {
                    $lengthMB = [math]::Round($f.Length/1MB,2)
                    if ($MaxFileMB -gt 0 -and $f.Length -gt ($MaxFileMB*1MB)) {
                        Write-Host "      ⏭️ pulando >$MaxFileMB MB: $($f.FullPath) ($lengthMB MB)" -ForegroundColor Yellow
                        continue
                    }
                    # compute relative path
                    $rel = $f.FullPath.Substring($remoteFolder.Length).TrimStart('\\')
                    $target = Join-Path $destFolder $rel
                    $targetDir = Split-Path $target -Parent
                    New-LocalDir -Path $targetDir

                    if ($f.Length -eq 0) {
                        # create empty file locally
                        [System.IO.File]::WriteAllBytes($target, @())
                        $totalSaved++
                        Write-Host "        ✓ $rel (0 bytes)" -ForegroundColor Gray
                    } else {
                        # fetch bytes and write
                        $bytes = Get-VMFileBytes -ResourceGroup $rg -VMName $vmName -RemoteFile $f.FullPath
                        [System.IO.File]::WriteAllBytes($target, $bytes)
                        $totalSaved++
                        Write-Host "        ✓ $rel ($($bytes.Length) bytes)" -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host "        ❌ Falha em $($f.FullPath): $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
    }
}

Write-Host ""; Write-Host "✅ Concluído. Total de arquivos salvos: $totalSaved" -ForegroundColor Green
