function Show-LimpiaMenu {
    Clear-Host
    Write-Host "=== Módulo de Limpieza ===" -ForegroundColor Cyan
    Write-Host "1. Utilidad de limpieza de disco con todas las opciones"
    Write-Host "2. Paquetes DISM (puede llevar un tiempo)"
    Write-Host "3. Windows Temp"
    Write-Host "4. Updates (Servicio Windows Update)"
    Write-Host "5. TODO (Ejecutar opciones 1, 2, 3 y 4)"
    Write-Host "6. Vaciar todas las papeleras"
    Write-Host "7. Perfiles de usuario"
    Write-Host ""
    Write-Host "0. Exit"
}

function Clear-WindowsTemp {
    Write-Host "Iniciando la limpieza de la carpeta Temp de Windows ..." -ForegroundColor Yellow
    Remove-Item "$env:SystemRoot\TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:LOCALAPPDATA\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
}

function Clear-Dism {
    Write-Host "Iniciando la limpieza de DISM (puede llevar un tiempo) ..." -ForegroundColor Yellow
    if ([Environment]::OSVersion.Version -lt (New-Object 'Version' 6, 2)) { 
        Dism.exe /online /Cleanup-Image /SpSuperseded
    }
    else { 
        Dism.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase
    }
}

function Clear-WinClean {    
    Write-Host "Configurando Cleanmgr con todas las opciones..." -ForegroundColor Yellow
    $CleanMgrKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    $items = Get-ChildItem -Path $CleanMgrKey
    foreach ($item in $items) {
        Set-ItemProperty -Path $item.PSPath -Name StateFlags0001 -Type DWORD -Value 2 -ErrorAction SilentlyContinue
    }
    
    Write-Host "Iniciando Cleanmgr (puede tomar un tiempo) ..." -ForegroundColor Yellow
    $Process = Start-Process -FilePath "$env:systemroot\system32\cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait -PassThru
    Write-Host "Proceso terminado con código de salida [$($Process.ExitCode)]."
}

function Clear-RecycleBin {
    Write-Host "Iniciando la limpieza de las papeleras..." -ForegroundColor Yellow
    # Usar el comando nativo de PowerShell si está disponible (PS 5.0+)
    if (Get-Command Clear-RecycleBin -ErrorAction SilentlyContinue) {
        Clear-RecycleBin -Confirm:$false -ErrorAction SilentlyContinue
    }
    else {
        if (Test-Path "C:\`$Recycle.Bin") {
            Start-Process -FilePath "cmd" -ArgumentList '/c "rd /s /q c:\$Recycle.Bin"' -Wait -WindowStyle Hidden
        }
    }
    Write-Host "Papeleras vaciadas." -ForegroundColor Green
}

function Clear-WindowsUpdates {
    Write-Host "Iniciando limpieza de Windows Updates..." -ForegroundColor Yellow
    Write-Host "Deteniendo servicio Windows Update..."
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Write-Host "Borrando carpeta SoftwareDistribution..."
    Remove-Item "$env:SystemRoot\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Iniciando servicio Windows Update..."
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
    Write-Host "Limpieza de updates finalizada." -ForegroundColor Green
}

function Remove-UserProfiles {
    $ProfilesList = Get-CimInstance -Class Win32_UserProfile | Where-Object { -not $_.Special }
    
    if ($ProfilesList.Count -eq 0) {
        Write-Host "No se encontraron perfiles de usuario eliminables." -ForegroundColor Yellow
        pause
        return
    }

    do {
        Clear-Host
        Write-Host "=== Perfiles de Usuario ===" -ForegroundColor Cyan
        $i = 1
        foreach ($UserProfile in $ProfilesList) {
            Write-Host "$i - $($UserProfile.LocalPath)"
            $i++
        }
        Write-Host "0 - SALIR"
        Write-Host ""
        
        $userSelection = Read-Host "Selecciona el número del perfil a borrar (o varios separados por coma)"
        if ($userSelection -eq "0" -or -not $userSelection) { break }

        $selectedIndices = $userSelection.Split(",")
        foreach ($idx in $selectedIndices) {
            $index = [int]$idx.Trim() - 1
            if ($index -ge 0 -and $index -lt $ProfilesList.Count) {
                $target = $ProfilesList[$index]
                Write-Host "Borrando perfil: $($target.LocalPath)..." -ForegroundColor Red
                try {
                    $target | Remove-CimInstance
                    Write-Host "Perfil eliminado correctamente." -ForegroundColor Green
                }
                catch {
                    Write-Host "Error al eliminar el perfil: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        pause
        $ProfilesList = Get-CimInstance -Class Win32_UserProfile | Where-Object { -not $_.Special }
    } while ($ProfilesList.Count -gt 0)
}

function Start-GbLimpia {
    <#
    .SYNOPSIS
        Inicia el proceso de limpieza del sistema.
    #>
    
    # Comprobar privilegios de administrador
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "ESTE SCRIPT REQUIERE PRIVILEGIOS DE ADMINISTRADOR." -ForegroundColor Red
        Write-Host "Por favor, reinicie como administrador."
        pause
        return
    }

    do {
        Show-LimpiaMenu
        $opt = Read-Host "Selecciona una opción"
        
        if ($opt -eq "0" -or -not $opt) { break }
        
        # Calcular espacio antes
        $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
        $before = $drive.FreeSpace / 1GB
        Write-Host ("Espacio libre actual: {0:N2} GB" -f $before) -ForegroundColor Cyan

        switch ($opt) {
            "1" { Clear-WinClean }
            "2" { Clear-Dism }
            "3" { Clear-WindowsTemp }
            "4" { Clear-WindowsUpdates }
            "5" { 
                Clear-WinClean
                Clear-Dism
                Clear-WindowsTemp
                Clear-WindowsUpdates
            }
            "6" { Clear-RecycleBin }
            "7" { Remove-UserProfiles }
        }

        # Calcular espacio después
        $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
        $after = $drive.FreeSpace / 1GB
        Write-Host ("Espacio libre ahora: {0:N2} GB" -f $after) -ForegroundColor Green
        Write-Host ("Se han liberado: {0:N2} GB" -f ($after - $before)) -ForegroundColor Cyan
        
        pause
    } while ($true)
}
