<#
.SYNOPSIS
    Modulo de TESTING para instalacion de Office 64-bit

.DESCRIPTION
    IMPORTANTE: Este modulo es SOLO para TESTING y NO INSTALA NADA REAL.
    
    Simula el flujo de instalacion de Office 64-bit sin ejecutar instalaciones reales.
    
    Uso remoto:
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/test.psm1') | iex ; Start-Deploy

.NOTES
    Version:        0.2.0
    Author:         Garikoitz Belarbide    
    Creation Date:  19/01/2026
    Purpose:        TESTING ONLY - Version simplificada

#>

Function Test-Installed {
    <#
    .SYNOPSIS
        Verifica Office instalado y prerequisitos
    #>
    
    [CmdletBinding()]
    param()
    
    Write-Host "`n=== VERIFICACION DE INSTALACION ===" -ForegroundColor Cyan
    
    # Verificar Office instalado
    $officeInstalled = $false
    $officeArch = $null
    $officeVersion = $null
    
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )
    
    foreach ($regPath in $regPaths) {
        if (Test-Path -Path $regPath) {
            $config = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
            if ($config -and $config.VersionToReport) {
                $officeInstalled = $true
                $officeVersion = $config.VersionToReport
                $officeArch = if ($config.Platform -eq "x64") { "64-bit" } else { "32-bit" }
                break
            }
        }
    }
    
    # Verificar permisos de administrador
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $hasAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    # Verificar espacio en disco
    $drive = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    $freeSpaceGB = if ($drive) { [math]::Round($drive.Free / 1GB, 2) } else { 0 }
    $hasDiskSpace = $freeSpaceGB -ge 10
    
    # Mostrar resultados
    Write-Host "`nEstado de Office:" -ForegroundColor Yellow
    Write-Host "  Instalado: $officeInstalled" -ForegroundColor $(if ($officeInstalled) { "Green" } else { "Red" })
    if ($officeInstalled) {
        Write-Host "  Arquitectura: $officeArch" -ForegroundColor $(if ($officeArch -eq "64-bit") { "Green" } else { "Yellow" })
        Write-Host "  Version: $officeVersion" -ForegroundColor Cyan
    }
    
    Write-Host "`nPrerequisitos:" -ForegroundColor Yellow
    Write-Host "  Admin: $hasAdmin" -ForegroundColor $(if ($hasAdmin) { "Green" } else { "Red" })
    Write-Host "  Espacio: $freeSpaceGB GB" -ForegroundColor $(if ($hasDiskSpace) { "Green" } else { "Red" })
    
    return [PSCustomObject]@{
        IsInstalled    = $officeInstalled
        Architecture   = $officeArch
        Version        = $officeVersion
        NeedsMigration = $officeInstalled -and $officeArch -eq "32-bit"
        HasAdminRights = $hasAdmin
        HasDiskSpace   = $hasDiskSpace
        CanProceed     = $hasAdmin -and $hasDiskSpace
    }
}

Function Start-Preinstall {
    <#
    .SYNOPSIS
        Simula preparacion de archivos
    #>
    
    [CmdletBinding()]
    param(
        [string]$InstallPath = "C:\Temp\OfficeTest"
    )
    
    Write-Host "`n=== PREPARACION ===" -ForegroundColor Cyan
    Write-Host "[SIMULACION - No se descargaran archivos reales]" -ForegroundColor Magenta
    
    # Crear directorio
    if (-not (Test-Path -Path $InstallPath)) {
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
        Write-Host "[OK] Directorio creado: $InstallPath" -ForegroundColor Green
    }
    
    # Crear configuracion XML
    $configXML = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" MigrateArch="TRUE" ForceAppShutdown="TRUE">
    <Product ID="O365BusinessRetail">
      <Language ID="es-es" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Teams" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Updates Enabled="TRUE" Channel="MonthlyEnterprise" />
  <RemoveMSI />
</Configuration>
"@
    
    $configPath = Join-Path -Path $InstallPath -ChildPath "configuration.xml"
    $configXML | Out-File -FilePath $configPath -Encoding UTF8 -Force
    Write-Host "[OK] Configuracion XML creada" -ForegroundColor Green
    
    $setupPath = Join-Path -Path $InstallPath -ChildPath "setup.exe"
    Write-Host "[SIMULADO] Setup.exe estaria en: $setupPath" -ForegroundColor Magenta
    
    return [PSCustomObject]@{
        Success       = $true
        InstallPath   = $InstallPath
        SetupExePath  = $setupPath
        ConfigXmlPath = $configPath
    }
}

Function Start-Install {
    <#
    .SYNOPSIS
        Simula instalacion
    #>
    
    [CmdletBinding()]
    param(
        [string]$SetupExePath,
        [string]$ConfigXmlPath,
        [bool]$NeedsMigration = $false
    )
    
    Write-Host "`n=== INSTALACION ===" -ForegroundColor Cyan
    Write-Host "[SIMULACION - No se ejecutara instalacion real]" -ForegroundColor Magenta
    
    if (-not (Test-Path -Path $ConfigXmlPath)) {
        throw "No se encontro el archivo de configuracion"
    }
    
    $tipo = if ($NeedsMigration) { "MIGRACION 32-bit -> 64-bit" } else { "INSTALACION NUEVA" }
    Write-Host "`nTipo: $tipo" -ForegroundColor Yellow
    Write-Host "Comando: $SetupExePath /configure `"$ConfigXmlPath`"" -ForegroundColor Cyan
    
    Write-Host "`n[SIMULADO] Instalacion completada" -ForegroundColor Magenta
    Start-Sleep -Seconds 2
    
    return [PSCustomObject]@{
        Success      = $true
        ExitCode     = 0
        Duration     = New-TimeSpan -Minutes 8 -Seconds 34
        WasMigration = $NeedsMigration
    }
}

Function Start-PostInstall {
    <#
    .SYNOPSIS
        Verifica y limpia
    #>
    
    [CmdletBinding()]
    param(
        [string]$InstallPath,
        [switch]$KeepFiles
    )
    
    Write-Host "`n=== VERIFICACION POST-INSTALACION ===" -ForegroundColor Cyan
    
    # Verificar estado actual
    $verification = Test-Installed
    
    # Limpiar archivos
    if (-not $KeepFiles -and (Test-Path -Path $InstallPath)) {
        try {
            Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Archivos temporales eliminados" -ForegroundColor Green
        }
        catch {
            Write-Warning "No se pudieron eliminar archivos: $_"
        }
    }
    
    return [PSCustomObject]@{
        VerificationSuccess   = $true
        InstalledVersion      = if ($verification.IsInstalled) { $verification.Version } else { "No instalado" }
        InstalledArchitecture = if ($verification.IsInstalled) { $verification.Architecture } else { "N/A" }
    }
}

Function Start-Deploy {
    <#
    .SYNOPSIS
        Orquesta el flujo completo de test
    #>
    
    [CmdletBinding()]
    param(
        [string]$InstallPath = "C:\Temp\OfficeTest",
        [switch]$Force,
        [switch]$KeepTempFiles
    )
    
    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Host "  TEST: DESPLIEGUE DE MICROSOFT OFFICE 64-BIT" -ForegroundColor Cyan
    Write-Host "  MODO: SIMULACION (No se instalara nada real)" -ForegroundColor Magenta
    Write-Host "================================================================" -ForegroundColor Cyan
    
    try {
        # FASE 1: Verificacion
        Write-Host "`n--- FASE 1: VERIFICACION ---" -ForegroundColor Cyan
        $testResult = Test-Installed
        
        if (-not $testResult.CanProceed) {
            if (-not $testResult.HasAdminRights) {
                throw "Se requieren permisos de administrador"
            }
            if (-not $testResult.HasDiskSpace) {
                throw "Espacio en disco insuficiente"
            }
        }
        
        if ($testResult.IsInstalled -and $testResult.Architecture -eq "64-bit" -and -not $Force) {
            Write-Host "`n[OK] Office 64-bit ya esta instalado" -ForegroundColor Green
            Write-Host "Version: $($testResult.Version)" -ForegroundColor Cyan
            Write-Host "Use -Force para simular reinstalacion" -ForegroundColor Yellow
            return [PSCustomObject]@{ Success = $true; AlreadyInstalled = $true }
        }
        
        # FASE 2: Preparacion
        Write-Host "`n--- FASE 2: PREPARACION ---" -ForegroundColor Cyan
        $preinstall = Start-Preinstall -InstallPath $InstallPath
        
        if (-not $preinstall.Success) {
            throw "Error en preparacion"
        }
        
        # FASE 3: Instalacion
        Write-Host "`n--- FASE 3: INSTALACION ---" -ForegroundColor Cyan
        $install = Start-Install -SetupExePath $preinstall.SetupExePath `
            -ConfigXmlPath $preinstall.ConfigXmlPath `
            -NeedsMigration $testResult.NeedsMigration
        
        if (-not $install.Success) {
            throw "Error en instalacion"
        }
        
        # FASE 4: Verificacion
        Write-Host "`n--- FASE 4: VERIFICACION ---" -ForegroundColor Cyan
        $postInstall = Start-PostInstall -InstallPath $InstallPath -KeepFiles:$KeepTempFiles
        
        # Resumen
        Write-Host "`n================================================================" -ForegroundColor Green
        Write-Host "  [OK] SIMULACION COMPLETADA EXITOSAMENTE" -ForegroundColor Green
        Write-Host "================================================================" -ForegroundColor Green
        Write-Host "`nResumen:" -ForegroundColor Cyan
        Write-Host "  Version actual: $($postInstall.InstalledVersion)" -ForegroundColor Green
        Write-Host "  Arquitectura: $($postInstall.InstalledArchitecture)" -ForegroundColor Green
        Write-Host "  Duracion simulada: $($install.Duration.ToString('mm\:ss'))" -ForegroundColor Cyan
        Write-Host "  Tipo: $(if ($testResult.NeedsMigration) { 'Migracion 32->64' } else { 'Instalacion nueva' })" -ForegroundColor Cyan
        
        return [PSCustomObject]@{
            Success           = $true
            TestResult        = $testResult
            PreinstallResult  = $preinstall
            InstallResult     = $install
            PostInstallResult = $postInstall
        }
    }
    catch {
        Write-Host "`n================================================================" -ForegroundColor Red
        Write-Host "  [ERROR] ERROR EN EL DESPLIEGUE" -ForegroundColor Red
        Write-Host "================================================================" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        
        return [PSCustomObject]@{
            Success      = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Las funciones se exportan automaticamente cuando se ejecuta con iex
# Si se importa como modulo, descomentar la siguiente linea:
# Export-ModuleMember -Function Test-Installed, Start-Preinstall, Start-Install, Start-PostInstall, Start-Deploy
