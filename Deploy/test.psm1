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
        Verifica si existe el archivo de test
    #>
    
    [CmdletBinding()]
    param()
    
    Write-Host "`n=== VERIFICACION DE INSTALACION ===" -ForegroundColor Cyan
    
    # Verificar si existe el archivo de test
    $testFile = "C:\temp\test.txt"
    $isInstalled = Test-Path -Path $testFile
    
    # Verificar permisos de administrador
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $hasAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    # Verificar espacio en disco
    $drive = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    $freeSpaceGB = if ($drive) { [math]::Round($drive.Free / 1GB, 2) } else { 0 }
    $hasDiskSpace = $freeSpaceGB -ge 1
    
    # Mostrar resultados
    Write-Host "`nEstado de instalacion:" -ForegroundColor Yellow
    Write-Host "  Archivo test: $testFile" -ForegroundColor Cyan
    Write-Host "  Instalado: $isInstalled" -ForegroundColor $(if ($isInstalled) { "Green" } else { "Red" })
    
    Write-Host "`nPrerequisitos:" -ForegroundColor Yellow
    Write-Host "  Admin: $hasAdmin" -ForegroundColor $(if ($hasAdmin) { "Green" } else { "Red" })
    Write-Host "  Espacio: $freeSpaceGB GB" -ForegroundColor $(if ($hasDiskSpace) { "Green" } else { "Red" })
    
    return [PSCustomObject]@{
        IsInstalled    = $isInstalled
        TestFile       = $testFile
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
        Simula instalacion creando el archivo test.txt
    #>
    
    [CmdletBinding()]
    param(
        [string]$SetupExePath,
        [string]$ConfigXmlPath
    )
    
    Write-Host "`n=== INSTALACION ===" -ForegroundColor Cyan
    Write-Host "[SIMULACION - No se ejecutara instalacion real]" -ForegroundColor Magenta
    
    if (-not (Test-Path -Path $ConfigXmlPath)) {
        throw "No se encontro el archivo de configuracion"
    }
    
    Write-Host "`nComando: $SetupExePath /configure `"$ConfigXmlPath`"" -ForegroundColor Cyan
    
    Write-Host "`n[SIMULADO] Iniciando instalacion..." -ForegroundColor Magenta
    
    # Simular instalacion de 30 segundos con progreso
    $totalSeconds = 30
    $interval = 5
    for ($i = $interval; $i -le $totalSeconds; $i += $interval) {
        Start-Sleep -Seconds $interval
        $percentage = [math]::Round(($i / $totalSeconds) * 100)
        Write-Host "[SIMULADO] Progreso: $percentage% ($i/$totalSeconds segundos)" -ForegroundColor Magenta
    }
    
    # Crear el archivo de test
    $testFile = "C:\temp\test.txt"
    $testDir = Split-Path -Path $testFile -Parent
    
    if (-not (Test-Path -Path $testDir)) {
        New-Item -Path $testDir -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "Instalacion completada: $timestamp" | Out-File -FilePath $testFile -Encoding UTF8 -Force
    
    Write-Host "[SIMULADO] Instalacion completada" -ForegroundColor Magenta
    Write-Host "[OK] Archivo creado: $testFile" -ForegroundColor Green
    
    return [PSCustomObject]@{
        Success  = $true
        ExitCode = 0
        Duration = New-TimeSpan -Seconds $totalSeconds
        TestFile = $testFile
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
    
    if ($verification.IsInstalled) {
        Write-Host "[OK] Instalacion verificada correctamente" -ForegroundColor Green
        Write-Host "Archivo: $($verification.TestFile)" -ForegroundColor Cyan
    }
    else {
        Write-Warning "No se pudo verificar la instalacion"
    }
    
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
        VerificationSuccess = $verification.IsInstalled
        TestFile            = $verification.TestFile
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
        $isInstalled = Test-Installed
        
        if ($isInstalled -and -not $Force) {
            Write-Host "`n[OK] Ya esta instalado (archivo test.txt existe)" -ForegroundColor Green
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
            -ConfigXmlPath $preinstall.ConfigXmlPath
        
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
        Write-Host "  Archivo creado: $($install.TestFile)" -ForegroundColor Green
        Write-Host "  Duracion simulada: $($install.Duration.ToString('mm\:ss'))" -ForegroundColor Cyan
        
        return [PSCustomObject]@{
            Success           = $true
            IsInstalled       = $isInstalled
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

Function Get-DeployCnf {
    <#
    .SYNOPSIS
        Devuelve la configuracion por defecto para el despliegue de Test
    
    .DESCRIPTION
        Esta funcion proporciona los valores por defecto para N, Every y Message
        que seran utilizados por Start-GbDeploy si no se especifican manualmente.
    
    .OUTPUTS
        PSCustomObject con las propiedades N, Every y Message
    
    .EXAMPLE
        $config = Get-DeployCnf
        Start-GbDeploy -Name "test" -N $config.N -Every $config.Every -Message $config.Message
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()
    
    return [PSCustomObject]@{
        N       = 3
        Every   = 60
        Message = "Se va a actualizar Office a la version de 64-bit. Durante la actualizacion podras usar tu ordenador, pero no podras usar las aplicaciones de Office."
    }
}

# Las funciones se exportan automaticamente cuando se ejecuta con iex
# Si se importa como modulo, descomentar la siguiente linea:
# Export-ModuleMember -Function Test-Installed, Start-Preinstall, Start-Install, Start-PostInstall, Start-Deploy, Get-DeployCnf
