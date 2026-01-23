function Write-GbLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "VERBOSE", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $logDir = "C:\temp"
    $logFile = Join-Path $logDir "gbdeploy.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    try {
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $logMessage | Out-File -FilePath $logFile -Append -Encoding UTF8
    }
    catch {
        # Fallback silencioso si no se puede escribir en el log
    }

    # Mostrar en consola con colores
    switch ($Level) {
        "WARNING" { Write-Warning $Message }
        "ERROR" { Write-Error $Message }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "VERBOSE" { Write-Verbose $Message }
        default { Write-Host $Message -ForegroundColor Cyan }
    }
}


Function Test-Installed {
    <#
    .SYNOPSIS
        Verifica si existe el archivo de test
    #>
    
    [CmdletBinding()]
    param()
    
    Write-GbLog -Message "=== VERIFICACION DE INSTALACION ===" -Level "INFO"
    Start-Sleep -Seconds 10
    
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
    Write-GbLog -Message "Estado de instalacion:" -Level "INFO"
    Write-GbLog -Message "  Archivo test: $testFile" -Level "INFO"
    Write-GbLog -Message "  Instalado: $isInstalled" -Level $(if ($isInstalled) { "SUCCESS" } else { "WARNING" })
    
    Write-GbLog -Message "Prerequisitos:" -Level "INFO"
    Write-GbLog -Message "  Admin: $hasAdmin" -Level $(if ($hasAdmin) { "SUCCESS" } else { "WARNING" })
    Write-GbLog -Message "  Espacio: $freeSpaceGB GB" -Level $(if ($hasDiskSpace) { "SUCCESS" } else { "WARNING" })
    
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
    
    Write-GbLog -Message "=== PREPARACION ===" -Level "INFO"
    Write-GbLog -Message "[SIMULACION - No se descargaran archivos reales]" -Level "VERBOSE"
    Start-Sleep -Seconds 10
    
    # Crear directorio
    if (-not (Test-Path -Path $InstallPath)) {
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
        Write-GbLog -Message "[OK] Directorio creado: $InstallPath" -Level "SUCCESS"
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
    Write-GbLog -Message "[OK] Configuracion XML creada" -Level "SUCCESS"
    
    $setupPath = Join-Path -Path $InstallPath -ChildPath "setup.exe"
    Write-GbLog -Message "[SIMULADO] Setup.exe estaria en: $setupPath" -Level "VERBOSE"
    
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
    
    Write-GbLog -Message "=== INSTALACION ===" -Level "INFO"
    Write-GbLog -Message "[SIMULACION - No se ejecutara instalacion real]" -Level "VERBOSE"
    
    if (-not (Test-Path -Path $ConfigXmlPath)) {
        throw "No se encontro el archivo de configuracion"
    }
    
    Write-GbLog -Message "Comando: $SetupExePath /configure `"$ConfigXmlPath`"" -Level "INFO"
    
    Write-GbLog -Message "[SIMULADO] Iniciando instalacion..." -Level "VERBOSE"
    
    # Simular instalacion de 30 segundos con progreso
    $totalSeconds = 30
    $interval = 5
    for ($i = $interval; $i -le $totalSeconds; $i += $interval) {
        Start-Sleep -Seconds $interval
        $percentage = [math]::Round(($i / $totalSeconds) * 100)
        Write-GbLog -Message "[SIMULADO] Progreso: $percentage% ($i/$totalSeconds segundos)" -Level "VERBOSE"
    }
    
    # Crear el archivo de test
    $testFile = "C:\temp\test.txt"
    $testDir = Split-Path -Path $testFile -Parent
    
    if (-not (Test-Path -Path $testDir)) {
        New-Item -Path $testDir -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "Instalacion completada: $timestamp" | Out-File -FilePath $testFile -Encoding UTF8 -Force
    
    Write-GbLog -Message "[SIMULADO] Instalacion completada" -Level "SUCCESS"
    Write-GbLog -Message "[OK] Archivo creado: $testFile" -Level "SUCCESS"
    
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
    
    Write-GbLog -Message "=== VERIFICACION POST-INSTALACION ===" -Level "INFO"
    Start-Sleep -Seconds 10
    
    # Verificar estado actual
    $verification = Test-Installed
    
    if ($verification.IsInstalled) {
        Write-GbLog -Message "[OK] Instalacion verificada correctamente" -Level "SUCCESS"
        Write-GbLog -Message "Archivo: $($verification.TestFile)" -Level "INFO"
    }
    else {
        Write-GbLog -Message "No se pudo verificar la instalacion" -Level "WARNING"
    }
    
    # Limpiar archivos
    if (-not $KeepFiles -and (Test-Path -Path $InstallPath)) {
        try {
            Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction Stop
            Write-GbLog -Message "[OK] Archivos temporales eliminados" -Level "SUCCESS"
        }
        catch {
            Write-GbLog -Message "No se pudieron eliminar archivos: $_" -Level "WARNING"
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
    
    Write-GbLog -Message "================================================================" -Level "INFO"
    Write-GbLog -Message "  TEST: DESPLIEGUE DE MICROSOFT OFFICE 64-BIT" -Level "INFO"
    Write-GbLog -Message "  MODO: SIMULACION (No se instalara nada real)" -Level "VERBOSE"
    Write-GbLog -Message "================================================================" -Level "INFO"
    
    try {
        # FASE 1: Verificacion
        Write-GbLog -Message "--- FASE 1: VERIFICACION ---" -Level "INFO"
        $isInstalled = Test-Installed
        
        if ($isInstalled -and -not $Force) {
            Write-GbLog -Message "[OK] Ya esta instalado (archivo test.txt existe)" -Level "SUCCESS"
            Write-GbLog -Message "Use -Force para simular reinstalacion" -Level "WARNING"
            return [PSCustomObject]@{ Success = $true; AlreadyInstalled = $true }
        }
        
        # FASE 2: Preparacion
        Write-GbLog -Message "--- FASE 2: PREPARACION ---" -Level "INFO"
        $preinstall = Start-Preinstall -InstallPath $InstallPath
        
        if (-not $preinstall.Success) {
            throw "Error en preparacion"
        }
        
        # FASE 3: Instalacion
        Write-GbLog -Message "--- FASE 3: INSTALACION ---" -Level "INFO"
        $install = Start-Install -SetupExePath $preinstall.SetupExePath `
            -ConfigXmlPath $preinstall.ConfigXmlPath
        
        if (-not $install.Success) {
            throw "Error en instalacion"
        }
        
        # FASE 4: Verificacion
        Write-GbLog -Message "--- FASE 4: VERIFICACION ---" -Level "INFO"
        $postInstall = Start-PostInstall -InstallPath $InstallPath -KeepFiles:$KeepTempFiles
        
        # Resumen
        Write-GbLog -Message "================================================================" -Level "SUCCESS"
        Write-GbLog -Message "  [OK] SIMULACION COMPLETADA EXITOSAMENTE" -Level "SUCCESS"
        Write-GbLog -Message "================================================================" -Level "SUCCESS"
        Write-GbLog -Message "Resumen:" -Level "INFO"
        Write-GbLog -Message "  Archivo creado: $($install.TestFile)" -Level "SUCCESS"
        Write-GbLog -Message "  Duracion simulada: $($install.Duration.ToString('mm\:ss'))" -Level "INFO"
        
        return [PSCustomObject]@{
            Success           = $true
            IsInstalled       = $isInstalled
            PreinstallResult  = $preinstall
            InstallResult     = $install
            PostInstallResult = $postInstall
        }
    }
    catch {
        Write-GbLog -Message "================================================================" -Level "ERROR"
        Write-GbLog -Message "  [ERROR] ERROR EN EL DESPLIEGUE" -Level "ERROR"
        Write-GbLog -Message "================================================================" -Level "ERROR"
        Write-GbLog -Message "Error: $_" -Level "ERROR"
        
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
        Every   = 2
        Message = "Se requiere actualizar Office a la version de 64-bit. Durante la actualizacion podras usar tu ordenador, pero no podras usar las aplicaciones de Office (Word, Excel, Outlook, etc.). GUARDA y CIERRA todos los documentos antes de iniciar, ya que el instalador forzara el cierre de Office."
    }
}

# Las funciones se exportan automaticamente cuando se ejecuta con iex
# Si se importa como modulo, descomentar la siguiente linea:
# Export-ModuleMember -Function Test-Installed, Start-Preinstall, Start-Install, Start-PostInstall, Start-Deploy, Get-DeployCnf
