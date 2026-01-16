<#
.SYNOPSIS
    Módulo de TESTING/SIMULACIÓN para instalación de Office 64-bit

.DESCRIPTION
    ⚠️ IMPORTANTE: Este módulo es SOLO para TESTING y NO INSTALA NADA REAL.
    
    Simula todo el flujo de instalación de Office 64-bit sin ejecutar instalaciones reales:
    - Verifica prerequisitos del sistema
    - Simula descarga de archivos ODT y Office
    - Simula la ejecución de setup.exe
    - Verifica el estado actual del sistema
    
    Útil para:
    - Probar el flujo de trabajo antes de ejecutar la instalación real
    - Verificar prerequisitos sin modificar el sistema
    - Validar configuración XML y parámetros
    - Entrenar/demostrar el proceso de instalación
    
    Uso remoto:
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/test.psm1') | iex ; Start-Deploy

.NOTES
    Version:        0.1.0
    Author:         Garikoitz Belarbide    
    Creation Date:  16/01/2026
    Purpose:        TESTING ONLY - No realiza instalaciones reales

#>

#region [Functions]-------------------------------------------------------------

Function Test-Installed {
    <#
    .SYNOPSIS
        Verifica Office instalado, arquitectura y prerequisitos
    
    .DESCRIPTION
        Comprueba el estado actual de Office en el sistema, incluyendo:
        - Si Office está instalado
        - Arquitectura (32-bit o 64-bit)
        - Versión instalada
        - Prerequisitos del sistema (permisos, espacio en disco)
        - Si necesita migración de 32-bit a 64-bit
    
    .OUTPUTS
        PSCustomObject con información detallada del estado de Office y prerequisitos
    
    .EXAMPLE
        $status = Test-Installed
        if ($status.IsInstalled -and $status.Architecture -eq "64-bit") {
            Write-Host "Office 64-bit ya instalado"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()
    
    Write-Host "`n=== TEST: VERIFICACIÓN DE INSTALACIÓN ===" -ForegroundColor Cyan
    
    # Función auxiliar para verificar Office instalado
    function Get-OfficeInfo {
        $result = [PSCustomObject]@{
            IsInstalled  = $false
            Architecture = $null
            Version      = $null
        }
        
        # Comprobar en el registro para ClickToRun (Office 365/2016+)
        $registryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
        )
        
        foreach ($regPath in $registryPaths) {
            if (Test-Path -Path $regPath) {
                $config = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                
                if ($config -and $config.VersionToReport) {
                    $platform = $config.Platform
                    
                    $result.IsInstalled = $true
                    $result.Version = $config.VersionToReport
                    
                    if ($platform -eq "x64") {
                        $result.Architecture = "64-bit"
                    }
                    elseif ($platform -eq "x86") {
                        $result.Architecture = "32-bit"
                    }
                    else {
                        # Fallback: determinar por la ruta del registro
                        if ($regPath -like "*WOW6432Node*") {
                            $result.Architecture = "32-bit"
                        }
                        else {
                            $result.Architecture = "64-bit"
                        }
                    }
                    
                    return $result
                }
            }
        }
        
        # Comprobar instalaciones MSI antiguas (Office 2013 y anteriores)
        $msiPaths = @(
            @{Path = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot"; Arch = "64-bit" },
            @{Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot"; Arch = "32-bit" },
            @{Path = "HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot"; Arch = "64-bit" },
            @{Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot"; Arch = "32-bit" }
        )
        
        foreach ($msiInfo in $msiPaths) {
            if (Test-Path -Path $msiInfo.Path) {
                $installPath = (Get-ItemProperty -Path $msiInfo.Path -Name "Path" -ErrorAction SilentlyContinue).Path
                if ($installPath -and (Test-Path -Path $installPath)) {
                    $result.IsInstalled = $true
                    $result.Architecture = $msiInfo.Arch
                    return $result
                }
            }
        }
        
        return $result
    }
    
    # Obtener información de Office instalado
    $officeInfo = Get-OfficeInfo
    
    # Verificar permisos de administrador
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $hasAdminRights = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    # Verificar espacio en disco (unidad C:)
    $drive = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    $freeSpaceGB = if ($drive) { [math]::Round($drive.Free / 1GB, 2) } else { 0 }
    $minDiskSpaceGB = 10
    $hasDiskSpace = $freeSpaceGB -ge $minDiskSpaceGB
    
    # Determinar si necesita migración
    $needsMigration = $officeInfo.IsInstalled -and $officeInfo.Architecture -eq "32-bit"
    
    # Crear objeto de resultado
    $result = [PSCustomObject]@{
        IsInstalled    = $officeInfo.IsInstalled
        Architecture   = $officeInfo.Architecture
        Version        = $officeInfo.Version
        NeedsMigration = $needsMigration
        Prerequisites  = [PSCustomObject]@{
            HasAdminRights = $hasAdminRights
            HasDiskSpace   = $hasDiskSpace
            FreeSpaceGB    = $freeSpaceGB
            MinDiskSpaceGB = $minDiskSpaceGB
        }
        CanProceed     = $hasAdminRights -and $hasDiskSpace
    }
    
    # Mostrar resultados
    Write-Host "`nEstado de Office:" -ForegroundColor Yellow
    Write-Host "  Instalado: " -NoNewline; Write-Host $result.IsInstalled -ForegroundColor $(if ($result.IsInstalled) { "Green" } else { "Red" })
    if ($result.IsInstalled) {
        Write-Host "  Arquitectura: " -NoNewline; Write-Host $result.Architecture -ForegroundColor $(if ($result.Architecture -eq "64-bit") { "Green" } else { "Yellow" })
        Write-Host "  Versión: " -NoNewline; Write-Host $result.Version -ForegroundColor Cyan
        Write-Host "  Necesita migración: " -NoNewline; Write-Host $result.NeedsMigration -ForegroundColor $(if ($result.NeedsMigration) { "Yellow" } else { "Green" })
    }
    
    Write-Host "`nPrerequisitos:" -ForegroundColor Yellow
    Write-Host "  Permisos de administrador: " -NoNewline; Write-Host $hasAdminRights -ForegroundColor $(if ($hasAdminRights) { "Green" } else { "Red" })
    Write-Host "  Espacio en disco: " -NoNewline; Write-Host "$freeSpaceGB GB " -NoNewline -ForegroundColor $(if ($hasDiskSpace) { "Green" } else { "Red" }); Write-Host "(mínimo: $minDiskSpaceGB GB)"
    Write-Host "  Puede proceder: " -NoNewline; Write-Host $result.CanProceed -ForegroundColor $(if ($result.CanProceed) { "Green" } else { "Red" })
    
    return $result
}

Function Start-Preinstall {
    <#
    .SYNOPSIS
        Descarga ODT y archivos de Office (modo test)
    
    .DESCRIPTION
        Simula la preparación de archivos para la instalación de Office 64-bit:
        - Crea el directorio de instalación
        - Genera el archivo de configuración XML
        - Simula la descarga de ODT y archivos de Office (sin descargar realmente)
    
    .PARAMETER InstallPath
        Ruta donde se crearían los archivos (por defecto: C:\Temp\OfficeTest)
    
    .PARAMETER NeedsMigration
        Indica si se necesita migración de 32-bit a 64-bit
    
    .PARAMETER SimulateOnly
        Si es $true, solo simula sin crear archivos reales
    
    .OUTPUTS
        PSCustomObject con información sobre los archivos que se crearían
    
    .EXAMPLE
        $preinstall = Start-Preinstall
        if ($preinstall.Success) {
            Write-Host "Preparación simulada exitosamente"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InstallPath = "C:\Temp\OfficeTest",
        
        [Parameter(Mandatory = $false)]
        [bool]$NeedsMigration = $false,
        
        [Parameter(Mandatory = $false)]
        [bool]$SimulateOnly = $true
    )
    
    $result = [PSCustomObject]@{
        Success         = $false
        InstallPath     = $InstallPath
        SetupExePath    = $null
        ConfigXmlPath   = $null
        FilesDownloaded = $false
        ErrorMessage    = $null
        IsSimulation    = $SimulateOnly
    }
    
    try {
        Write-Host "`n=== TEST: PREPARACIÓN DE INSTALACIÓN ===" -ForegroundColor Cyan
        
        if ($SimulateOnly) {
            Write-Host "[MODO SIMULACIÓN - No se descargarán archivos reales]" -ForegroundColor Magenta
        }
        
        # Crear directorio si no existe
        if (-not (Test-Path -Path $InstallPath)) {
            New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
            Write-Host "✓ Directorio creado: $InstallPath" -ForegroundColor Green
        }
        else {
            Write-Host "✓ Directorio ya existe: $InstallPath" -ForegroundColor Green
        }
        
        # Crear archivo de configuración XML
        Write-Host "`nCreando archivo de configuración XML..." -ForegroundColor Yellow
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
        Write-Host "✓ Archivo de configuración creado: $configPath" -ForegroundColor Green
        $result.ConfigXmlPath = $configPath
        
        # Simular descarga de ODT
        Write-Host "`nSimulando descarga de Office Deployment Tool..." -ForegroundColor Yellow
        $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17328-20162.exe"
        Write-Host "  URL: $odtUrl" -ForegroundColor Cyan
        
        if ($SimulateOnly) {
            Write-Host "  [SIMULADO] ODT no descargado (modo test)" -ForegroundColor Magenta
        }
        else {
            $odtPath = Join-Path -Path $InstallPath -ChildPath "ODT.exe"
            Write-Host "  Descargando a: $odtPath" -ForegroundColor Cyan
            # Aquí iría la descarga real si no estuviéramos en modo simulación
        }
        
        # Simular extracción de ODT
        Write-Host "`nSimulando extracción de Office Deployment Tool..." -ForegroundColor Yellow
        $setupPath = Join-Path -Path $InstallPath -ChildPath "setup.exe"
        
        if ($SimulateOnly) {
            Write-Host "  [SIMULADO] Setup.exe estaría en: $setupPath" -ForegroundColor Magenta
            $result.SetupExePath = $setupPath
        }
        else {
            # Aquí iría la extracción real
            $result.SetupExePath = $setupPath
        }
        
        # Simular descarga de archivos de Office
        Write-Host "`nSimulando descarga de archivos de Office 64-bit..." -ForegroundColor Yellow
        Write-Host "  Canal: MonthlyEnterprise" -ForegroundColor Cyan
        Write-Host "  Idioma: Español (es-es)" -ForegroundColor Cyan
        Write-Host "  Arquitectura: 64-bit" -ForegroundColor Cyan
        
        if ($SimulateOnly) {
            Write-Host "  [SIMULADO] Archivos de Office no descargados (modo test)" -ForegroundColor Magenta
            Write-Host "  En producción, esto descargaría ~3-4 GB de datos" -ForegroundColor Magenta
        }
        
        $result.FilesDownloaded = $SimulateOnly
        $result.Success = $true
        
        Write-Host "`n✓ Preparación simulada completada exitosamente" -ForegroundColor Green
        
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        Write-Error "Error durante la preparación: $_"
    }
    
    return $result
}

Function Start-Install {
    <#
    .SYNOPSIS
        Ejecuta la instalación/migración (modo test)
    
    .DESCRIPTION
        Simula la ejecución de la instalación de Office 64-bit sin ejecutar realmente setup.exe.
        Útil para verificar que todos los parámetros y rutas son correctos.
    
    .PARAMETER SetupExePath
        Ruta completa al archivo setup.exe del Office Deployment Tool
    
    .PARAMETER ConfigXmlPath
        Ruta completa al archivo de configuración XML
    
    .PARAMETER NeedsMigration
        Indica si es una migración de 32-bit a 64-bit
    
    .PARAMETER SimulateOnly
        Si es $true, solo simula sin ejecutar la instalación real
    
    .OUTPUTS
        PSCustomObject con el resultado de la instalación simulada
    
    .EXAMPLE
        $install = Start-Install -SetupExePath "C:\Temp\OfficeTest\setup.exe" -ConfigXmlPath "C:\Temp\OfficeTest\configuration.xml"
        if ($install.Success) {
            Write-Host "Instalación simulada exitosamente"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SetupExePath,
        
        [Parameter(Mandatory = $true)]
        [string]$ConfigXmlPath,
        
        [Parameter(Mandatory = $false)]
        [bool]$NeedsMigration = $false,
        
        [Parameter(Mandatory = $false)]
        [bool]$SimulateOnly = $true
    )
    
    $result = [PSCustomObject]@{
        Success      = $false
        ExitCode     = -1
        Duration     = [TimeSpan]::Zero
        WasMigration = $NeedsMigration
        ErrorMessage = $null
        IsSimulation = $SimulateOnly
    }
    
    try {
        Write-Host "`n=== TEST: INSTALACIÓN DE OFFICE 64-BIT ===" -ForegroundColor Cyan
        
        if ($SimulateOnly) {
            Write-Host "[MODO SIMULACIÓN - No se ejecutará la instalación real]" -ForegroundColor Magenta
        }
        
        # Verificar que el archivo de configuración existe
        if (-not (Test-Path -Path $ConfigXmlPath)) {
            throw "No se encontró el archivo de configuración en: $ConfigXmlPath"
        }
        Write-Host "✓ Archivo de configuración encontrado" -ForegroundColor Green
        
        # Mensaje según tipo de instalación
        if ($NeedsMigration) {
            Write-Host "`nTipo de operación: " -NoNewline
            Write-Host "MIGRACIÓN" -ForegroundColor Yellow
            Write-Host "  De: Office 32-bit" -ForegroundColor Cyan
            Write-Host "  A: Office 64-bit" -ForegroundColor Cyan
            Write-Host "  El parámetro MigrateArch preservará configuración y datos" -ForegroundColor Cyan
        }
        else {
            Write-Host "`nTipo de operación: " -NoNewline
            Write-Host "INSTALACIÓN NUEVA" -ForegroundColor Green
            Write-Host "  Office 64-bit" -ForegroundColor Cyan
        }
        
        Write-Host "`nConfiguración:" -ForegroundColor Yellow
        Write-Host "  Canal: MonthlyEnterprise (Empresas)" -ForegroundColor Cyan
        Write-Host "  Idioma: Español (es-es)" -ForegroundColor Cyan
        Write-Host "  Arquitectura: 64-bit" -ForegroundColor Cyan
        
        # Simular la instalación
        Write-Host "`nComando que se ejecutaría:" -ForegroundColor Yellow
        Write-Host "  $SetupExePath /configure `"$ConfigXmlPath`"" -ForegroundColor Cyan
        
        if ($SimulateOnly) {
            Write-Host "`n[SIMULADO] Instalación no ejecutada (modo test)" -ForegroundColor Magenta
            Write-Host "En producción, esto tardaría varios minutos..." -ForegroundColor Magenta
            
            # Simular duración
            Start-Sleep -Seconds 2
            $result.ExitCode = 0
            $result.Duration = New-TimeSpan -Minutes 8 -Seconds 34
        }
        else {
            # Aquí iría la ejecución real
            $startTime = Get-Date
            # Start-Process -FilePath $SetupExePath -ArgumentList "/configure `"$ConfigXmlPath`"" -Wait -NoNewWindow -PassThru
            $endTime = Get-Date
            $result.Duration = $endTime - $startTime
            $result.ExitCode = 0
        }
        
        if ($result.ExitCode -eq 0) {
            $result.Success = $true
            
            if ($NeedsMigration) {
                Write-Host "`n✓ Migración simulada completada exitosamente" -ForegroundColor Green
            }
            else {
                Write-Host "`n✓ Instalación simulada completada exitosamente" -ForegroundColor Green
            }
            
            Write-Host "Duración simulada: $($result.Duration.ToString('mm\:ss'))" -ForegroundColor Cyan
        }
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        Write-Error "Error durante la instalación: $_"
    }
    
    return $result
}

Function Start-PostInstall {
    <#
    .SYNOPSIS
        Verifica instalación y limpia archivos temporales (modo test)
    
    .DESCRIPTION
        Verifica que Office 64-bit se instaló correctamente (o simula la verificación)
        y opcionalmente limpia los archivos temporales de instalación
    
    .PARAMETER InstallPath
        Ruta de los archivos temporales a limpiar
    
    .PARAMETER KeepFiles
        Si se especifica, no elimina los archivos temporales
    
    .PARAMETER SimulateOnly
        Si es $true, solo simula la verificación
    
    .OUTPUTS
        PSCustomObject con el resultado de la verificación
    
    .EXAMPLE
        $postinstall = Start-PostInstall -InstallPath "C:\Temp\OfficeTest"
        if ($postinstall.VerificationSuccess) {
            Write-Host "Verificación exitosa"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstallPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$KeepFiles,
        
        [Parameter(Mandatory = $false)]
        [bool]$SimulateOnly = $true
    )
    
    $result = [PSCustomObject]@{
        VerificationSuccess   = $false
        InstalledVersion      = $null
        InstalledArchitecture = $null
        FilesCleanedUp        = $false
        TempFilesRemaining    = @()
        ErrorMessage          = $null
        IsSimulation          = $SimulateOnly
    }
    
    try {
        Write-Host "`n=== TEST: VERIFICACIÓN POST-INSTALACIÓN ===" -ForegroundColor Cyan
        
        if ($SimulateOnly) {
            Write-Host "[MODO SIMULACIÓN - Verificación simulada]" -ForegroundColor Magenta
        }
        
        # Esperar un momento para simular actualización del sistema
        Write-Host "`nEsperando actualización del sistema..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        
        # Verificar la instalación real (esto sí se ejecuta siempre)
        function Get-OfficeInfo {
            $officeResult = [PSCustomObject]@{
                IsInstalled  = $false
                Architecture = $null
                Version      = $null
            }
            
            $registryPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
            )
            
            foreach ($regPath in $registryPaths) {
                if (Test-Path -Path $regPath) {
                    $config = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    
                    if ($config -and $config.VersionToReport) {
                        $platform = $config.Platform
                        
                        $officeResult.IsInstalled = $true
                        $officeResult.Version = $config.VersionToReport
                        
                        if ($platform -eq "x64") {
                            $officeResult.Architecture = "64-bit"
                        }
                        elseif ($platform -eq "x86") {
                            $officeResult.Architecture = "32-bit"
                        }
                        
                        return $officeResult
                    }
                }
            }
            
            return $officeResult
        }
        
        $verificationInfo = Get-OfficeInfo
        
        if ($SimulateOnly) {
            # En modo simulación, mostrar el estado actual
            Write-Host "`nEstado actual del sistema:" -ForegroundColor Yellow
            if ($verificationInfo.IsInstalled) {
                Write-Host "  Office instalado: Sí" -ForegroundColor Green
                Write-Host "  Versión: $($verificationInfo.Version)" -ForegroundColor Cyan
                Write-Host "  Arquitectura: $($verificationInfo.Architecture)" -ForegroundColor Cyan
            }
            else {
                Write-Host "  Office instalado: No" -ForegroundColor Yellow
            }
            
            Write-Host "`n[SIMULADO] En producción, aquí se verificaría Office 64-bit" -ForegroundColor Magenta
            $result.VerificationSuccess = $true
            $result.InstalledVersion = "16.0.XXXXX.XXXXX (simulado)"
            $result.InstalledArchitecture = "64-bit (simulado)"
        }
        else {
            # Verificación real
            if ($verificationInfo.IsInstalled -and $verificationInfo.Architecture -eq "64-bit") {
                $result.VerificationSuccess = $true
                $result.InstalledVersion = $verificationInfo.Version
                $result.InstalledArchitecture = "64-bit"
                
                Write-Host "`n✓ INSTALACIÓN EXITOSA" -ForegroundColor Green
                Write-Host "Office 64-bit instalado correctamente" -ForegroundColor Green
                Write-Host "Versión: $($verificationInfo.Version)" -ForegroundColor Cyan
                Write-Host "Arquitectura: 64-bit" -ForegroundColor Cyan
            }
            elseif ($verificationInfo.IsInstalled -and $verificationInfo.Architecture -eq "32-bit") {
                $result.InstalledVersion = $verificationInfo.Version
                $result.InstalledArchitecture = "32-bit"
                
                Write-Host "`n⚠ ADVERTENCIA" -ForegroundColor Yellow
                Write-Host "Office está instalado pero sigue siendo la versión de 32-bit" -ForegroundColor Yellow
                Write-Host "Versión: $($verificationInfo.Version)" -ForegroundColor Cyan
            }
            else {
                Write-Host "`n✗ ERROR" -ForegroundColor Red
                Write-Host "No se pudo verificar la instalación de Office 64-bit" -ForegroundColor Red
            }
        }
        
        # Limpiar archivos temporales si se solicita
        if (-not $KeepFiles -and (Test-Path -Path $InstallPath)) {
            Write-Host "`nLimpiando archivos temporales..." -ForegroundColor Yellow
            
            try {
                Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction Stop
                $result.FilesCleanedUp = $true
                Write-Host "✓ Archivos temporales eliminados" -ForegroundColor Green
            }
            catch {
                Write-Warning "No se pudieron eliminar todos los archivos temporales: $_"
                $result.TempFilesRemaining = @($InstallPath)
            }
        }
        elseif ($KeepFiles) {
            Write-Host "`nArchivos temporales conservados en: $InstallPath" -ForegroundColor Cyan
            $result.TempFilesRemaining = @($InstallPath)
        }
        
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        Write-Error "Error durante la verificación post-instalación: $_"
    }
    
    return $result
}

Function Start-Deploy {
    <#
    .SYNOPSIS
        Orquesta todo el flujo completo de despliegue (modo test)
    
    .DESCRIPTION
        Función principal que coordina todo el flujo de instalación de Office 64-bit en modo test:
        1. Verifica el estado actual y prerequisitos (Test-Installed)
        2. Simula la descarga de archivos necesarios (Start-Preinstall)
        3. Simula la ejecución de la instalación/migración (Start-Install)
        4. Verifica y limpia (Start-PostInstall)
        
        Por defecto ejecuta en modo simulación sin descargar ni instalar nada real.
    
    .PARAMETER InstallPath
        Ruta donde se crearían los archivos temporales (por defecto: C:\Temp\OfficeTest)
    
    .PARAMETER Force
        Fuerza la simulación incluso si Office 64-bit ya está instalado
    
    .PARAMETER KeepTempFiles
        No elimina los archivos temporales después de la simulación
    
    .PARAMETER SimulateOnly
        Si es $true (por defecto), solo simula sin ejecutar instalación real
    
    .OUTPUTS
        PSCustomObject con el resultado completo del despliegue
    
    .EXAMPLE
        Start-Deploy
        Simula el despliegue de Office 64-bit con configuración por defecto
    
    .EXAMPLE
        Start-Deploy -Force -KeepTempFiles
        Fuerza la simulación y conserva los archivos temporales
    
    .EXAMPLE
        $result = Start-Deploy -InstallPath "D:\Temp\OfficeTest"
        if ($result.Success) {
            Write-Host "Simulación exitosa"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InstallPath = "C:\Temp\OfficeTest",
        
        [Parameter(Mandatory = $false)]
        [switch]$Force,
        
        [Parameter(Mandatory = $false)]
        [switch]$KeepTempFiles,
        
        [Parameter(Mandatory = $false)]
        [bool]$SimulateOnly = $true
    )
    
    $deployResult = [PSCustomObject]@{
        Success           = $false
        Phase             = $null
        TestResult        = $null
        PreinstallResult  = $null
        InstallResult     = $null
        PostInstallResult = $null
        ErrorMessage      = $null
        IsSimulation      = $SimulateOnly
    }
    
    try {
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║  TEST: DESPLIEGUE DE MICROSOFT OFFICE 64-BIT               ║" -ForegroundColor Cyan
        Write-Host "║  Canal: MonthlyEnterprise | Idioma: Español (es-es)       ║" -ForegroundColor Cyan
        if ($SimulateOnly) {
            Write-Host "║  MODO: SIMULACIÓN (No se instalará nada real)              ║" -ForegroundColor Magenta
        }
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        
        # FASE 1: Verificación
        $deployResult.Phase = "Verificación"
        Write-Host "═══ FASE 1: VERIFICACIÓN ═══" -ForegroundColor Cyan
        $testResult = Test-Installed
        $deployResult.TestResult = $testResult
        
        # Verificar prerequisitos
        if (-not $testResult.CanProceed) {
            if (-not $testResult.Prerequisites.HasAdminRights) {
                throw "Se requieren permisos de administrador para instalar Office"
            }
            if (-not $testResult.Prerequisites.HasDiskSpace) {
                throw "Espacio en disco insuficiente. Se requieren al menos $($testResult.Prerequisites.MinDiskSpaceGB) GB"
            }
        }
        
        # Decidir si proceder
        if ($testResult.IsInstalled -and $testResult.Architecture -eq "64-bit" -and -not $Force) {
            Write-Host "`n✓ Office 64-bit ya está instalado" -ForegroundColor Green
            Write-Host "Versión: $($testResult.Version)" -ForegroundColor Cyan
            Write-Host "Use el parámetro -Force para simular reinstalación" -ForegroundColor Yellow
            $deployResult.Success = $true
            return $deployResult
        }
        
        # FASE 2: Preparación
        $deployResult.Phase = "Preparación"
        Write-Host "`n═══ FASE 2: PREPARACIÓN ═══" -ForegroundColor Cyan
        $preinstallResult = Start-Preinstall -InstallPath $InstallPath -NeedsMigration $testResult.NeedsMigration -SimulateOnly $SimulateOnly
        $deployResult.PreinstallResult = $preinstallResult
        
        if (-not $preinstallResult.Success) {
            throw "Error en la preparación: $($preinstallResult.ErrorMessage)"
        }
        
        # FASE 3: Instalación
        $deployResult.Phase = "Instalación"
        Write-Host "`n═══ FASE 3: INSTALACIÓN ═══" -ForegroundColor Cyan
        $installResult = Start-Install -SetupExePath $preinstallResult.SetupExePath `
            -ConfigXmlPath $preinstallResult.ConfigXmlPath `
            -NeedsMigration $testResult.NeedsMigration `
            -SimulateOnly $SimulateOnly
        $deployResult.InstallResult = $installResult
        
        if (-not $installResult.Success) {
            throw "Error en la instalación: $($installResult.ErrorMessage)"
        }
        
        # FASE 4: Verificación Post-Instalación
        $deployResult.Phase = "Verificación Post-Instalación"
        Write-Host "`n═══ FASE 4: VERIFICACIÓN POST-INSTALACIÓN ═══" -ForegroundColor Cyan
        $postInstallResult = Start-PostInstall -InstallPath $InstallPath -KeepFiles:$KeepTempFiles -SimulateOnly $SimulateOnly
        $deployResult.PostInstallResult = $postInstallResult
        
        if ($postInstallResult.VerificationSuccess) {
            $deployResult.Success = $true
            
            Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
            if ($SimulateOnly) {
                Write-Host "║  ✓ SIMULACIÓN COMPLETADA EXITOSAMENTE                      ║" -ForegroundColor Green
            }
            else {
                Write-Host "║  ✓ DESPLIEGUE COMPLETADO EXITOSAMENTE                      ║" -ForegroundColor Green
            }
            Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
            Write-Host ""
            Write-Host "Resumen:" -ForegroundColor Cyan
            Write-Host "  Versión: $($postInstallResult.InstalledVersion)" -ForegroundColor Green
            Write-Host "  Arquitectura: $($postInstallResult.InstalledArchitecture)" -ForegroundColor Green
            Write-Host "  Duración de instalación: $($installResult.Duration.ToString('mm\:ss'))" -ForegroundColor Cyan
            Write-Host "  Canal: MonthlyEnterprise (Empresas)" -ForegroundColor Cyan
            Write-Host "  Idioma: Español (es-es)" -ForegroundColor Cyan
            
            if ($testResult.NeedsMigration) {
                Write-Host "  Tipo: Migración de 32-bit a 64-bit" -ForegroundColor Cyan
            }
            else {
                Write-Host "  Tipo: Instalación nueva" -ForegroundColor Cyan
            }
            
            if ($SimulateOnly) {
                Write-Host "`n  [MODO SIMULACIÓN - No se realizó instalación real]" -ForegroundColor Magenta
            }
        }
        else {
            Write-Host "`n⚠ ADVERTENCIA: La verificación post-instalación falló" -ForegroundColor Yellow
            Write-Host "La instalación puede no haberse completado correctamente" -ForegroundColor Yellow
        }
    }
    catch {
        $deployResult.ErrorMessage = $_.Exception.Message
        
        Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Red
        Write-Host "║  ✗ ERROR EN EL DESPLIEGUE                                  ║" -ForegroundColor Red
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        Write-Host ""
        Write-Host "Fase: $($deployResult.Phase)" -ForegroundColor Yellow
        Write-Host "Error: $($deployResult.ErrorMessage)" -ForegroundColor Red
        
        Write-Error "Error durante el despliegue: $_"
    }
    
    return $deployResult
}

#endregion

# Exportar funciones
Export-ModuleMember -Function Test-Installed, Start-Preinstall, Start-Install, Start-PostInstall, Start-Deploy
