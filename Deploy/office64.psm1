<#
.SYNOPSIS
    


.DESCRIPTION
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/office64.psm1') | iex ; Start-Install
    
   

.NOTES
    Version:        0.1.0
    Author:         Garikoitz Belarbide    
    Creation Date:  14/01/2026

#>

#region [Functions]-------------------------------------------------------------

Function Test-OfficeInstalled {
    <#
    .SYNOPSIS
        Comprueba si Microsoft Office esta instalado en el sistema
    
    .DESCRIPTION
        Verifica la existencia de Office comprobando el registro de Windows
        y determina si es version de 32 o 64 bits
    
    .OUTPUTS
        PSCustomObject con propiedades:
        - IsInstalled: Boolean indicando si Office esta instalado
        - Architecture: String con "32-bit", "64-bit" o $null
        - Version: String con la version instalada o $null
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()
    
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
                # Verificar la plataforma instalada usando la propiedad Platform
                $platform = $config.Platform
                
                $result.IsInstalled = $true
                $result.Version = $config.VersionToReport
                
                if ($platform -eq "x64") {
                    $result.Architecture = "64-bit"
                    Write-Verbose "Office 64-bit encontrado: $($config.VersionToReport)"
                }
                elseif ($platform -eq "x86") {
                    $result.Architecture = "32-bit"
                    Write-Verbose "Office 32-bit encontrado: $($config.VersionToReport)"
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
                Write-Verbose "Office $($msiInfo.Arch) encontrado (MSI): $installPath"
                return $result
            }
        }
    }
    
    return $result
}

Function Test-Installed {
    <#
    .SYNOPSIS
        Verifica el estado de Office y los prerequisitos del sistema
    
    .DESCRIPTION
        Comprueba si Office esta instalado, su arquitectura, version y verifica
        que el sistema cumple con los prerequisitos necesarios para la instalacion
    
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
    
    Write-Verbose "Verificando estado de Office y prerequisitos del sistema..."
    
    # Obtener información de Office instalado
    $officeInfo = Test-OfficeInstalled
    
    # Verificar permisos de administrador
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $hasAdminRights = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    # Verificar espacio en disco (unidad C:)
    $drive = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    $freeSpaceGB = if ($drive) { [math]::Round($drive.Free / 1GB, 2) } else { 0 }
    $minDiskSpaceGB = 10
    $hasDiskSpace = $freeSpaceGB -ge $minDiskSpaceGB
    
    # Determinar si necesita migracion
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
    
    # Logging
    Write-Verbose "Estado de Office:"
    Write-Verbose "  - Instalado: $($result.IsInstalled)"
    Write-Verbose "  - Arquitectura: $($result.Architecture)"
    Write-Verbose "  - Version: $($result.Version)"
    Write-Verbose "  - Necesita migracion: $($result.NeedsMigration)"
    Write-Verbose "Prerequisitos:"
    Write-Verbose "  - Permisos de administrador: $hasAdminRights"
    Write-Verbose "  - Espacio en disco: $freeSpaceGB GB (mínimo: $minDiskSpaceGB GB)"
    Write-Verbose "  - Puede proceder: $($result.CanProceed)"
    
    return $result
}

Function Start-Preinstall {
    <#
    .SYNOPSIS
        Prepara los archivos necesarios para la instalacion de Office 64-bit
    
    .DESCRIPTION
        Descarga Office Deployment Tool, lo extrae, crea el archivo de configuración XML
        y descarga los archivos de instalacion de Office 64-bit
    
    .PARAMETER InstallPath
        Ruta donde se descargarán los archivos (por defecto: C:\Temp\Office)
    
    .PARAMETER NeedsMigration
        Indica si se necesita migracion de 32-bit a 64-bit
    
    .OUTPUTS
        PSCustomObject con información sobre los archivos descargados
    
    .EXAMPLE
        $preinstall = Start-Preinstall
        if ($preinstall.Success) {
            Write-Host "Archivos listos en: $($preinstall.InstallPath)"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InstallPath = "C:\Temp\Office",
        
        [Parameter(Mandatory = $false)]
        [bool]$NeedsMigration = $false
    )
    
    $result = [PSCustomObject]@{
        Success         = $false
        InstallPath     = $InstallPath
        SetupExePath    = $null
        ConfigXmlPath   = $null
        FilesDownloaded = $false
        ErrorMessage    = $null
    }
    
    try {
        Write-Host "=== PREPARACIÓN DE INSTALACION ===" -ForegroundColor Cyan
        
        # Crear directorio si no existe
        if (-not (Test-Path -Path $InstallPath)) {
            New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
            Write-Host "Directorio creado: $InstallPath" -ForegroundColor Green
        }
        
        # Crear archivo de configuración XML
        Write-Host "Creando archivo de configuración..." -ForegroundColor Yellow
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
        Write-Host "Archivo de configuración creado: $configPath" -ForegroundColor Green
        $result.ConfigXmlPath = $configPath
        
        # Descargar Office Deployment Tool
        $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17328-20162.exe"
        $odtPath = Join-Path -Path $InstallPath -ChildPath "ODT.exe"
        
        Write-Host "Descargando Office Deployment Tool..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing
        Write-Host "Office Deployment Tool descargado" -ForegroundColor Green
        
        # Extraer ODT
        Write-Host "Extrayendo Office Deployment Tool..." -ForegroundColor Yellow
        Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$InstallPath" -Wait -NoNewWindow
        Write-Host "Office Deployment Tool extraído" -ForegroundColor Green
        
        # Verificar que setup.exe existe
        $setupPath = Join-Path -Path $InstallPath -ChildPath "setup.exe"
        if (-not (Test-Path -Path $setupPath)) {
            throw "No se encontró setup.exe después de extraer ODT"
        }
        $result.SetupExePath = $setupPath
        
        # Descargar archivos de Office
        Write-Host "Descargando archivos de Office 64-bit..." -ForegroundColor Yellow
        Write-Host "Esto puede tardar varios minutos dependiendo de la conexión..." -ForegroundColor Cyan
        Start-Process -FilePath $setupPath -ArgumentList "/download `"$configPath`"" -Wait -NoNewWindow
        Write-Host "Archivos de Office descargados correctamente" -ForegroundColor Green
        
        $result.FilesDownloaded = $true
        $result.Success = $true
        
        Write-Host "`n✓ Preparacion completada exitosamente" -ForegroundColor Green
        
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        Write-Error "Error durante la preparacion: $_"
    }
    
    return $result
}

Function Start-PostInstall {
    <#
    .SYNOPSIS
        Verifica la instalacion y limpia archivos temporales
    
    .DESCRIPTION
        Verifica que Office 64-bit se instaló correctamente y opcionalmente
        limpia los archivos temporales de instalacion
    
    .PARAMETER InstallPath
        Ruta de los archivos temporales a limpiar
    
    .PARAMETER KeepFiles
        Si se especifica, no elimina los archivos temporales
    
    .OUTPUTS
        PSCustomObject con el resultado de la verificacion
    
    .EXAMPLE
        $postinstall = Start-PostInstall -InstallPath "C:\Temp\Office"
        if ($postinstall.VerificationSuccess) {
            Write-Host "Office 64-bit instalado correctamente"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstallPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$KeepFiles
    )
    
    $result = [PSCustomObject]@{
        VerificationSuccess   = $false
        InstalledVersion      = $null
        InstalledArchitecture = $null
        FilesCleanedUp        = $false
        TempFilesRemaining    = @()
        ErrorMessage          = $null
    }
    
    try {
        Write-Host "`n=== VERIFICACIÓN POST-INSTALACION ===" -ForegroundColor Cyan
        
        # Esperar un momento para que el sistema se actualice
        Write-Host "Esperando actualización del sistema..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        
        # Verificar la instalacion
        $verificationInfo = Test-OfficeInstalled
        
        if ($verificationInfo.IsInstalled -and $verificationInfo.Architecture -eq "64-bit") {
            $result.VerificationSuccess = $true
            $result.InstalledVersion = $verificationInfo.Version
            $result.InstalledArchitecture = "64-bit"
            
            Write-Host "`n[OK] INSTALACION EXITOSA" -ForegroundColor Green
            Write-Host "Office 64-bit instalado correctamente" -ForegroundColor Green
            Write-Host "Version: $($verificationInfo.Version)" -ForegroundColor Cyan
            Write-Host "Arquitectura: 64-bit" -ForegroundColor Cyan
        }
        elseif ($verificationInfo.IsInstalled -and $verificationInfo.Architecture -eq "32-bit") {
            $result.InstalledVersion = $verificationInfo.Version
            $result.InstalledArchitecture = "32-bit"
            
            Write-Host "`n[!] ADVERTENCIA" -ForegroundColor Yellow
            Write-Host "Office esta instalado pero sigue siendo la version de 32-bit" -ForegroundColor Yellow
            Write-Host "Version: $($verificationInfo.Version)" -ForegroundColor Cyan
            Write-Host "Es posible que la migracion no se haya completado correctamente." -ForegroundColor Yellow
        }
        else {
            Write-Host "`n[X] ERROR" -ForegroundColor Red
            Write-Host "No se pudo verificar la instalacion de Office 64-bit" -ForegroundColor Red
        }
        
        # Limpiar archivos temporales si se solicita
        if (-not $KeepFiles -and (Test-Path -Path $InstallPath)) {
            Write-Host "`nLimpiando archivos temporales..." -ForegroundColor Yellow
            
            try {
                Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction Stop
                $result.FilesCleanedUp = $true
                Write-Host "Archivos temporales eliminados" -ForegroundColor Green
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
        Write-Error "Error durante la verificacion post-instalacion: $_"
    }
    
    return $result
}

Function Start-Install {
    <#
    .SYNOPSIS
        Ejecuta la instalacion/migracion de Office 64-bit
    
    .DESCRIPTION
        Ejecuta setup.exe con el archivo de configuración XML para instalar o migrar Office 64-bit.
        Esta función asume que los archivos ya han sido descargados por Start-Preinstall.
    
    .PARAMETER SetupExePath
        Ruta completa al archivo setup.exe del Office Deployment Tool
    
    .PARAMETER ConfigXmlPath
        Ruta completa al archivo de configuración XML
    
    .PARAMETER NeedsMigration
        Indica si es una migracion de 32-bit a 64-bit
    
    .OUTPUTS
        PSCustomObject con el resultado de la instalacion
    
    .EXAMPLE
        $install = Start-Install -SetupExePath "C:\Temp\Office\setup.exe" -ConfigXmlPath "C:\Temp\Office\configuration.xml"
        if ($install.Success) {
            Write-Host "Instalacion exitosa"
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
        [bool]$NeedsMigration = $false
    )
    
    $result = [PSCustomObject]@{
        Success      = $false
        ExitCode     = -1
        Duration     = [TimeSpan]::Zero
        WasMigration = $NeedsMigration
        ErrorMessage = $null
    }
    
    try {
        Write-Host "`n=== INSTALACION DE OFFICE 64-BIT ===" -ForegroundColor Cyan
        
        # Verificar que los archivos existen
        if (-not (Test-Path -Path $SetupExePath)) {
            throw "No se encontró setup.exe en: $SetupExePath"
        }
        
        if (-not (Test-Path -Path $ConfigXmlPath)) {
            throw "No se encontró el archivo de configuración en: $ConfigXmlPath"
        }
        
        # Mensaje según tipo de instalacion
        if ($NeedsMigration) {
            Write-Host "Migrando de Office 32-bit a 64-bit..." -ForegroundColor Yellow
            Write-Host "El parametro MigrateArch preservará su configuración y datos." -ForegroundColor Cyan
        }
        else {
            Write-Host "Instalando Office 64-bit..." -ForegroundColor Yellow
        }
        
        Write-Host "Canal: MonthlyEnterprise (Empresas)" -ForegroundColor Cyan
        Write-Host "Idioma: Español (es-es)" -ForegroundColor Cyan
        Write-Host "`nEsto puede tardar varios minutos..." -ForegroundColor Yellow
        
        # Ejecutar la instalacion
        $startTime = Get-Date
        $process = Start-Process -FilePath $SetupExePath -ArgumentList "/configure `"$ConfigXmlPath`"" -Wait -NoNewWindow -PassThru
        $endTime = Get-Date
        
        $result.ExitCode = $process.ExitCode
        $result.Duration = $endTime - $startTime
        
        if ($process.ExitCode -eq 0) {
            $result.Success = $true
            
            if ($NeedsMigration) {
                Write-Host "`n[OK] Migracion completada exitosamente" -ForegroundColor Green
            }
            else {
                Write-Host "`n[OK] Instalacion completada exitosamente" -ForegroundColor Green
            }
            
            Write-Host "Duración: $($result.Duration.ToString('mm\:ss'))" -ForegroundColor Cyan
        }
        else {
            $result.ErrorMessage = "Setup.exe finalizó con código de salida: $($process.ExitCode)"
            Write-Warning $result.ErrorMessage
        }
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        Write-Error "Error durante la instalacion: $_"
    }
    
    return $result
}

Function Start-Deploy {
    <#
    .SYNOPSIS
        Orquesta el despliegue completo de Office 64-bit
    
    .DESCRIPTION
        Función principal que coordina todo el flujo de instalacion de Office 64-bit:
        1. Verifica el estado actual y prerequisitos (Test-Installed)
        2. Descarga los archivos necesarios (Start-Preinstall)
        3. Ejecuta la instalacion/migracion (Start-Install)
        4. Verifica y limpia (Start-PostInstall)
    
    .PARAMETER InstallPath
        Ruta donde se descargarán los archivos temporales (por defecto: C:\Temp\Office)
    
    .PARAMETER Force
        Fuerza la instalacion incluso si Office 64-bit ya esta instalado
    
    .PARAMETER KeepFiles
        No elimina los archivos temporales después de la instalacion
    
    .OUTPUTS
        PSCustomObject con el resultado completo del despliegue
    
    .EXAMPLE
        Start-Deploy
        Despliega Office 64-bit con configuración por defecto
    
    .EXAMPLE
        Start-Deploy -Force -KeepTempFiles
        Fuerza la reinstalacion y conserva los archivos temporales
    
    .EXAMPLE
        $result = Start-Deploy -InstallPath "D:\Temp\Office"
        if ($result.Success) {
            Write-Host "Despliegue exitoso"
        }
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InstallPath = "C:\Temp\Office",
        
        [Parameter(Mandatory = $false)]
        [switch]$Force,
        
        [Parameter(Mandatory = $false)]
        [switch]$KeepTempFiles
    )
    
    $deployResult = [PSCustomObject]@{
        Success           = $false
        Phase             = $null
        TestResult        = $null
        PreinstallResult  = $null
        InstallResult     = $null
        PostInstallResult = $null
        ErrorMessage      = $null
    }
    
    try {
        Write-Host "================================================================" -ForegroundColor Cyan
        Write-Host "  DESPLIEGUE DE MICROSOFT OFFICE 64-BIT                       " -ForegroundColor Cyan
        Write-Host "  Canal: MonthlyEnterprise | Idioma: Espanol (es-es)         " -ForegroundColor Cyan
        Write-Host "================================================================" -ForegroundColor Cyan
        Write-Host ""
        
        # FASE 1: Verificacion
        $deployResult.Phase = "Verificacion"
        Write-Host "=== FASE 1: VERIFICACION ===" -ForegroundColor Cyan
        $testResult = Test-Installed
        $deployResult.TestResult = $testResult
        
        # Mostrar información del sistema
        Write-Host "`nEstado del sistema:" -ForegroundColor Yellow
        Write-Host "  Office instalado: $($testResult.IsInstalled)" -ForegroundColor Cyan
        if ($testResult.IsInstalled) {
            Write-Host "  Arquitectura: $($testResult.Architecture)" -ForegroundColor Cyan
            Write-Host "  Version: $($testResult.Version)" -ForegroundColor Cyan
        }
        Write-Host "  Permisos de administrador: $($testResult.Prerequisites.HasAdminRights)" -ForegroundColor Cyan
        Write-Host "  Espacio en disco: $($testResult.Prerequisites.FreeSpaceGB) GB" -ForegroundColor Cyan
        
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
            Write-Host "`n[OK] Office 64-bit ya esta instalado" -ForegroundColor Green
            Write-Host "Version: $($testResult.Version)" -ForegroundColor Cyan
            Write-Host "Use el parametro -Force para reinstalar" -ForegroundColor Yellow
            $deployResult.Success = $true
            return $deployResult
        }
        
        # FASE 2: Preparacion
        $deployResult.Phase = "Preparacion"
        Write-Host "`n=== FASE 2: PREPARACIÓN ===" -ForegroundColor Cyan
        $preinstallResult = Start-Preinstall -InstallPath $InstallPath -NeedsMigration $testResult.NeedsMigration
        $deployResult.PreinstallResult = $preinstallResult
        
        if (-not $preinstallResult.Success) {
            throw "Error en la preparacion: $($preinstallResult.ErrorMessage)"
        }
        
        # FASE 3: Instalacion
        $deployResult.Phase = "Instalacion"
        Write-Host "`n=== FASE 3: INSTALACION ===" -ForegroundColor Cyan
        $installResult = Start-Install -SetupExePath $preinstallResult.SetupExePath `
            -ConfigXmlPath $preinstallResult.ConfigXmlPath `
            -NeedsMigration $testResult.NeedsMigration
        $deployResult.InstallResult = $installResult
        
        if (-not $installResult.Success) {
            throw "Error en la instalacion: $($installResult.ErrorMessage)"
        }
        
        # FASE 4: Verificacion Post-Instalacion
        $deployResult.Phase = "Verificacion Post-Instalacion"
        Write-Host "`n=== FASE 4: VERIFICACIÓN POST-INSTALACION ===" -ForegroundColor Cyan
        $postInstallResult = Start-PostInstall -InstallPath $InstallPath -KeepFiles:$KeepTempFiles
        $deployResult.PostInstallResult = $postInstallResult
        
        if ($postInstallResult.VerificationSuccess) {
            $deployResult.Success = $true
            
            Write-Host "`n================================================================" -ForegroundColor Green
            Write-Host "  [OK] DESPLIEGUE COMPLETADO EXITOSAMENTE                    " -ForegroundColor Green
            Write-Host "================================================================" -ForegroundColor Green
            Write-Host ""
            Write-Host "Resumen:" -ForegroundColor Cyan
            Write-Host "  Version instalada: $($postInstallResult.InstalledVersion)" -ForegroundColor Green
            Write-Host "  Arquitectura: $($postInstallResult.InstalledArchitecture)" -ForegroundColor Green
            Write-Host "  Duración de instalacion: $($installResult.Duration.ToString('mm\:ss'))" -ForegroundColor Cyan
            Write-Host "  Canal: MonthlyEnterprise (Empresas)" -ForegroundColor Cyan
            Write-Host "  Idioma: Español (es-es)" -ForegroundColor Cyan
            
            if ($testResult.NeedsMigration) {
                Write-Host "  Tipo: Migracion de 32-bit a 64-bit" -ForegroundColor Cyan
            }
            else {
                Write-Host "  Tipo: Instalacion nueva" -ForegroundColor Cyan
            }
        }
        else {
            Write-Host "`n[!] ADVERTENCIA: La verificacion post-instalacion fallo" -ForegroundColor Yellow
            Write-Host "La instalacion puede no haberse completado correctamente" -ForegroundColor Yellow
        }
    }
    catch {
        $deployResult.ErrorMessage = $_.Exception.Message
        
        Write-Host "`n================================================================" -ForegroundColor Red
        Write-Host "  [ERROR] ERROR EN EL DESPLIEGUE                              " -ForegroundColor Red
        Write-Host "================================================================" -ForegroundColor Red
        Write-Host ""
        Write-Host "Fase: $($deployResult.Phase)" -ForegroundColor Yellow
        Write-Host "Error: $($deployResult.ErrorMessage)" -ForegroundColor Red
        
        Write-Error "Error durante el despliegue: $_"
    }
    
    return $deployResult
}

Function Get-DeployCnf {
    <#
    .SYNOPSIS
        Devuelve la configuracion por defecto para el despliegue de Office 64-bit
    
    .DESCRIPTION
        Esta funcion proporciona los valores por defecto para N, Every y Message
        que seran utilizados por Start-GbDeploy si no se especifican manualmente.
    
    .OUTPUTS
        PSCustomObject con las propiedades N, Every y Message
    
    .EXAMPLE
        $config = Get-DeployCnf
        Start-GbDeploy -Name "office64" -N $config.N -Every $config.Every -Message $config.Message
    #>
    
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()
    
    return [PSCustomObject]@{
        N       = 3
        Every   = 60
        Message = "Gbelarbide:Se requiere actualizar Office a la version de 64-bit para mejorar el rendimiento y compatibilidad. Durante la actualizacion podras usar tu ordenador, pero no podras usar las aplicaciones de Office."
    }
}

#endregion