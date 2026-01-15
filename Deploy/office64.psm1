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
        Comprueba si Microsoft Office está instalado en el sistema
    
    .DESCRIPTION
        Verifica la existencia de Office comprobando el registro de Windows
        y determina si es versión de 32 o 64 bits
    
    .OUTPUTS
        PSCustomObject con propiedades:
        - IsInstalled: Boolean indicando si Office está instalado
        - Architecture: String con "32-bit", "64-bit" o $null
        - Version: String con la versión instalada o $null
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

Function Start-Install {
    <#
    .SYNOPSIS
        Instala Microsoft Office 64-bit en castellano con el canal para empresas
    
    .DESCRIPTION
        Descarga e instala Microsoft Office 64-bit en español utilizando el canal 
        de actualización para empresas (MonthlyEnterprise). Incluye Word, Excel, 
        PowerPoint, Outlook, OneNote, Access y Publisher.
        
        Proceso de instalación cuando se detecta Office 32-bit:
        - Utiliza el parámetro MigrateArch="TRUE" para migrar automáticamente de 32-bit a 64-bit
        - La migración preserva la configuración y datos del usuario
        - No requiere desinstalación manual previa
        
        Comportamiento según versión instalada:
        - Si Office 64-bit ya está instalado: Se detiene la instalación
        - Si Office 32-bit está instalado: Migra automáticamente a 64-bit
        - Si no hay Office instalado: Procede con la instalación directamente
        
        Use el parámetro -Force para reinstalar sin importar la versión existente.
    
    .PARAMETER InstallPath
        Ruta donde se descargará el instalador de Office. Por defecto: C:\Temp\Office
    
    .PARAMETER Force
        Fuerza la instalación incluso si Office 64-bit ya está instalado en el sistema
    
    .EXAMPLE
        Start-Install
        Instala Office 64-bit en español. Si detecta Office 32-bit, migra automáticamente a 64-bit usando MigrateArch.
    
    .EXAMPLE
        Start-Install -InstallPath "D:\Downloads\Office"
        Instala Office 64-bit en español en la ruta especificada
    
    .EXAMPLE
        Start-Install -Force
        Reinstala Office 64-bit incluso si ya está instalado
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InstallPath = "C:\Temp\Office",
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    begin {
        Write-Host "Iniciando instalación de Office 64-bit en castellano..." -ForegroundColor Cyan
        
        # Variable para controlar si debemos proceder con la instalación
        $script:shouldProceed = $true
        $script:needsMigration = $false
        
        # Comprobar si Office ya está instalado
        if (-not $Force) {
            Write-Host "Comprobando si Office ya está instalado..." -ForegroundColor Yellow
            $officeInfo = Test-OfficeInstalled
            
            if ($officeInfo.IsInstalled) {
                if ($officeInfo.Architecture -eq "64-bit") {
                    Write-Host "Office 64-bit ya está instalado en el sistema." -ForegroundColor Green
                    Write-Host "Versión: $($officeInfo.Version)" -ForegroundColor Cyan
                    Write-Host "Si desea reinstalar, use el parámetro -Force" -ForegroundColor Yellow
                    $script:shouldProceed = $false
                    return
                }
                elseif ($officeInfo.Architecture -eq "32-bit") {
                    Write-Host "Office 32-bit detectado en el sistema." -ForegroundColor Yellow
                    Write-Host "Versión: $($officeInfo.Version)" -ForegroundColor Cyan
                    Write-Host "Se migrará automáticamente a Office 64-bit usando MigrateArch." -ForegroundColor Green
                    
                    # Marcar que necesitamos migrar de 32-bit a 64-bit
                    $script:needsMigration = $true
                }
            }
            else {
                Write-Host "Office no está instalado. Procediendo con la instalación..." -ForegroundColor Green
            }
        }
        else {
            Write-Host "Instalación forzada. Omitiendo comprobación..." -ForegroundColor Yellow
        }
        
        # Crear directorio si no existe
        if ($script:shouldProceed -and -not (Test-Path -Path $InstallPath)) {
            New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
            Write-Host "Directorio creado: $InstallPath" -ForegroundColor Green
        }
    }
    
    process {
        if (-not $script:shouldProceed) { return }
        
        try {
            # Crear archivo de configuración XML
            # MigrateArch="TRUE" permite migrar automáticamente de 32-bit a 64-bit
            # FORCEAPPSHUTDOWN="TRUE" cierra automáticamente las aplicaciones de Office abiertas
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
            
            # Descargar Office
            $setupPath = Join-Path -Path $InstallPath -ChildPath "setup.exe"
            Write-Host "Descargando archivos de Office 64-bit..." -ForegroundColor Yellow
            Start-Process -FilePath $setupPath -ArgumentList "/download `"$configPath`"" -Wait -NoNewWindow
            Write-Host "Archivos de Office descargados correctamente" -ForegroundColor Green
            
            # Instalar Office (con migración automática si es necesario)
            if ($script:needsMigration) {
                Write-Host "`nMigrando de Office 32-bit a 64-bit en castellano..." -ForegroundColor Yellow
                Write-Host "El parámetro MigrateArch preservará su configuración y datos." -ForegroundColor Cyan
            }
            else {
                Write-Host "`nInstalando Office 64-bit en castellano..." -ForegroundColor Yellow
            }
            
            Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -Wait -NoNewWindow
            
            if ($script:needsMigration) {
                Write-Host "Migración completada exitosamente" -ForegroundColor Green
            }
            else {
                Write-Host "Instalación completada exitosamente" -ForegroundColor Green
            }
        }
        catch {
            Write-Error "Error durante la instalación: $_"
            throw
        }
    }
    
    end {
        if (-not $script:shouldProceed) { return }
        
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Verificando instalación de Office 64-bit" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        
        # Esperar un momento para que el sistema se actualice
        Start-Sleep -Seconds 3
        
        # Verificar la instalación
        $verificationInfo = Test-OfficeInstalled
        
        if ($verificationInfo.IsInstalled -and $verificationInfo.Architecture -eq "64-bit") {
            Write-Host "`n✓ INSTALACIÓN EXITOSA" -ForegroundColor Green
            Write-Host "Office 64-bit instalado correctamente" -ForegroundColor Green
            Write-Host "Versión: $($verificationInfo.Version)" -ForegroundColor Cyan
            Write-Host "Arquitectura: 64-bit" -ForegroundColor Cyan
            Write-Host "Canal de actualización: MonthlyEnterprise (Empresas)" -ForegroundColor Cyan
            Write-Host "Idioma: Español (es-es)" -ForegroundColor Cyan
        }
        elseif ($verificationInfo.IsInstalled -and $verificationInfo.Architecture -eq "32-bit") {
            Write-Host "`n⚠ ADVERTENCIA" -ForegroundColor Yellow
            Write-Host "Office está instalado pero sigue siendo la versión de 32-bit" -ForegroundColor Yellow
            Write-Host "Versión: $($verificationInfo.Version)" -ForegroundColor Cyan
            Write-Host "Es posible que la migración no se haya completado correctamente." -ForegroundColor Yellow
            Write-Host "Intente ejecutar nuevamente con el parámetro -Force" -ForegroundColor Yellow
        }
        else {
            Write-Host "`n✗ ERROR" -ForegroundColor Red
            Write-Host "No se pudo verificar la instalación de Office 64-bit" -ForegroundColor Red
            Write-Host "Por favor, verifique manualmente la instalación." -ForegroundColor Yellow
        }
        
        Write-Host "`n========================================" -ForegroundColor Cyan
    }
}

#endregion