<#
.SYNOPSIS
    


.DESCRIPTION
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/gbtools.psm1') | iex ; Get-HolaMundo
    
   

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
        @{Path = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"; Arch = "64-bit" },
        @{Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"; Arch = "32-bit" }
    )
    
    foreach ($regInfo in $registryPaths) {
        if (Test-Path -Path $regInfo.Path) {
            $config = Get-ItemProperty -Path $regInfo.Path -ErrorAction SilentlyContinue
            
            if ($config) {
                # Verificar la plataforma instalada
                $platform = $config.Platform -replace '\s', ''
                
                if ($platform -eq "x64" -or $regInfo.Arch -eq "64-bit") {
                    $result.IsInstalled = $true
                    $result.Architecture = "64-bit"
                    $result.Version = $config.VersionToReport
                    Write-Verbose "Office 64-bit encontrado: $($config.VersionToReport)"
                    return $result
                }
                elseif ($platform -eq "x86" -or $regInfo.Arch -eq "32-bit") {
                    $result.IsInstalled = $true
                    $result.Architecture = "32-bit"
                    $result.Version = $config.VersionToReport
                    Write-Verbose "Office 32-bit encontrado: $($config.VersionToReport)"
                    return $result
                }
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
        
        Antes de instalar, comprueba si Office ya está instalado en el sistema:
        - Si Office 64-bit ya está instalado: Se detiene la instalación
        - Si Office 32-bit está instalado: Procede con la instalación de 64-bit
        - Si no hay Office instalado: Procede con la instalación
        
        Use el parámetro -Force para reinstalar sin importar la versión existente.
    
    .PARAMETER InstallPath
        Ruta donde se descargará el instalador de Office. Por defecto: C:\Temp\Office
    
    .PARAMETER Force
        Fuerza la instalación incluso si Office 64-bit ya está instalado en el sistema
    
    .EXAMPLE
        Start-Install
        Instala Office 64-bit en español. Si detecta Office 32-bit, procede con la instalación.
    
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
        
        # Comprobar si Office ya está instalado
        if (-not $Force) {
            Write-Host "Comprobando si Office ya está instalado..." -ForegroundColor Yellow
            $officeInfo = Test-OfficeInstalled
            
            if ($officeInfo.IsInstalled) {
                if ($officeInfo.Architecture -eq "64-bit") {
                    Write-Host "Office 64-bit ya está instalado en el sistema." -ForegroundColor Green
                    Write-Host "Versión: $($officeInfo.Version)" -ForegroundColor Cyan
                    Write-Host "Si desea reinstalar, use el parámetro -Force" -ForegroundColor Yellow
                    return
                }
                elseif ($officeInfo.Architecture -eq "32-bit") {
                    Write-Host "Office 32-bit detectado en el sistema." -ForegroundColor Yellow
                    Write-Host "Versión: $($officeInfo.Version)" -ForegroundColor Cyan
                    Write-Host "Se procederá a instalar Office 64-bit." -ForegroundColor Green
                    Write-Host "NOTA: Es recomendable desinstalar Office 32-bit primero para evitar conflictos." -ForegroundColor Yellow
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
        if (-not (Test-Path -Path $InstallPath)) {
            New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
            Write-Host "Directorio creado: $InstallPath" -ForegroundColor Green
        }
    }
    
    process {
        try {
            # Crear archivo de configuración XML
            $configXML = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365BusinessRetail">
      <Language ID="es-es" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Teams" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
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
            Write-Host "Archivos de Office descargados" -ForegroundColor Green
            
            # Instalar Office
            Write-Host "Instalando Office 64-bit en castellano..." -ForegroundColor Yellow
            Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -Wait -NoNewWindow
            Write-Host "Instalación completada exitosamente" -ForegroundColor Green
            
        }
        catch {
            Write-Error "Error durante la instalación: $_"
            throw
        }
    }
    
    end {
        Write-Host "`nInstalación de Office 64-bit finalizada" -ForegroundColor Cyan
        Write-Host "Canal de actualización: MonthlyEnterprise (Empresas)" -ForegroundColor Cyan
        Write-Host "Idioma: Español (es-es)" -ForegroundColor Cyan
    }
}

#endregion