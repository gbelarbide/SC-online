function Install-AdobeReader {
    Write-Host "------------------------------"
    Write-Host "Adobe Acrobat Reader DC"
    Write-Host "------------------------------"

    Write-Host "Borrando version previa si existe..."
    $app = Get-CimInstance -Query "SELECT * FROM Win32_Product WHERE Name like '%Adobe Reader%'"
    if ($app) { $app | Invoke-CimMethod -MethodName Uninstall }
    $app = Get-CimInstance -Query "SELECT * FROM Win32_Product WHERE Name like '%Adobe Acrobat Reader%'"
    if ($app) { $app | Invoke-CimMethod -MethodName Uninstall }

    Write-Host "Instalando nueva version"
    $msi = "\\ondoan.com\soporte$\Instal\Adobe Acrobat\AcrobatReader\AcroRead.msi"
    $mst = '"\\ondoan.com\soporte$\Instal\Adobe Acrobat\AcrobatReader\AcroRead.mst"'
    
    Start-Process msiexec.exe -ArgumentList "/i `"$msi`" TRANSFORMS=$mst /qn" -Wait
}

function Install-ZWCAD {
    Write-Host "------------------------------"
    Write-Host "ZWCAD 64 2020"
    Write-Host "------------------------------"

    Write-Host "Instalando ZWCad 2020..."
    $p = "\\ondoan.com\soporte$\Install\ZWCad\Deploy\Setup\setup.exe"
    Start-Process -FilePath $p -Wait
}

function Install-FortiConfig {
    Write-Host "------------------------------"
    Write-Host "FortiClient Configuracion (Desde Internet)"
    Write-Host "------------------------------"

    Write-Host "Bajando app"
    $tempZip = "c:\temp\forticlientc.zip"
    $tempDir = "c:\temp\forticlientc"
    Invoke-WebRequest 'https://legoan.ondoan.com/servicios/it/forticlientc.zip' -OutFile $tempZip -Verbose

    Write-Host "Descomprimiendo app"
    Expand-Archive -Path $tempZip -DestinationPath $tempDir -Force

    Write-Host "Instalando vpn ondoan"
    $p = "C:\Program Files\Fortinet\FortiClient\FCConfig.exe"
    if (!(Test-Path $p)) {
        $p = "C:\Program Files (x86)\Fortinet\FortiClient\FCConfig.exe"
    }
    
    if (Test-Path $p) {
        Start-Process -FilePath $p -ArgumentList "-m all -f `"$tempDir\config.conf`" -o import -i 1" -Wait
    }
    else {
        Write-Host "Error: FCConfig.exe no encontrado." -ForegroundColor Red
    }

    Write-Host "Limpiando temporales"
    Remove-Item -Path $tempZip -Force
    Remove-Item -Path $tempDir -Force -Recurse
    Write-Host "Acciones finalizadas"
}

function Install-FortiClient {
    Write-Host "------------------------------"
    Write-Host "FortiClient (Desde Internet)"
    Write-Host "------------------------------"

    Write-Host "Bajando app"
    $tempZip = "c:\temp\forticlient.zip"
    $tempDir = "c:\temp\forticlient"
    Invoke-WebRequest 'https://legoan.ondoan.com/servicios/it/forticlient.zip' -OutFile $tempZip -Verbose

    Write-Host "Descomprimiendo app"
    Expand-Archive -Path $tempZip -DestinationPath $tempDir -Force

    Write-Host "Instalando app"
    $p = "$tempDir\FortiClientVPN.exe"
    Start-Process -FilePath $p -ArgumentList "/quiet /norestart" -Wait

    Write-Host "Instalando vpn ondoan"
    $configExe = "C:\Program Files\Fortinet\FortiClient\FCConfig.exe"
    if (!(Test-Path $configExe)) {
        $configExe = "C:\Program Files (x86)\Fortinet\FortiClient\FCConfig.exe"
    }

    if (Test-Path $configExe) {
        Start-Process -FilePath $configExe -ArgumentList "-m all -f `"$tempDir\config.conf`" -o import -i 1" -Wait
    }

    Write-Host "Limpiando temporales"
    Remove-Item -Path $tempZip -Force
    Remove-Item -Path $tempDir -Force -Recurse
    Write-Host "Acciones finalizadas"
}

function Install-Foxit {
    Write-Host "------------------------------"
    Write-Host "Foxit Reader"
    Write-Host "------------------------------"
    
    $p = 'c:\temp\foxit.msi'
    Write-Host "Bajando app"
    Invoke-WebRequest 'https://www.foxit.com/downloads/latest.html?product=Foxit-Enterprise-Reader&platform=&version=&package_type=msi&language=Spanish&distID=' -OutFile $p -Verbose -UseBasicParsing

    Write-Host "Instalando nueva version"
    Start-Process msiexec.exe -ArgumentList "/i `"$p`" /quiet DESKTOPSHORTCUT=0 REBOOT=REALLYSUPPRESS INSTALLLEVEL=3" -Wait
    Remove-Item -Path $p -Force
}

function Install-Presto {
    Write-Host "------------------------------"
    Write-Host "Presto"
    Write-Host "------------------------------"
    $msi = "\\servidor2\soporte$\Install\Presto\PrestoSetup2303x64.msi"
    Start-Process msiexec.exe -ArgumentList "/i `"$msi`" /qn" -Wait
    Write-Host "Instalado"
}

function Install-Macrium {
    Write-Host "------------------------------"
    Write-Host "Macrium Reflect"
    Write-Host "------------------------------"
    $tempZip = 'c:\temp\mr.zip'
    $tempDir = 'c:\temp\mr'
    
    Write-Host "Bajando app"
    Invoke-WebRequest 'https://legoan.ondoan.com/servicios/it/mr.zip' -OutFile $tempZip -Verbose
    
    Write-Host "Descomprimiendo app"
    Expand-Archive -Path $tempZip -DestinationPath $tempDir -Force
    
    Write-Host "Instalando app"
    $p = "$tempDir\ReflectBin.exe"
    Start-Process -FilePath $p -Wait

    Write-Host "Limpiando temporales"
    Remove-Item -Path $tempZip -Force
    Remove-Item -Path $tempDir -Force -Recurse
    Write-Host "Acciones finalizadas"
}

function Install-Izenpe {
    Write-Host "------------------------------"
    Write-Host "Izempe"
    Write-Host "------------------------------"
    $p = "c:\temp\izenpe.exe"
    $u = 'https://www.izenpe.eus/contenidos/informacion/software_izenpe/es_def/adjuntos/izenpe-full-install64.exe'
    
    Write-Host "Bajando app"
    Invoke-WebRequest $u -OutFile $p -Verbose
    
    Write-Host "Instalando app"
    Start-Process -FilePath $p -ArgumentList "/s" -Wait
    
    Write-Host "Limpiando temporales"
    Remove-Item -Path $p -Force
    Write-Host "Acciones finalizadas"
}

function Install-Idazki {
    Write-Host "------------------------------"
    Write-Host "Idazki"
    Write-Host "------------------------------"
    $p = "c:\temp\idazki.msi"
    $u = 'https://www.izenpe.eus/contenidos/informacion/idazki_izenpe/es_def/adjuntos/idazki-desktop-win_64.msi'

    Write-Host "Bajando app"
    Invoke-WebRequest $u -OutFile $p -Verbose
    
    Write-Host "Instalando app"
    Start-Process msiexec.exe -ArgumentList "/i `"$p`" /quiet DESKTOPSHORTCUT=0 REBOOT=REALLYSUPPRESS INSTALLLEVEL=3" -Wait
    
    Write-Host "Limpiando temporales"
    Remove-Item -Path $p -Force
    Write-Host "Acciones finalizadas"
}

function Show-SoftwareMenu {
    Clear-Host
    Write-Host "Usuario: $env:UserName" -ForegroundColor Cyan
    Write-Host "Instalar..." -ForegroundColor Yellow
    Write-Host "1. Adobe Acrobat Reader DC"
    Write-Host "2. ZWCAD 64"
    Write-Host "3. FortiClient Configuracion (Desde Internet)"
    Write-Host "4. FortiClient (Desde Internet)"
    Write-Host "5. Foxit Reader"
    Write-Host "6. Presto"
    Write-Host "7. Macrium Reflect"
    Write-Host "8. Izempe"
    Write-Host "9. Idazki Desktop"
    Write-Host ""
    Write-Host "98. Reopen as admin"
    Write-Host "99. Open admin ps shell"
    Write-Host "0. Exit"
}

function Start-GbInstala {
    <#
    .SYNOPSIS
        Instala aplicaciones de Windows de forma interactiva.
    
    .DESCRIPTION
        Muestra un menú para seleccionar e instalar diversas aplicaciones corporativas.
    #>
    
    do {
        Show-SoftwareMenu
        $userInput = Read-Host "Selecciona una opcion (puedes usar comas para varias)"
        
        if ($userInput -eq "0" -or -not $userInput) { break }
        
        $options = $userInput.Split(",")
        foreach ($opt in $options) {
            $opt = $opt.Trim()
            switch ($opt) {
                "1" { Install-AdobeReader }
                "2" { Install-ZWCAD }
                "3" { Install-FortiConfig }
                "4" { Install-FortiClient }
                "5" { Install-Foxit }
                "6" { Install-Presto }
                "7" { Install-Macrium }
                "8" { Install-Izenpe }
                "9" { Install-Idazki }
                "98" { 
                    Write-Host "Reabriendo como admin..."
                    Start-Process powershell.exe -Verb RunAs
                }
                "99" {
                    Write-Host "Abriendo shell de admin..."
                    Start-Process powershell.exe -Verb RunAs
                }
                default { Write-Host "Opcion invalida: $opt" -ForegroundColor Red }
            }
            if ($opt -ne "0") { pause }
        }
    } while ($true)
}
