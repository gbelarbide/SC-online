function Get-InstalledApp {
    param ($appname)
    $32bit = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName | Where-Object { $_.DisplayName -match "^*$appname*" }
    $64bit = Get-ItemProperty 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName | Where-Object { $_.DisplayName -match "^*$appname*" }

    if ($64bit -eq "" -or $64bit.count -eq 0) {
        switch ($32bit.DisplayName.count) {
            0 { return $null }
            1 {               
                return $32bit
            }
            default { return $null }
        }
    }
    else {
        switch ($64bit.DisplayName.count) {
            0 { return $null }
            1 {               
                return $64bit.UninstallString                
            }
            default { return $null }
        }
    }
}

function Get-InstalledAppUS {
    param ($appname)

    $app = Get-InstalledApp($appname)

    if ($app -match "msiexec.exe") {
        return $app.UninstallString -replace 'msiexec.exe /i', 'msiexec.exe /x'
    }
}

function Get-InstalledApplist {
    $32bit = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName 
    $64bit = Get-ItemProperty 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName 
    $t = ($32bit + $64bit) 
    return $t | Where-Object { $_.UninstallString.Trim -ne "" } | Where-Object { $_.DisplayName.Trim -ne "" }
}

function Get-UninstallString {
    param ($appname)
    $32bit = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName | Where-Object { $_.DisplayName -match "^*$appname*" }
    $64bit = Get-ItemProperty 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName | Where-Object { $_.DisplayName -match "^*$appname*" }

    if ($64bit -eq "" -or $64bit.count -eq 0) {
        switch ($32bit.DisplayName.count) {
            0 { return $null }
            1 {
                if ($32bit -match "msiexec.exe") {
                    return $32bit.UninstallString -replace 'msiexec.exe /i', 'msiexec.exe /x'
                }
                else {
                    return $32bit.UninstallString 
                }
            }
            default { return $null }
        }
    }
    else {
        switch ($64bit.DisplayName.count) {
            0 { return $null }
            1 {
                if ($64bit -match "msiexec.exe") {
                    return $64bit.UninstallString -replace 'msiexec.exe /i', 'msiexec.exe /x'
                }
                else {
                    return $64bit.UninstallString 
                }
            }
            default { return $null }
        }
    }
}

function Invoke-Desinst {
    param ($App)

    Write-Host "Desinstalando "$app.DisplayName    

    $UninstallString = Get-UninstallString($App.DisplayName)

    if ($UninstallString -match "msiexec") {
        $UninstallString = $UninstallString + " /qn /norestart REBOOT=REALLYSUPPRESS"
        $UninstallString = $UninstallString -replace '{', ' "{'
        $UninstallString = $UninstallString -replace '}', '}"'

        Write-Host $UninstallString

        $exe = "msiexec.exe"
        $params = $UninstallString -replace "msiexec.exe ", ""   
        $sp = Start-Process -FilePath $exe -ArgumentList $params -Wait -PassThru 
        $sp.ExitCode 
       
    }
    elseif ($UninstallString -match "uninstaller.exe") {
        $UninstallString = $UninstallString + " /S"
        $s = $UninstallString.split("uninstaller.exe ")

        $exe = $s[0] + "uninstaller.exe"
        $params = $s[1]   
        Start-Process -FilePath $exe -ArgumentList $params -Wait -NoNewWindow 

    }
    else {
        Write-Host "no se reconoce el desinstalador, se ejecutara la sifuiente cadena"
        Write-Host $UninstallString
    }

    # Comprueba si la app sigue existiendo
    if (Get-InstalledApp($App.DisplayName)) {
        Write-Host $App.DisplayName " no se desinstalo" -ForegroundColor Red
    }
    else {
        Write-Host $App.DisplayName " se desinstalo adecuadamente" -ForegroundColor Green       
    }
}

function Start-GbDesintala {
    <#
    .SYNOPSIS
        Desinstala aplicaciones de Windows de forma interactiva.
    
    .DESCRIPTION
        Permite buscar, seleccionar y desinstalar aplicaciones instaladas en Windows.
        Soporta desinstalación de múltiples aplicaciones a la vez.
    
    .PARAMETER Filter
        Filtro opcional para buscar aplicaciones específicas.
    
    .EXAMPLE
        Start-GbDesintala
        Inicia el proceso de desinstalación interactivo.
    
    .EXAMPLE
        Start-GbDesintala -Filter "Adobe"
        Busca aplicaciones que contengan "Adobe" en el nombre.
    #>
    
    param (
        [string]$Filter
    )
    
    Clear-Host

    if (-not $Filter) {
        $f = Read-Host "filtro (enter para no filtrar)"
    }
    else {
        $f = $Filter
    }
      
    Clear-Host
    Write-Host "Listando, tarda unos segundos..."
    
    if ($f -eq "") {
        $apps = Get-InstalledApplist
    }
    else {
        $apps = Get-InstalledApplist | Where-Object { $_.DisplayName -match "^*$f*" }
    }  

    Clear-Host

    if ($apps -is [array]) {
        $i = 0
        foreach ($app in $apps) {
            $i += 1
            Write-Host $i " - " $app.DisplayName  " ("  $app.DisplayVersion  ")"
        }

        Write-Host
        $input = Read-Host "Selecciona una aplicacion o varias separadas con ','(coma)"
        Clear-Host
     
        if (-Not($input)) { return }     
     
        Write-Host "Se desistalaran las siguientes aplicaciones:"
        Write-Host "CTRL + C para cancelar"
        Write-Host

        $inputs = $input.Split(",")
        foreach ($inputItem in $inputs) {
            $app = $apps[$inputItem - 1]        
            Write-Host $app.DisplayName  " ("  $app.DisplayVersion  ")"        
        }

        Write-Host
        pause
        Clear-Host
     
        foreach ($inputItem in $inputs) {
            $app = $apps[$inputItem - 1]             
            Invoke-Desinst($app)       
        }
    }
    else {
        Write-Host "Se desistalara la siguiente aplicacion:"
        Write-Host "CTRL + C para cancelar"
        Write-Host

        Write-Host $apps.DisplayName 

        Write-Host
        pause
        Clear-Host

        Invoke-Desinst($apps)
    }

    Write-Host
    pause
}


