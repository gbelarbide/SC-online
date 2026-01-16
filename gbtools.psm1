<#
.SYNOPSIS
    


.DESCRIPTION
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/gbtools.psm1') | Invoke-Expression ; Get-GbHelp
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Tools/desinstala.psm1') | Invoke-Expression ; Start-GbDesintala
   

.NOTES
    Version:        0.1.0
    Author:         Garikoitz Belarbide    
    Creation Date:  14/01/2026

#>

#region [Functions]-------------------------------------------------------------
Function Get-GbHelp {
    <#
    .SYNOPSIS
        Muestra las funciones disponibles en el módulo gbtools
    
    .DESCRIPTION
        Muestra una lista de todas las funciones disponibles en el módulo gbtools con una breve descripción de cada una
    
    .EXAMPLE
        Get-GbHelp
    #>
    
    Write-Host "`n=== Funciones Disponibles en gbtools ===" -ForegroundColor Cyan
    Write-Host ""
    
    $functions = @(
        [PSCustomObject]@{
            Funcion     = "Get-GbHelp"
            Descripcion = "Muestra esta ayuda con las funciones disponibles"
        },
        [PSCustomObject]@{
            Funcion     = "Start-GbDesintala"
            Descripcion = "Descarga y ejecuta el módulo de desinstalación de aplicaciones"
        }
    )
    
    $functions | Format-Table -AutoSize -Wrap
    
    Write-Host "Para más información sobre una función específica, usa: Get-Help <Nombre-Función> -Detailed" -ForegroundColor Yellow
    Write-Host ""
}


Function Start-GbDesintala {
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Tools/desinstala.psm1') | Invoke-Expression
    Start-GbDesintala
}
#endregion
