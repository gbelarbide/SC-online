<#
.SYNOPSIS
    


.DESCRIPTION
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/gbtools.psm1') | Invoke-Expression ; Get-HolaMundo
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Tools/desinstala.psm1') | Invoke-Expression ; Start-GbDesintala
   

.NOTES
    Version:        0.1.0
    Author:         Garikoitz Belarbide    
    Creation Date:  14/01/2026

#>

#region [Functions]-------------------------------------------------------------

Function Start-GbDesintala {
    (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Tools/desinstala.psm1') | Invoke-Expression
    Start-GbDesintala
}
#endregion
