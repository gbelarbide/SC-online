# Script de prueba para Show-UserPrompt
# Importar el modulo
Import-Module "c:\Users\gbelarbide\OneDrive - Ondoan\Desktop\Proyectos\SC online\Deploy\gbdeploy.psm1" -Force

Write-Host "Probando Show-UserPrompt..." -ForegroundColor Cyan

# Prueba 1: Dialogo simple OK/Cancel
Write-Host "`nPrueba 1: Dialogo OK/Cancel" -ForegroundColor Yellow
$result1 = Show-UserPrompt -Message "Esta es una prueba de dialogo. ¿Desea continuar?" -Title "Prueba 1" -Verbose
Write-Host "Resultado: $result1" -ForegroundColor Green

# Prueba 2: Dialogo Yes/No con icono de advertencia
Write-Host "`nPrueba 2: Dialogo Yes/No con advertencia" -ForegroundColor Yellow
$result2 = Show-UserPrompt -Message "¿Desea proceder con esta accion?" -Title "Advertencia" -Buttons "YesNo" -Icon "Warning" -Verbose
Write-Host "Resultado: $result2" -ForegroundColor Green

Write-Host "`nPruebas completadas!" -ForegroundColor Cyan
