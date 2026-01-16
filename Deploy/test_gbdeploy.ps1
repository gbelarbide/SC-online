# Script de prueba para Start-GbDeploy
# IMPORTANTE: Este script debe ejecutarse con privilegios elevados (como SYSTEM o Administrador)

Import-Module "c:\Users\gbelarbide\OneDrive - Ondoan\Desktop\Proyectos\SC online\Deploy\gbdeploy.psm1" -Force

Write-Host "`n=== Prueba de Start-GbDeploy ===" -ForegroundColor Yellow
Write-Host "Esta funcion crea una tarea programada que pregunta al usuario si desea instalar una aplicacion." -ForegroundColor White
Write-Host "Si el usuario acepta, se instala inmediatamente. Si rechaza, se vuelve a preguntar." -ForegroundColor White
Write-Host "En la ultima ejecucion, se muestra un aviso y se instala automaticamente." -ForegroundColor White

Write-Host "`n=== EJEMPLO DE USO ===" -ForegroundColor Cyan
Write-Host "Start-GbDeploy -Name 'office64' -N 5 -Every 60" -ForegroundColor Green
Write-Host "  - Pregunta al usuario 4 veces (cada hora)" -ForegroundColor Gray
Write-Host "  - En la 5ta vez, avisa y ejecuta automaticamente" -ForegroundColor Gray

Write-Host "`n=== PRUEBA CON PARAMETROS DE PRUEBA ===" -ForegroundColor Cyan
Write-Host "Vamos a crear una tarea de prueba con:" -ForegroundColor White
Write-Host "  - Name: 'TestApp'" -ForegroundColor Gray
Write-Host "  - N: 3 (3 intentos totales)" -ForegroundColor Gray
Write-Host "  - Every: 2 (cada 2 minutos)" -ForegroundColor Gray

$confirm = Read-Host "`n¿Desea ejecutar la prueba? (S/N)"

if ($confirm -eq 'S' -or $confirm -eq 's') {
    Write-Host "`nEjecutando Start-GbDeploy..." -ForegroundColor Green
    Start-GbDeploy -Name "TestApp" -N 3 -Every 2 -Verbose
    
    Write-Host "`n=== TAREA CREADA ===" -ForegroundColor Green
    Write-Host "La tarea 'Deploy_TestApp' ha sido creada en la carpeta \Ondoan\" -ForegroundColor White
    Write-Host "Puedes verificarla en el Programador de tareas" -ForegroundColor White
    
    Write-Host "`n=== COMPORTAMIENTO ESPERADO ===" -ForegroundColor Yellow
    Write-Host "1. En 2 minutos, aparecera un cuadro de dialogo preguntando si desea instalar TestApp" -ForegroundColor Gray
    Write-Host "2. Si haces clic en OK, se ejecutara la instalacion (fallara porque TestApp.psm1 no existe)" -ForegroundColor Gray
    Write-Host "3. Si haces clic en Cancelar, se programara el siguiente intento en 2 minutos" -ForegroundColor Gray
    Write-Host "4. En el 3er intento, solo mostrara un aviso y ejecutara automaticamente" -ForegroundColor Gray
    
    Write-Host "`n=== PARA CANCELAR LA PRUEBA ===" -ForegroundColor Red
    Write-Host "Remove-GbScheduledTask -TaskName 'Deploy_TestApp' -Force" -ForegroundColor Cyan
}
else {
    Write-Host "`nPrueba cancelada." -ForegroundColor Yellow
}

Write-Host "`n=== EJEMPLO REAL CON OFFICE 64 ===" -ForegroundColor Cyan
Write-Host "Para desplegar Office 64 con 5 intentos cada hora:" -ForegroundColor White
Write-Host "Start-GbDeploy -Name 'office64' -N 5 -Every 60" -ForegroundColor Green
