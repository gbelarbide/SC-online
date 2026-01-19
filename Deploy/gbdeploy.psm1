<#
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression
Start-GbDeploy -Name "Test" -N 5 -Every 1
#>

function Show-UserMessage {
    <#
    .SYNOPSIS
        Muestra un mensaje al usuario que tiene iniciada la sesión.
    
    .DESCRIPTION
        Esta función muestra un mensaje al usuario activo cuando se ejecuta en el contexto de SYSTEM.
        Utiliza el comando msg.exe para enviar mensajes a las sesiones de usuario activas.
    
    .PARAMETER Message
        El mensaje que se mostrará al usuario.
    
    .PARAMETER Title
        El título de la ventana del mensaje (opcional).
    
    .PARAMETER Timeout
        Tiempo en segundos antes de que el mensaje se cierre automáticamente (0 = sin timeout).
        Por defecto: 0 (el usuario debe cerrar el mensaje manualmente).
    
    .EXAMPLE
        Show-UserMessage -Message "La instalación se ha completado correctamente."
    
    .EXAMPLE
        Show-UserMessage -Message "El sistema se reiniciará en 5 minutos." -Timeout 60
    
    .EXAMPLE
        Show-UserMessage -Message "Actualización disponible" -Title "Notificación del Sistema" -Timeout 30
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Notificacion del Sistema",
        
        [Parameter(Mandatory = $false)]
        [int]$Timeout = 0
    )
    
    try {
        # Obtener todas las sesiones de usuario activas
        $sessions = query user 2>$null | Select-Object -Skip 1
        
        if (-not $sessions) {
            Write-Warning "No se encontraron sesiones de usuario activas."
            return
        }
        
        # Procesar cada línea de sesión
        foreach ($session in $sessions) {
            if ([string]::IsNullOrWhiteSpace($session)) {
                continue
            }
            
            # Extraer el ID de sesión (formato de query user puede variar)
            # Formato típico: USERNAME SESSIONNAME ID STATE IDLE TIME LOGON TIME
            $sessionInfo = $session -split '\s+' | Where-Object { $_ -ne '' }
            
            # El ID de sesión suele estar en la tercera posición (índice 2)
            # pero puede variar si hay un nombre de sesión
            $sessionId = $null
            foreach ($item in $sessionInfo) {
                if ($item -match '^\d+$') {
                    $sessionId = $item
                    break
                }
            }
            
            if ($sessionId) {
                # Construir el mensaje completo con el título
                $fullMessage = if ($Title) {
                    "[$Title] $Message"
                }
                else {
                    $Message
                }
                
                # Enviar mensaje a la sesión
                if ($Timeout -gt 0) {
                    msg.exe $sessionId /TIME:$Timeout $fullMessage 2>$null
                }
                else {
                    msg.exe $sessionId $fullMessage 2>$null
                }
                
                Write-Verbose "Mensaje enviado a la sesion ID: $sessionId"
            }
        }
        
        Write-Host "Mensaje enviado a todas las sesiones de usuario activas." -ForegroundColor Green
    }
    catch {
        Write-Error "Error al enviar el mensaje: $_"
    }
}

function Show-UserPrompt {
    <#
    .SYNOPSIS
        Muestra un cuadro de dialogo interactivo al usuario con botones de accion.
    
    .DESCRIPTION
        Esta funcion muestra un cuadro de dialogo interactivo al usuario activo cuando se ejecuta en el contexto de SYSTEM.
        Utiliza un script que se ejecuta en la sesion interactiva del usuario.
    
    .PARAMETER Message
        El mensaje que se mostrara al usuario.
    
    .PARAMETER Title
        El titulo de la ventana del dialogo.
    
    .PARAMETER Buttons
        Tipo de botones a mostrar. Valores validos:
        - OKCancel (OK y Cancelar)
        - YesNo (Si y No)
        Por defecto: OKCancel
    
    .PARAMETER Icon
        Icono a mostrar en el dialogo. Valores validos:
        - Information (Informacion)
        - Question (Pregunta)
        - Warning (Advertencia)
        - Error (Error)
        Por defecto: Question
    
    .PARAMETER TimeoutSeconds
        Tiempo en segundos antes de que el dialogo se cierre automaticamente.
        Si se alcanza el timeout, se considera como "Cancelar".
        0 = sin timeout
        Por defecto: 0
    
    .EXAMPLE
        $result = Show-UserPrompt -Message "¿Desea continuar con la instalacion?" -Title "Confirmacion"
        if ($result -eq "OK") {
            Write-Host "Usuario acepto continuar"
        }
    
    .EXAMPLE
        $result = Show-UserPrompt -Message "¿Desea reiniciar el equipo ahora?" -Buttons "YesNo" -Icon "Warning"
        if ($result -eq "Yes") {
            Restart-Computer -Force
        }
    
    .OUTPUTS
        String - Devuelve la respuesta del usuario: "OK", "Cancel", "Yes", "No", o "Timeout"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Confirmacion del Sistema",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("OKCancel", "YesNo")]
        [string]$Buttons = "OKCancel",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Information", "Question", "Warning", "Error")]
        [string]$Icon = "Question",
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 0
    )
    
    try {
        # Detectar si estamos ejecutando como SYSTEM o como usuario
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $isSystem = $currentUser.IsSystem
        
        Write-Verbose "Ejecutando como: $($currentUser.Name), IsSystem: $isSystem"
        
        # Crear archivo temporal para el resultado en una ubicacion accesible por todos
        # Usar ProgramData en lugar de TEMP para evitar problemas de permisos
        $tempFolder = if ($isSystem) { "C:\ProgramData\Temp" } else { $env:TEMP }
        
        # Crear la carpeta si no existe
        if (-not (Test-Path $tempFolder)) {
            New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
        }
        
        $resultPath = "$tempFolder\UserPrompt_Result_$(Get-Random).txt"
        
        # Escapar caracteres especiales para VBScript
        # Reemplazar saltos de linea con el codigo VBScript apropiado
        $escapedMessage = $Message -replace '"', '""' -replace '\r?\n', '" & vbCrLf & "'
        $escapedTitle = $Title -replace '"', '""'
        
        # Mapear tipos de botones a valores VBScript MsgBox
        $buttonValue = if ($Buttons -eq "YesNo") { 4 } else { 1 }
        
        # Mapear iconos a valores VBScript MsgBox
        $iconValue = switch ($Icon) {
            "Error" { 16 }
            "Question" { 32 }
            "Warning" { 48 }
            "Information" { 64 }
            default { 32 }
        }
        
        $style = $buttonValue + $iconValue + 4096  # 4096 = vbSystemModal para que aparezca al frente
        
        # Crear script VBScript que se ejecutara en la sesion del usuario
        $vbsContent = @"
Dim objShell, result, fso, file
Set objShell = CreateObject("WScript.Shell")

result = MsgBox("$escapedMessage", $style, "$escapedTitle")

' Mapear resultado a texto
Dim resultText
Select Case result
    Case 1
        resultText = "OK"
    Case 2
        resultText = "Cancel"
    Case 6
        resultText = "Yes"
    Case 7
        resultText = "No"
    Case Else
        resultText = "Cancel"
End Select

' Escribir resultado en archivo
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("$($resultPath -replace '\\', '\\')", True)
file.WriteLine resultText
file.Close

WScript.Quit
"@
        
        # Guardar el script VBScript en la ubicacion accesible
        $vbsPath = "$tempFolder\UserPrompt_$(Get-Random).vbs"
        Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII -Force
        
        if (-not $isSystem) {
            # Si NO estamos ejecutando como SYSTEM, ejecutar directamente
            Write-Verbose "Ejecutando VBScript directamente en la sesion actual del usuario"
            $process = Start-Process -FilePath "wscript.exe" -ArgumentList "`"$vbsPath`"" -Wait -PassThru -WindowStyle Hidden
        }
        else {
            # Si estamos ejecutando como SYSTEM, necesitamos ejecutar en la sesion del usuario
            Write-Verbose "Ejecutando como SYSTEM, buscando sesion interactiva del usuario"
            
            # Obtener la sesion interactiva del usuario
            $sessionId = (Get-Process -Name "explorer" -ErrorAction SilentlyContinue | Select-Object -First 1).SessionId
            
            if ($null -eq $sessionId) {
                Write-Warning "No se encontro una sesion interactiva de usuario."
                return "Cancel"
            }
            
            Write-Verbose "Sesion interactiva encontrada: $sessionId"
            
            # Usar PsExec si esta disponible
            $psExecPath = "C:\Windows\System32\PsExec.exe"
            
            if (Test-Path $psExecPath) {
                # Ejecutar con PsExec en la sesion interactiva
                Write-Verbose "Usando PsExec para ejecutar en la sesion $sessionId"
                $null = Start-Process -FilePath $psExecPath -ArgumentList "-accepteula -s -i $sessionId wscript.exe `"$vbsPath`"" -WindowStyle Hidden -PassThru
            }
            else {
                # Metodo alternativo: usar WMI para crear proceso en la sesion del usuario
                Write-Verbose "PsExec no disponible, usando WMI/PowerShell Remoting"
                
                # Crear un script que use schtasks de forma mas directa
                $batchPath = "$tempFolder\UserPrompt_$(Get-Random).bat"
                $batchContent = "@echo off`r`nwscript.exe `"$vbsPath`""
                Set-Content -Path $batchPath -Value $batchContent -Encoding ASCII -Force
                
                # Obtener el usuario de la sesion
                $sessionUser = (query user | Select-String -Pattern "^\s*\S+\s+console\s+$sessionId" | ForEach-Object {
                        ($_ -split '\s+')[1]
                    }) | Select-Object -First 1
                
                if (-not $sessionUser) {
                    # Intentar obtener el usuario de otra forma
                    $sessionUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                    if ($sessionUser) {
                        $sessionUser = $sessionUser.Split('\')[-1]
                    }
                }
                
                Write-Verbose "Usuario de sesion: $sessionUser"
                
                if ($sessionUser) {
                    # Crear tarea que se ejecute inmediatamente y se elimine
                    $taskName = "UserPrompt_$(Get-Random)"
                    $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>User Prompt Dialog</Description>
  </RegistrationInfo>
  <Triggers />
  <Principals>
    <Principal id="Author">
      <UserId>$sessionUser</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>wscript.exe</Command>
      <Arguments>"$vbsPath"</Arguments>
    </Exec>
  </Actions>
</Task>
"@
                    
                    $taskXmlPath = "$tempFolder\UserPrompt_$(Get-Random).xml"
                    Set-Content -Path $taskXmlPath -Value $taskXml -Encoding Unicode -Force
                    
                    # Registrar y ejecutar la tarea
                    schtasks /create /tn $taskName /xml $taskXmlPath /f | Out-Null
                    Start-Sleep -Milliseconds 500
                    schtasks /run /tn $taskName | Out-Null
                    
                    # Guardar el nombre de la tarea para limpiarla despues
                    $script:taskNameToClean = $taskName
                    $script:taskXmlPathToClean = $taskXmlPath
                    $script:batchPathToClean = $batchPath
                }
                else {
                    Write-Warning "No se pudo determinar el usuario de la sesion interactiva"
                    return "Cancel"
                }
            }
        }
        
        # Esperar a que el usuario responda o se alcance el timeout
        $maxWait = if ($TimeoutSeconds -gt 0) { $TimeoutSeconds + 10 } else { 300 }
        $elapsed = 0
        $checkInterval = 1
        
        Write-Verbose "Esperando respuesta del usuario (max: $maxWait segundos)..."
        
        while ($elapsed -lt $maxWait) {
            Start-Sleep -Seconds $checkInterval
            $elapsed += $checkInterval
            
            if (Test-Path $resultPath) {
                Write-Verbose "Archivo de resultado encontrado"
                break
            }
        }
        
        # Leer el resultado
        $userResponse = "Cancel"
        
        if (Test-Path $resultPath) {
            $resultValue = (Get-Content $resultPath -Raw).Trim()
            
            if (-not [string]::IsNullOrWhiteSpace($resultValue)) {
                $userResponse = $resultValue
            }
        }
        else {
            Write-Warning "No se recibio respuesta del usuario (timeout o error)."
        }
        
        # Limpiar archivos temporales y tarea programada
        try {
            if (Test-Path $vbsPath) { Remove-Item $vbsPath -Force -ErrorAction SilentlyContinue }
            if (Test-Path $resultPath) { Remove-Item $resultPath -Force -ErrorAction SilentlyContinue }
            
            # Eliminar tarea programada si se creo
            if ($taskName) {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Ignorar errores de limpieza
        }
        
        Write-Verbose "Respuesta del usuario: $userResponse"
        return $userResponse
    }
    catch {
        Write-Error "Error al mostrar el dialogo: $_"
        return "Cancel"
    }
}

function New-GbScheduledTask {
    <#
    .SYNOPSIS
        Crea una tarea programada que ejecuta un comando de PowerShell periodicamente.
    
    .DESCRIPTION
        Esta funcion crea una tarea programada en la carpeta \Ondoan\ que ejecuta un comando
        de PowerShell a intervalos especificados.
    
    .PARAMETER TaskName
        Nombre de la tarea programada (sin la ruta de carpeta).
    
    .PARAMETER ScriptBlock
        Bloque de script de PowerShell a ejecutar.
    
    .PARAMETER IntervalMinutes
        Intervalo en minutos entre ejecuciones (por defecto: 60 minutos).
    
    .PARAMETER RunAsSystem
        Si se especifica, la tarea se ejecuta como SYSTEM. Si no, se ejecuta como el usuario actual.
    
    .PARAMETER StartTime
        Hora de inicio de la tarea (por defecto: ahora).
    
    .PARAMETER Description
        Descripcion de la tarea (opcional).
    
    .EXAMPLE
        New-GbScheduledTask -TaskName "MiTarea" -ScriptBlock { Write-Host "Hola" } -IntervalMinutes 30
    
    .EXAMPLE
        $script = { Get-Process | Out-File C:\logs\procesos.txt }
        New-GbScheduledTask -TaskName "LogProcesos" -ScriptBlock $script -IntervalMinutes 15 -RunAsSystem
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskName,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [int]$IntervalMinutes = 60,
        
        [Parameter(Mandatory = $false)]
        [switch]$RunAsSystem,
        
        [Parameter(Mandatory = $false)]
        [datetime]$StartTime = (Get-Date),
        
        [Parameter(Mandatory = $false)]
        [string]$Description = "Tarea creada por gbdeploy"
    )
    
    try {
        # Crear la carpeta Ondoan si no existe
        $taskPath = "\Ondoan\"
        
        # Verificar si la tarea ya existe
        $existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $taskPath -ErrorAction SilentlyContinue
        if ($existingTask) {
            Write-Warning "La tarea '$TaskName' ya existe en la carpeta Ondoan. Sera reemplazada."
            Unregister-ScheduledTask -TaskName $TaskName -TaskPath $taskPath -Confirm:$false
        }
        
        # Convertir el scriptblock a string y codificarlo en base64
        $scriptString = $ScriptBlock.ToString()
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptString)
        $encodedCommand = [Convert]::ToBase64String($bytes)
        
        # Crear la accion
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -EncodedCommand $encodedCommand"
        
        # Crear el trigger (repetir cada X minutos)
        $trigger = New-ScheduledTaskTrigger -Once -At $StartTime -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) -RepetitionDuration ([TimeSpan]::MaxValue)
        
        # Crear el principal (usuario o SYSTEM)
        if ($RunAsSystem) {
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        }
        else {
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Highest
        }
        
        # Configuracion de la tarea
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false -MultipleInstances IgnoreNew
        
        # Registrar la tarea
        $task = Register-ScheduledTask -TaskName $TaskName -TaskPath $taskPath -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $Description -Force
        
        Write-Host "Tarea programada '$TaskName' creada exitosamente en la carpeta Ondoan" -ForegroundColor Green
        Write-Host "  - Intervalo: cada $IntervalMinutes minutos" -ForegroundColor Cyan
        Write-Host "  - Usuario: $(if ($RunAsSystem) { 'SYSTEM' } else { $currentUser })" -ForegroundColor Cyan
        Write-Host "  - Proxima ejecucion: $StartTime" -ForegroundColor Cyan
        
        return $task
    }
    catch {
        Write-Error "Error al crear la tarea programada: $_"
        return $null
    }
}

function Remove-GbScheduledTask {
    <#
    .SYNOPSIS
        Elimina una tarea programada de la carpeta Ondoan.
    
    .DESCRIPTION
        Esta funcion elimina una tarea programada previamente creada en la carpeta \Ondoan\.
    
    .PARAMETER TaskName
        Nombre de la tarea programada a eliminar.
    
    .PARAMETER Force
        Si se especifica, no solicita confirmacion antes de eliminar.
    
    .EXAMPLE
        Remove-GbScheduledTask -TaskName "MiTarea"
    
    .EXAMPLE
        Remove-GbScheduledTask -TaskName "MiTarea" -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    try {
        $taskPath = "\Ondoan\"
        
        # Verificar si la tarea existe
        $task = Get-ScheduledTask -TaskName $TaskName -TaskPath $taskPath -ErrorAction SilentlyContinue
        
        if (-not $task) {
            Write-Warning "La tarea '$TaskName' no existe en la carpeta Ondoan."
            return $false
        }
        
        # Eliminar la tarea
        if ($Force) {
            Unregister-ScheduledTask -TaskName $TaskName -TaskPath $taskPath -Confirm:$false
        }
        else {
            Unregister-ScheduledTask -TaskName $TaskName -TaskPath $taskPath
        }
        
        Write-Host "Tarea programada '$TaskName' eliminada exitosamente" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Error al eliminar la tarea programada: $_"
        return $false
    }
}

function Start-GbDeploy {
    <#
    .SYNOPSIS
        Gestiona el despliegue de una aplicacion mediante prompts programados al usuario.
    
    .DESCRIPTION
        Esta funcion crea una tarea programada que pregunta al usuario si desea instalar
        una aplicacion. Si el usuario acepta, se ejecuta la instalacion inmediatamente.
        Si rechaza, se vuelve a preguntar en el siguiente intervalo.
        En la ultima ejecucion, se muestra un aviso y se instala automaticamente.
    
    .PARAMETER Name
        Nombre de la aplicacion/modulo a desplegar.
    
    .PARAMETER N
        Numero total de intentos antes de la instalacion forzada.
    
    .PARAMETER Every
        Intervalo en minutos entre cada intento.
    
    .EXAMPLE
        Start-GbDeploy -Name "office64" -N 5 -Every 60
        # Pregunta 4 veces cada hora, en la 5ta vez instala automaticamente
    
    .EXAMPLE
        Start-GbDeploy -Name "MyApp" -N 3 -Every 30
        # Pregunta 2 veces cada 30 minutos, en la 3ra vez instala automaticamente
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [int]$N,
        
        [Parameter(Mandatory = $true)]
        [int]$Every
    )
    
    try {
        $taskName = "Deploy_$Name"
        $taskPath = "\Ondoan\"
        
        # Intentar obtener la tarea existente para leer los metadatos
        $existingTask = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue
        
        $currentAttempt = 1
        
        if ($existingTask) {
            # Leer metadatos de la descripcion de la tarea
            try {
                $metadata = $existingTask.Description | ConvertFrom-Json
                $currentAttempt = $metadata.CurrentAttempt
                
                Write-Verbose "Tarea existente encontrada. Intento actual: $currentAttempt de $N"
            }
            catch {
                Write-Warning "No se pudieron leer los metadatos de la tarea. Asumiendo primer intento."
                $currentAttempt = 1
            }
        }
        else {
            Write-Verbose "Primera ejecucion. Creando tarea programada."
        }
        
        # Determinar si es la ultima ejecucion
        $isLastAttempt = ($currentAttempt -ge $N)
        
        if ($isLastAttempt) {
            # ULTIMA EJECUCION: Mostrar aviso y ejecutar
            Write-Host "=== ULTIMA EJECUCION ===" -ForegroundColor Red
            Write-Host "Se instalara $Name automaticamente en 5 minutos" -ForegroundColor Yellow
            
            # Mostrar mensaje al usuario
            Show-UserMessage -Message "La aplicacion $Name se instalara en 5 minutos. Por favor, guarde su trabajo." -Title "Instalacion Programada"
            
            # Esperar 5 minutos
            Write-Verbose "Esperando 5 minutos antes de la instalacion..."
            Start-Sleep -Seconds 300
            
            # Ejecutar la instalacion
            Write-Host "Ejecutando instalacion de $Name..." -ForegroundColor Green
            $deployResult = Invoke-GbDeployment -Name $Name
            
            # Guardar resultado en el registro
            Save-DeploymentResult -AppName $Name -Result $deployResult | Out-Null
            
            # Eliminar la tarea programada
            Write-Verbose "Eliminando tarea programada..."
            Remove-GbScheduledTask -TaskName $taskName -Force -ErrorAction SilentlyContinue
            
            if ($deployResult.Success) {
                Write-Host "Despliegue de $Name completado exitosamente." -ForegroundColor Green
            }
            else {
                Write-Warning "El despliegue de $Name finalizo con errores: $($deployResult.Message)"
            }
        }
        else {
            # EJECUCIONES INTERMEDIAS: Preguntar al usuario
            Write-Host "=== INTENTO $currentAttempt de $N ===" -ForegroundColor Cyan
            
            # Preguntar al usuario
            $response = Show-UserPrompt -Message "¿Desea instalar $Name ahora?`n`nSi selecciona 'Cancelar', se le volvera a preguntar en $Every minutos.`n`nIntentos restantes: $($N - $currentAttempt)" -Title "Instalacion de $Name" -Buttons "OKCancel" -Icon "Question"
            
            if ($response -eq "OK") {
                # Usuario acepto: Ejecutar instalacion y eliminar tarea
                Write-Host "Usuario acepto la instalacion." -ForegroundColor Green
                
                # Ejecutar la instalacion
                Write-Host "Ejecutando instalacion de $Name..." -ForegroundColor Green
                $deployResult = Invoke-GbDeployment -Name $Name
                
                # Guardar resultado en el registro
                Save-DeploymentResult -AppName $Name -Result $deployResult | Out-Null
                
                # Eliminar la tarea programada
                Write-Verbose "Eliminando tarea programada..."
                Remove-GbScheduledTask -TaskName $taskName -Force -ErrorAction SilentlyContinue
                
                if ($deployResult.Success) {
                    Write-Host "Despliegue de $Name completado exitosamente." -ForegroundColor Green
                }
                else {
                    Write-Warning "El despliegue de $Name finalizo con errores: $($deployResult.Message)"
                }
            }
            else {
                # Usuario rechazo: Programar siguiente ejecucion
                Write-Host "Usuario rechazo la instalacion. Programando siguiente intento..." -ForegroundColor Yellow
                
                # Incrementar contador de intentos
                $nextAttempt = $currentAttempt + 1
                
                # Crear metadatos actualizados
                $metadata = @{
                    CurrentAttempt  = $nextAttempt
                    TotalAttempts   = $N
                    IntervalMinutes = $Every
                    AppName         = $Name
                    LastAttempt     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                } | ConvertTo-Json -Compress
                
                # Eliminar tarea existente si existe
                if ($existingTask) {
                    Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false -ErrorAction SilentlyContinue
                }
                
                # Crear el script que se ejecutara en la tarea
                $scriptBlock = [scriptblock]::Create(@"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression
Start-GbDeploy -Name '$Name' -N $N -Every $Every
"@)
                
                # Calcular hora de siguiente ejecucion
                $nextRunTime = (Get-Date).AddMinutes($Every)
                
                # Crear nueva tarea programada
                Write-Verbose "Creando tarea para siguiente ejecucion en $nextRunTime"
                
                $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$($scriptBlock.ToString())`""
                
                # Crear dos triggers: uno por tiempo y otro al logon
                # Esto asegura que si el ordenador se apaga, la tarea se ejecute al iniciar sesion
                $triggerTime = New-ScheduledTaskTrigger -Once -At $nextRunTime
                $triggerLogon = New-ScheduledTaskTrigger -AtLogOn
                
                $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                
                # Registrar con ambos triggers
                Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Trigger @($triggerTime, $triggerLogon) -Principal $principal -Settings $settings -Description $metadata -Force | Out-Null
                
                Write-Host "Siguiente intento programado para: $nextRunTime (o al iniciar sesion)" -ForegroundColor Cyan
            }
        }
    }
    catch {
        Write-Error "Error en Start-GbDeploy: $_"
        Write-Error $_.ScriptStackTrace
    }
}

function Invoke-GbDeployment {
    <#
    .SYNOPSIS
        Ejecuta el despliegue de una aplicacion descargando su modulo.
    
    .DESCRIPTION
        Funcion interna que descarga y ejecuta el modulo de despliegue de una aplicacion.
    
    .PARAMETER Name
        Nombre del modulo a desplegar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    
    try {
        Write-Host "Descargando modulo de despliegue para $Name..." -ForegroundColor Cyan
        
        # Convertir nombre a minusculas (GitHub es case-sensitive)
        $moduleName = $Name.ToLower()
        
        # Descargar el modulo desde GitHub
        $url = "https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/$moduleName.psm1"
        Write-Host "URL: $url" -ForegroundColor Yellow
        $moduleContent = (new-object Net.WebClient).DownloadString($url)
        
        # Ejecutar el modulo
        Invoke-Expression $moduleContent
        
        # Intentar ejecutar la funcion Start-Deploy o Start-Install si existe
        $deployResult = $null
        
        if (Get-Command Start-Deploy -ErrorAction SilentlyContinue) {
            Write-Host "Ejecutando Start-Deploy..." -ForegroundColor Green
            $deployResult = Start-Deploy
        }
        elseif (Get-Command Start-Install -ErrorAction SilentlyContinue) {
            Write-Host "Ejecutando Start-Install..." -ForegroundColor Green
            Start-Install
            # Start-Install no devuelve resultado, asumir exito si no hay excepcion
            $deployResult = [PSCustomObject]@{
                Success = $true
                Message = "Instalacion completada (Start-Install)"
            }
        }
        else {
            Write-Warning "No se encontro la funcion Start-Deploy ni Start-Install en el modulo $Name"
            $deployResult = [PSCustomObject]@{
                Success = $false
                Message = "No se encontro funcion de instalacion en el modulo"
            }
        }
        
        return $deployResult
    }
    catch {
        Write-Error "Error al ejecutar el despliegue de $Name : $_"
        return [PSCustomObject]@{
            Success = $false
            Message = "Error: $($_.Exception.Message)"
        }
    }
}

function Save-DeploymentResult {
    <#
    .SYNOPSIS
        Guarda el resultado del despliegue en el registro de Windows.
    
    .DESCRIPTION
        Guarda el resultado del despliegue en formato JSON en la clave de registro
        HKLM:\SOFTWARE\ondoan\Deployments\<AppName>
    
    .PARAMETER AppName
        Nombre de la aplicacion desplegada
    
    .PARAMETER Result
        Objeto con el resultado del despliegue
    
    .EXAMPLE
        Save-DeploymentResult -AppName "office64" -Result $deployResult
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Result
    )
    
    try {
        # Crear la clave base si no existe
        $basePath = "HKLM:\SOFTWARE\ondoan"
        if (-not (Test-Path $basePath)) {
            New-Item -Path $basePath -Force | Out-Null
            Write-Verbose "Clave de registro creada: $basePath"
        }
        
        # Crear subclave para deployments
        $deploymentsPath = "$basePath\Deployments"
        if (-not (Test-Path $deploymentsPath)) {
            New-Item -Path $deploymentsPath -Force | Out-Null
            Write-Verbose "Clave de registro creada: $deploymentsPath"
        }
        
        # Crear o actualizar la clave para esta aplicacion
        $appPath = "$deploymentsPath\$AppName"
        if (-not (Test-Path $appPath)) {
            New-Item -Path $appPath -Force | Out-Null
            Write-Verbose "Clave de registro creada: $appPath"
        }
        
        # Preparar datos para guardar
        $resultData = @{
            AppName   = $AppName
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Success   = $Result.Success
            Result    = $Result
        }
        
        # Convertir a JSON
        $jsonResult = $resultData | ConvertTo-Json -Depth 10 -Compress
        
        # Guardar en el registro
        Set-ItemProperty -Path $appPath -Name "LastDeployment" -Value $jsonResult -Type String
        Set-ItemProperty -Path $appPath -Name "LastDeploymentDate" -Value (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") -Type String
        Set-ItemProperty -Path $appPath -Name "Success" -Value $Result.Success.ToString() -Type String
        
        Write-Host "Resultado guardado en el registro: $appPath" -ForegroundColor Green
        Write-Verbose "JSON guardado: $jsonResult"
        
        return $true
    }
    catch {
        Write-Error "Error al guardar el resultado en el registro: $_"
        return $false
    }
}

# Exportar las funciones
#Export-ModuleMember -Function Show-UserMessage, Show-UserPrompt, New-GbScheduledTask, Remove-GbScheduledTask, Start-GbDeploy
