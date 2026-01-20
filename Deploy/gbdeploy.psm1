<#
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression; Start-GbDeploy -Name "Test"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression; Start-GbDeploy -Name "office64"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression; Get-DeploymentLog -AppName "office64"
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
        Si se alcanza el timeout, se considera como "OK" (Aceptar).
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
        String - Devuelve la respuesta del usuario: "OK", "Cancel", "Yes", "No"
        Nota: Si hay timeout, se devuelve "OK" automaticamente
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
        
        # Decidir si usar HTA (con countdown) o VBScript (sin countdown)
        $useHTA = ($TimeoutSeconds -gt 0)
        
        if ($useHTA) {
            # Usar HTA con cuenta atrás
            Write-Verbose "Usando HTA con timeout de $TimeoutSeconds segundos"
            
            # Escapar mensaje para HTML
            $escapedMessage = $Message -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;' -replace '\r?\n', '<br>'
            $escapedTitle = $Title -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
            
            # Determinar botones HTML
            $buttonsHtml = if ($Buttons -eq "YesNo") {
                @"
                <button onclick="saveResult('Yes')" style="padding: 10px 30px; margin: 5px; font-size: 14px;">Sí</button>
                <button onclick="saveResult('No')" style="padding: 10px 30px; margin: 5px; font-size: 14px;">No</button>
"@
            }
            else {
                @"
                <button onclick="saveResult('OK')" style="padding: 10px 30px; margin: 5px; font-size: 14px;">OK</button>
                <button onclick="saveResult('Cancel')" style="padding: 10px 30px; margin: 5px; font-size: 14px;">Cancelar</button>
"@
            }
            
            # Crear contenido HTA
            $htaContent = @"
<html>
<head>
    <title>$escapedTitle</title>
    <HTA:APPLICATION
        APPLICATIONNAME="UserPrompt"
        BORDER="dialog"
        BORDERSTYLE="normal"
        CAPTION="yes"
        MAXIMIZEBUTTON="no"
        MINIMIZEBUTTON="no"
        SCROLL="no"
        SHOWINTASKBAR="yes"
        SINGLEINSTANCE="yes"
        SYSMENU="yes"
        WINDOWSTATE="normal"
    />
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f0f0f0;
        }
        .container {
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 500px;
            margin: 0 auto;
        }
        .message {
            margin-bottom: 20px;
            font-size: 14px;
            line-height: 1.6;
            color: #333;
            max-height: 250px;
            overflow-y: auto;
            padding-right: 10px;
        }
        .countdown {
            font-size: 24px;
            font-weight: bold;
            color: #0078d4;
            text-align: center;
            margin: 20px 0;
            padding: 15px;
            background-color: #f8f8f8;
            border-radius: 5px;
        }
        .buttons {
            text-align: center;
            margin-top: 20px;
        }
        button {
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 500;
        }
        button:hover {
            background-color: #005a9e;
        }
        button:last-child {
            background-color: #6c757d;
        }
        button:last-child:hover {
            background-color: #5a6268;
        }
    </style>
    <script>
        var timeLeft = $TimeoutSeconds;
        var resultPath = "$($resultPath -replace '\\', '\\')";
        
        function formatTime(seconds) {
            var hours = Math.floor(seconds / 3600);
            var minutes = Math.floor((seconds % 3600) / 60);
            var secs = seconds % 60;
            
            if (hours > 0) {
                return hours + ':' + pad(minutes) + ':' + pad(secs);
            } else {
                return minutes + ':' + pad(secs);
            }
        }
        
        function pad(num) {
            return (num < 10 ? '0' : '') + num;
        }
        
        function updateCountdown() {
            if (timeLeft <= 0) {
                saveResult('OK');
                window.close();
                return;
            }
            
            document.getElementById('countdown').innerText = 'Tiempo restante: ' + formatTime(timeLeft);
            timeLeft--;
        }
        
        function saveResult(result) {
            try {
                var fso = new ActiveXObject('Scripting.FileSystemObject');
                var file = fso.CreateTextFile(resultPath, true);
                file.WriteLine(result);
                file.Close();
            } catch(e) {
                // Error al guardar, pero cerrar de todos modos
            }
            window.close();
        }
        
        window.onload = function() {
            // Centrar ventana con mayor altura
            window.resizeTo(550, 500);
            window.moveTo((screen.width - 550) / 2, (screen.height - 500) / 2);
            
            // Iniciar cuenta atrás
            updateCountdown();
            setInterval(updateCountdown, 1000);
        };
    </script>
</head>
<body>
    <div class="container">
        <div class="message">$escapedMessage</div>
        <div id="countdown" class="countdown">Tiempo restante: $(if ($TimeoutSeconds -ge 3600) { [Math]::Floor($TimeoutSeconds / 3600).ToString() + ':' } else { '' })$([Math]::Floor(($TimeoutSeconds % 3600) / 60).ToString('00')):$($TimeoutSeconds % 60).ToString('00')</div>
        <div class="buttons">
            $buttonsHtml
        </div>
    </div>
</body>
</html>
"@
            
            # Guardar el HTA
            $htaPath = "$tempFolder\UserPrompt_$(Get-Random).hta"
            Set-Content -Path $htaPath -Value $htaContent -Encoding UTF8 -Force
            $scriptPath = $htaPath
            $scriptExecutable = "mshta.exe"
        }
        else {
            # Usar VBScript tradicional (sin timeout)
            Write-Verbose "Usando VBScript sin timeout"
            
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
            $scriptPath = $vbsPath
            $scriptExecutable = "wscript.exe"
        }
        
        if (-not $isSystem) {
            # Si NO estamos ejecutando como SYSTEM, ejecutar directamente
            Write-Verbose "Ejecutando script directamente en la sesion actual del usuario"
            $process = Start-Process -FilePath $scriptExecutable -ArgumentList "`"$scriptPath`"" -Wait -PassThru -WindowStyle Hidden
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
                $null = Start-Process -FilePath $psExecPath -ArgumentList "-accepteula -s -i $sessionId $scriptExecutable `"$scriptPath`"" -WindowStyle Hidden -PassThru
            }
            else {
                # Metodo alternativo: usar WMI para crear proceso en la sesion del usuario
                Write-Verbose "PsExec no disponible, usando WMI/PowerShell Remoting"
                
                # Crear un script que use schtasks de forma mas directa
                $batchPath = "$tempFolder\UserPrompt_$(Get-Random).bat"
                $batchContent = "@echo off`r`n$scriptExecutable `"$scriptPath`""
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
      <Command>$scriptExecutable</Command>
      <Arguments>"$scriptPath"</Arguments>
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
        # Si hay timeout, se considera como aceptar (OK)
        $userResponse = "OK"
        
        if (Test-Path $resultPath) {
            $resultValue = (Get-Content $resultPath -Raw).Trim()
            
            if (-not [string]::IsNullOrWhiteSpace($resultValue)) {
                $userResponse = $resultValue
            }
        }
        else {
            Write-Warning "No se recibio respuesta del usuario (timeout). Se considera como aceptar."
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

function Add-DeploymentLog {
    <#
    .SYNOPSIS
        Registra eventos de despliegue en el registro de Windows.
    
    .DESCRIPTION
        Crea entradas de log en el registro para rastrear eventos del proceso de despliegue.
        Los logs se guardan en HKLM:\SOFTWARE\ondoan\Deployments\<AppName>\Logs
    
    .PARAMETER AppName
        Nombre de la aplicacion
    
    .PARAMETER EventType
        Tipo de evento: MessageShown, UserResponse, InstallationStarted, InstallationCompleted
    
    .PARAMETER Details
        Detalles adicionales del evento
    
    .PARAMETER Attempt
        Numero de intento actual
    
    .EXAMPLE
        Add-DeploymentLog -AppName "office64" -EventType "MessageShown" -Details "Intento 1 de 5" -Attempt 1
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("MessageShown", "UserResponse", "InstallationStarted", "InstallationCompleted")]
        [string]$EventType,
        
        [Parameter(Mandatory = $false)]
        [string]$Details = "",
        
        [Parameter(Mandatory = $false)]
        [int]$Attempt = 0
    )
    
    try {
        # Crear estructura de registro si no existe
        $basePath = "HKLM:\SOFTWARE\ondoan"
        if (-not (Test-Path $basePath)) {
            New-Item -Path $basePath -Force | Out-Null
        }
        
        $deploymentsPath = "$basePath\Deployments"
        if (-not (Test-Path $deploymentsPath)) {
            New-Item -Path $deploymentsPath -Force | Out-Null
        }
        
        $appPath = "$deploymentsPath\$AppName"
        if (-not (Test-Path $appPath)) {
            New-Item -Path $appPath -Force | Out-Null
        }
        
        $logsPath = "$appPath\Logs"
        if (-not (Test-Path $logsPath)) {
            New-Item -Path $logsPath -Force | Out-Null
        }
        
        # Crear entrada de log con timestamp unico
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss-fff"
        $logEntryPath = "$logsPath\$timestamp"
        
        New-Item -Path $logEntryPath -Force | Out-Null
        
        # Guardar datos del evento
        Set-ItemProperty -Path $logEntryPath -Name "Timestamp" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Type String
        Set-ItemProperty -Path $logEntryPath -Name "EventType" -Value $EventType -Type String
        Set-ItemProperty -Path $logEntryPath -Name "Details" -Value $Details -Type String
        Set-ItemProperty -Path $logEntryPath -Name "Attempt" -Value $Attempt -Type DWord
        
        Write-Verbose "Log registrado: $EventType - $Details"
        
        return $true
    }
    catch {
        Write-Warning "Error al registrar log: $_"
        return $false
    }
}

function Get-DeploymentLog {
    <#
    .SYNOPSIS
        Recupera los logs de despliegue del registro de Windows.
    
    .DESCRIPTION
        Lee las entradas de log del registro para una aplicación específica.
        Los logs se leen de HKLM:\SOFTWARE\ondoan\Deployments\<AppName>\Logs
    
    .PARAMETER AppName
        Nombre de la aplicacion para la cual recuperar los logs
    
    .PARAMETER EventType
        Filtrar por tipo de evento específico (opcional)
    
    .PARAMETER Attempt
        Filtrar por número de intento específico (opcional)
    
    .PARAMETER Last
        Devolver solo los últimos N logs
    
    .OUTPUTS
        Array de PSCustomObject con las propiedades Timestamp, EventType, Details, Attempt
    
    .EXAMPLE
        Get-DeploymentLog -AppName "office64"
        Obtiene todos los logs de office64
    
    .EXAMPLE
        Get-DeploymentLog -AppName "test" -EventType "UserResponse"
        Obtiene solo los logs de respuesta de usuario
    
    .EXAMPLE
        Get-DeploymentLog -AppName "office64" -Last 10
        Obtiene los últimos 10 logs
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("MessageShown", "UserResponse", "InstallationStarted", "InstallationCompleted")]
        [string]$EventType,
        
        [Parameter(Mandatory = $false)]
        [int]$Attempt,
        
        [Parameter(Mandatory = $false)]
        [int]$Last
    )
    
    try {
        $logsPath = "HKLM:\SOFTWARE\ondoan\Deployments\$AppName\Logs"
        
        if (-not (Test-Path $logsPath)) {
            # No hay logs, devolver JSON con resultado OK y log vacío
            $result = @{
                result = "OK"
                log    = ""
            }
            return ($result | ConvertTo-Json -Compress)
        }
        
        # Obtener todas las entradas de log
        $logEntries = Get-ChildItem -Path $logsPath | ForEach-Object {
            $logEntry = Get-ItemProperty -Path $_.PSPath
            
            [PSCustomObject]@{
                Timestamp = $logEntry.Timestamp
                EventType = $logEntry.EventType
                Details   = $logEntry.Details
                Attempt   = $logEntry.Attempt
                EntryName = $_.PSChildName
            }
        }
        
        # Aplicar filtros si se especificaron
        if ($EventType) {
            $logEntries = $logEntries | Where-Object { $_.EventType -eq $EventType }
        }
        
        if ($Attempt) {
            $logEntries = $logEntries | Where-Object { $_.Attempt -eq $Attempt }
        }
        
        # Ordenar por timestamp (más reciente primero)
        $logEntries = $logEntries | Sort-Object -Property EntryName -Descending
        
        # Aplicar límite si se especificó
        if ($Last) {
            $logEntries = $logEntries | Select-Object -First $Last
        }
        
        # Obtener el último log (más reciente)
        $lastLog = $logEntries | Select-Object -First 1
        
        if ($lastLog) {
            # Formatear el log como array de líneas
            $logLines = @(
                $lastLog.Timestamp,
                $lastLog.EventType,
                $lastLog.Details,
                $lastLog.Attempt.ToString()
            )
            
            $result = @{
                result = "OK"
                log    = $logLines
            }
        }
        else {
            # No hay logs después de aplicar filtros
            $result = @{
                result = "OK"
                log    = @()
            }
        }
        
        return ($result | ConvertTo-Json -Compress)
    }
    catch {
        # En caso de error, devolver JSON con resultado ERROR
        $result = @{
            result = "ERROR"
            log    = @("Error al recuperar logs: $_")
        }
        return ($result | ConvertTo-Json -Compress)
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
        Si no se especifica, se obtiene de Get-DeployCnf del modulo.
    
    .PARAMETER Every
        Intervalo en minutos entre cada intento.
        Si no se especifica, se obtiene de Get-DeployCnf del modulo.
    
    .PARAMETER Message
        Mensaje personalizado a mostrar en el dialogo de confirmacion.
        Si no se especifica, se obtiene de Get-DeployCnf del modulo (si existe).
    
    .EXAMPLE
        Start-GbDeploy -Name "office64"
        # Usa configuracion por defecto de Get-DeployCnf del modulo office64
    
    .EXAMPLE
        Start-GbDeploy -Name "office64" -N 5 -Every 60
        # Pregunta 4 veces cada hora, en la 5ta vez instala automaticamente
    
    .EXAMPLE
        Start-GbDeploy -Name "MyApp" -N 3 -Every 30
        # Pregunta 2 veces cada 30 minutos, en la 3ra vez instala automaticamente
    
    .EXAMPLE
        Start-GbDeploy -Name "office64" -N 5 -Every 60 -Message "Se requiere actualizar Office a la version 64-bit para mejorar el rendimiento."
        # Usa un mensaje personalizado
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $false)]
        [int]$N = 0,
        
        [Parameter(Mandatory = $false)]
        [int]$Every = 0,
        
        [Parameter(Mandatory = $false)]
        [string]$Message = ""
    )
    
    try {
        # Si no se especificaron N o Every, intentar obtener de Get-DeployCnf
        if ($N -eq 0 -or $Every -eq 0) {
            Write-Verbose "Parametros N o Every no especificados, intentando obtener de Get-DeployCnf..."
            
            try {
                # Descargar el modulo para obtener la configuracion
                $moduleName = $Name.ToLower()
                $url = "https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/$moduleName.psm1"
                Write-Verbose "Descargando modulo desde: $url"
                $moduleContent = (new-object Net.WebClient).DownloadString($url)
                
                # Ejecutar el modulo
                Invoke-Expression $moduleContent
                
                # Intentar obtener configuracion
                if (Get-Command Get-DeployCnf -ErrorAction SilentlyContinue) {
                    $config = Get-DeployCnf
                    
                    if ($N -eq 0 -and $config.N) {
                        $N = $config.N
                        Write-Verbose "N obtenido de Get-DeployCnf: $N"
                    }
                    
                    if ($Every -eq 0 -and $config.Every) {
                        $Every = $config.Every
                        Write-Verbose "Every obtenido de Get-DeployCnf: $Every"
                    }
                    
                    if ([string]::IsNullOrWhiteSpace($Message) -and $config.Message) {
                        $Message = $config.Message
                        Write-Verbose "Message obtenido de Get-DeployCnf: $Message"
                    }
                }
                else {
                    Write-Warning "No se encontro la funcion Get-DeployCnf en el modulo $Name"
                }
            }
            catch {
                Write-Warning "Error al obtener configuracion de Get-DeployCnf: $_"
            }
        }
        
        # Validar que ahora tenemos valores para N y Every
        if ($N -eq 0 -or $Every -eq 0) {
            throw "Los parametros N y Every son obligatorios si no se pueden obtener de Get-DeployCnf. N=$N, Every=$Every"
        }
        
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
            
            # Log: Instalacion forzada iniciada
            Add-DeploymentLog -AppName $Name -EventType "InstallationStarted" -Details "Instalacion forzada - ultimo intento" -Attempt $N
            
            # Mostrar mensaje de instalación en curso
            Show-UserMessage -Message "INSTALANDO $Name...`n`nPor favor, NO REINICIE ni APAGUE el equipo durante la instalacion.`n`nEste proceso puede tardar varios minutos." -Title "Instalacion en Curso" -Timeout 0
            
            # Ejecutar la instalacion
            Write-Host "Ejecutando instalacion de $Name..." -ForegroundColor Green
            $deployResult = Invoke-GbDeployment -Name $Name
            
            # Log: Instalacion completada
            $status = if ($deployResult.Success) { "Exitosa" } else { "Fallida" }
            Add-DeploymentLog -AppName $Name -EventType "InstallationCompleted" -Details "Estado: $status - $($deployResult.Message)" -Attempt $N
            
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
            
            # Detectar si es la primera ejecución (no existe tarea previa)
            $isFirstRun = ($currentAttempt -eq 1 -and -not $existingTask)
            
            if ($isFirstRun) {
                # PRIMERA EJECUCIÓN: Solo programar la tarea para 1 minuto después
                Write-Host "=== PRIMERA EJECUCION ===" -ForegroundColor Cyan
                Write-Host "Programando primera verificacion en 1 minuto..." -ForegroundColor Yellow
                
                # Log: Primera ejecución
                Add-DeploymentLog -AppName $Name -EventType "MessageShown" -Details "Primera ejecucion - programando tarea" -Attempt 1
                
                # Incrementar contador de intentos para la siguiente ejecución
                $nextAttempt = 2
                
                # Crear metadatos
                $metadata = @{
                    CurrentAttempt  = $nextAttempt
                    TotalAttempts   = $N
                    IntervalMinutes = $Every
                    AppName         = $Name
                    LastAttempt     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                } | ConvertTo-Json -Compress
                
                # Crear el script que se ejecutara en la tarea
                $messageParam = if ([string]::IsNullOrWhiteSpace($Message)) { "" } else { " -Message '$($Message -replace "'", "''")'" }
                $scriptBlock = [scriptblock]::Create(@"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression
Start-GbDeploy -Name '$Name' -N $N -Every $Every$messageParam
"@)
                
                # Calcular hora de siguiente ejecucion (1 minuto)
                $nextRunTime = (Get-Date).AddMinutes(1)
                
                # Crear nueva tarea programada
                Write-Verbose "Creando tarea para primera ejecucion real en $nextRunTime"
                
                $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$($scriptBlock.ToString())`""
                
                # Crear dos triggers: uno por tiempo y otro al logon
                $triggerTime = New-ScheduledTaskTrigger -Once -At $nextRunTime
                $triggerLogon = New-ScheduledTaskTrigger -AtLogOn
                
                $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                
                # Registrar con ambos triggers
                Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Trigger @($triggerTime, $triggerLogon) -Principal $principal -Settings $settings -Description $metadata -Force | Out-Null
                
                Write-Host "Tarea programada creada. Primera verificacion en: $nextRunTime" -ForegroundColor Green
                
                # Devolver JSON indicando que se programó la primera ejecución
                $jsonResult = @{
                    result = "OK"
                    log    = @("Primera ejecucion programada para: $nextRunTime")
                }
                return ($jsonResult | ConvertTo-Json -Compress)
            }
            else {
                # EJECUCIONES SUBSIGUIENTES: Preguntar al usuario
                Write-Host "=== INTENTO $currentAttempt de $N ===" -ForegroundColor Cyan
                
                # Construir mensaje para el usuario
                if ([string]::IsNullOrWhiteSpace($Message)) {
                    # Mensaje por defecto
                    $userMessage = "Desea instalar $Name ahora?`n`nSi selecciona 'Cancelar', se le volvera a preguntar en $Every minutos.`n`nIntentos restantes: $($N - $currentAttempt)"
                }
                else {
                    # Mensaje personalizado + info de intentos
                    $userMessage = "$Message`n`nDesea instalar $Name ahora?`n`nSi selecciona 'Cancelar', se le volvera a preguntar en $Every minutos.`n`nIntentos restantes: $($N - $currentAttempt)"
                }
                
                # Log: Mensaje mostrado al usuario
                Add-DeploymentLog -AppName $Name -EventType "MessageShown" -Details "Intento $currentAttempt de $N" -Attempt $currentAttempt
                
                # Preguntar al usuario (timeout de 15 minutos)
                $response = Show-UserPrompt -Message $userMessage -Title "Instalacion de $Name" -Buttons "OKCancel" -Icon "Question" -TimeoutSeconds 900
                
                # Log: Respuesta del usuario
                Add-DeploymentLog -AppName $Name -EventType "UserResponse" -Details "Respuesta: $response" -Attempt $currentAttempt
                
                if ($response -eq "OK") {
                    # Usuario acepto: Ejecutar instalacion y eliminar tarea
                    Write-Host "Usuario acepto la instalacion." -ForegroundColor Green
                    
                    # Log: Instalacion iniciada
                    Add-DeploymentLog -AppName $Name -EventType "InstallationStarted" -Details "Usuario acepto en intento $currentAttempt" -Attempt $currentAttempt
                    
                    # Mostrar mensaje de instalación en curso
                    Show-UserMessage -Message "INSTALANDO $Name...`n`nPor favor, NO REINICIE ni APAGUE el equipo durante la instalacion.`n`nEste proceso puede tardar varios minutos." -Title "Instalacion en Curso" -Timeout 0
                    
                    # Ejecutar la instalacion
                    Write-Host "Ejecutando instalacion de $Name..." -ForegroundColor Green
                    $deployResult = Invoke-GbDeployment -Name $Name
                    
                    # Log: Instalacion completada
                    $status = if ($deployResult.Success) { "Exitosa" } else { "Fallida" }
                    Add-DeploymentLog -AppName $Name -EventType "InstallationCompleted" -Details "Estado: $status - $($deployResult.Message)" -Attempt $currentAttempt
                    
                    # Guardar resultado en el registro
                    Save-DeploymentResult -AppName $Name -Result $deployResult | Out-Null
                    
                    # Eliminar la tarea programada
                    Write-Verbose "Eliminando tarea programada..."
                    Remove-GbScheduledTask -TaskName $taskName -Force -ErrorAction SilentlyContinue
                    
                    if ($deployResult.Success) {
                        Write-Host "Despliegue de $Name completado exitosamente." -ForegroundColor Green
                        
                        # Devolver JSON con resultado exitoso
                        $jsonResult = @{
                            result = "OK"
                            log    = @("Despliegue completado: $($deployResult.Message)")
                        }
                        return ($jsonResult | ConvertTo-Json -Compress)
                    }
                    else {
                        Write-Warning "El despliegue de $Name finalizo con errores: $($deployResult.Message)"
                        
                        # Devolver JSON con error
                        $jsonResult = @{
                            result = "ERROR"
                            log    = @("Despliegue fallido: $($deployResult.Message)")
                        }
                        return ($jsonResult | ConvertTo-Json -Compress)
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
                    # Incluir el parametro Message si esta presente
                    $messageParam = if ([string]::IsNullOrWhiteSpace($Message)) { "" } else { " -Message '$($Message -replace "'", "''")'" }
                    $scriptBlock = [scriptblock]::Create(@"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression
Start-GbDeploy -Name '$Name' -N $N -Every $Every$messageParam
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
                
                    # Devolver JSON indicando que se programó siguiente intento
                    $jsonResult = @{
                        result = "OK"
                        log    = @("Usuario rechazo. Siguiente intento: $nextRunTime (Intento $nextAttempt de $N)")
                    }
                    return ($jsonResult | ConvertTo-Json -Compress)
                }
            }
        }
    }
    catch {
        Write-Error "Error en Start-GbDeploy: $_"
        Write-Error $_.ScriptStackTrace
        
        # Devolver JSON con error
        $jsonResult = @{
            result = "ERROR"
            log    = @("Error en Start-GbDeploy: $_")
        }
        return ($jsonResult | ConvertTo-Json -Compress)
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
