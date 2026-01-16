
#(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression
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
        # Obtener la sesion interactiva del usuario
        $sessionId = (Get-Process -Name "explorer" -ErrorAction SilentlyContinue | Select-Object -First 1).SessionId
        
        if ($null -eq $sessionId) {
            Write-Warning "No se encontro una sesion interactiva de usuario."
            return "Cancel"
        }
        
        # Crear archivo temporal para el resultado
        $resultPath = "$env:TEMP\UserPrompt_Result_$(Get-Random).txt"
        
        # Escapar caracteres especiales para VBScript
        $escapedMessage = $Message -replace '"', '""'
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
        
        # Guardar el script VBScript
        $vbsPath = "$env:TEMP\UserPrompt_$(Get-Random).vbs"
        Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII -Force
        
        # Crear un script PowerShell que ejecute el VBScript en la sesion del usuario
        $psContent = @"
`$vbsPath = "$($vbsPath -replace '\\', '\\')"
Start-Process -FilePath "wscript.exe" -ArgumentList "`"`$vbsPath`"" -WindowStyle Hidden -Wait
"@
        
        $psPath = "$env:TEMP\UserPrompt_$(Get-Random).ps1"
        Set-Content -Path $psPath -Value $psContent -Encoding ASCII -Force
        
        # Usar PsExec si esta disponible, sino usar un metodo alternativo
        $psExecPath = "C:\Windows\System32\PsExec.exe"
        
        if (Test-Path $psExecPath) {
            # Ejecutar con PsExec en la sesion interactiva
            $process = Start-Process -FilePath $psExecPath -ArgumentList "-accepteula -s -i $sessionId powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$psPath`"" -WindowStyle Hidden -PassThru
        }
        else {
            # Metodo alternativo: crear una tarea programada que se ejecute inmediatamente
            $taskName = "UserPrompt_$(Get-Random)"
            $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsPath`""
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
            
            Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Force | Out-Null
            Start-ScheduledTask -TaskName $taskName
            
            # Esperar un momento para que la tarea se inicie
            Start-Sleep -Seconds 2
        }
        
        # Esperar a que el usuario responda o se alcance el timeout
        $maxWait = if ($TimeoutSeconds -gt 0) { $TimeoutSeconds + 10 } else { 300 }
        $elapsed = 0
        $checkInterval = 1
        
        while ($elapsed -lt $maxWait) {
            Start-Sleep -Seconds $checkInterval
            $elapsed += $checkInterval
            
            if (Test-Path $resultPath) {
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
            if (Test-Path $psPath) { Remove-Item $psPath -Force -ErrorAction SilentlyContinue }
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

# Exportar las funciones
#Export-ModuleMember -Function Show-UserMessage, Show-UserPrompt
