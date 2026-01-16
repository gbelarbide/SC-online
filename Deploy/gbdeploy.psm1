
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
        Utiliza VBScript para crear un MessageBox con botones personalizables y devuelve la respuesta del usuario.
    
    .PARAMETER Message
        El mensaje que se mostrara al usuario.
    
    .PARAMETER Title
        El titulo de la ventana del dialogo.
    
    .PARAMETER Buttons
        Tipo de botones a mostrar. Valores validos:
        - OKCancel (OK y Cancelar)
        - YesNo (Si y No)
        - YesNoCancel (Si, No y Cancelar)
        - RetryCancel (Reintentar y Cancelar)
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
        Por defecto: 300 (5 minutos)
    
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
        String - Devuelve la respuesta del usuario: "OK", "Cancel", "Yes", "No", "Retry", o "Timeout"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Title = "Confirmacion del Sistema",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("OKCancel", "YesNo", "YesNoCancel", "RetryCancel")]
        [string]$Buttons = "OKCancel",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Information", "Question", "Warning", "Error")]
        [string]$Icon = "Question",
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 300
    )
    
    try {
        # Obtener el usuario activo
        $activeUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
        
        if (-not $activeUser) {
            Write-Warning "No se encontro un usuario activo en el sistema."
            return "Cancel"
        }
        
        # Extraer solo el nombre de usuario (sin dominio)
        $userName = $activeUser.Split('\')[-1]
        
        # Mapear tipos de botones a valores VBScript
        $buttonValues = @{
            "OKCancel"    = 1
            "YesNo"       = 4
            "YesNoCancel" = 3
            "RetryCancel" = 5
        }
        
        # Mapear iconos a valores VBScript
        $iconValues = @{
            "Error"       = 16
            "Question"    = 32
            "Warning"     = 48
            "Information" = 64
        }
        
        $buttonValue = $buttonValues[$Buttons]
        $iconValue = $iconValues[$Icon]
        $style = $buttonValue + $iconValue
        
        # Crear script VBScript temporal
        $vbsPath = "$env:TEMP\UserPrompt_$(Get-Random).vbs"
        $resultPath = "$env:TEMP\UserPrompt_Result_$(Get-Random).txt"
        
        $vbsScript = @"
Dim objShell, result
Set objShell = CreateObject("WScript.Shell")

result = MsgBox("$Message", $style, "$Title")

' Escribir resultado en archivo
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("$resultPath", True)
file.WriteLine result
file.Close

WScript.Quit
"@
        
        # Guardar el script VBScript
        Set-Content -Path $vbsPath -Value $vbsScript -Encoding ASCII
        
        # Crear tarea programada temporal para ejecutar en el contexto del usuario
        $taskName = "UserPrompt_$(Get-Random)"
        
        # Ejecutar el script VBScript en la sesion del usuario usando schtasks
        $action = "wscript.exe `"$vbsPath`""
        
        # Crear y ejecutar la tarea
        schtasks /create /tn $taskName /tr $action /sc once /st 00:00 /ru $activeUser /rl highest /f | Out-Null
        schtasks /run /tn $taskName | Out-Null
        
        # Esperar a que el usuario responda o se alcance el timeout
        $elapsed = 0
        $checkInterval = 1
        
        while ($elapsed -lt $TimeoutSeconds) {
            Start-Sleep -Seconds $checkInterval
            $elapsed += $checkInterval
            
            if (Test-Path $resultPath) {
                break
            }
            
            # Verificar si la tarea sigue en ejecucion
            $taskStatus = schtasks /query /tn $taskName /fo csv | ConvertFrom-Csv
            if ($taskStatus.Status -notmatch "Running") {
                Start-Sleep -Seconds 2  # Esperar un poco mas para asegurar que el archivo se escriba
                break
            }
        }
        
        # Leer el resultado
        $userResponse = "Timeout"
        
        if (Test-Path $resultPath) {
            $resultValue = Get-Content $resultPath -Raw
            $resultValue = $resultValue.Trim()
            
            # Mapear valores de retorno VBScript a texto legible
            switch ($resultValue) {
                "1" { $userResponse = "OK" }
                "2" { $userResponse = "Cancel" }
                "3" { $userResponse = "Abort" }
                "4" { $userResponse = "Retry" }
                "5" { $userResponse = "Ignore" }
                "6" { $userResponse = "Yes" }
                "7" { $userResponse = "No" }
                "-1" { $userResponse = "Timeout" }
                default { $userResponse = "Cancel" }
            }
        }
        
        # Limpiar archivos temporales y tarea
        try {
            if (Test-Path $vbsPath) { Remove-Item $vbsPath -Force -ErrorAction SilentlyContinue }
            if (Test-Path $resultPath) { Remove-Item $resultPath -Force -ErrorAction SilentlyContinue }
            schtasks /delete /tn $taskName /f 2>$null | Out-Null
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
