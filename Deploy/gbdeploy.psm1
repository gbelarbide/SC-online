
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
        Utiliza mshta.exe (Microsoft HTML Application) para crear un dialogo visible en la sesion del usuario.
    
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
        # Crear archivo temporal para el resultado
        $resultPath = "$env:TEMP\UserPrompt_Result_$(Get-Random).txt"
        
        # Escapar caracteres especiales para HTML
        $escapedMessage = $Message -replace '"', '&quot;' -replace '<', '&lt;' -replace '>', '&gt;' -replace "'", "&#39;"
        $escapedTitle = $Title -replace '"', '&quot;' -replace '<', '&lt;' -replace '>', '&gt;' -replace "'", "&#39;"
        
        # Determinar los botones y el icono
        $button1Text = if ($Buttons -eq "YesNo") { "Si" } else { "OK" }
        $button2Text = if ($Buttons -eq "YesNo") { "No" } else { "Cancelar" }
        $button1Value = if ($Buttons -eq "YesNo") { "Yes" } else { "OK" }
        $button2Value = if ($Buttons -eq "YesNo") { "No" } else { "Cancel" }
        
        # Seleccionar icono
        $iconSymbol = switch ($Icon) {
            "Error" { "&#10060;" }  # ❌
            "Warning" { "&#9888;" }  # ⚠
            "Information" { "&#8505;" }  # ℹ
            default { "&#10067;" }  # ❓
        }
        
        # Crear HTML para mshta
        $htaContent = @"
<html>
<head>
    <title>$escapedTitle</title>
    <HTA:APPLICATION 
        ID="oHTA"
        APPLICATIONNAME="UserPrompt"
        BORDER="dialog"
        BORDERSTYLE="normal"
        CAPTION="yes"
        ICON=""
        MAXIMIZEBUTTON="no"
        MINIMIZEBUTTON="no"
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
            background: #f0f0f0;
        }
        .container {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 400px;
        }
        .icon {
            font-size: 48px;
            text-align: center;
            margin-bottom: 15px;
        }
        .message {
            font-size: 14px;
            margin-bottom: 20px;
            text-align: center;
            color: #333;
        }
        .buttons {
            text-align: center;
        }
        button {
            padding: 8px 20px;
            margin: 0 5px;
            font-size: 14px;
            border: 1px solid #0078d4;
            background: #0078d4;
            color: white;
            border-radius: 4px;
            cursor: pointer;
            min-width: 80px;
        }
        button:hover {
            background: #106ebe;
        }
        button.cancel {
            background: #6c757d;
            border-color: #6c757d;
        }
        button.cancel:hover {
            background: #5a6268;
        }
    </style>
    <script type="text/javascript">
        var timeoutSeconds = $TimeoutSeconds;
        var timeoutTimer = null;
        
        function resizeWindow() {
            window.resizeTo(450, 250);
            window.moveTo((screen.width - 450) / 2, (screen.height - 250) / 2);
        }
        
        function writeResult(result) {
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            var file = fso.CreateTextFile("$($resultPath -replace '\\', '\\')", true);
            file.WriteLine(result);
            file.Close();
            window.close();
        }
        
        function onButton1() {
            writeResult("$button1Value");
        }
        
        function onButton2() {
            writeResult("$button2Value");
        }
        
        function onTimeout() {
            writeResult("Timeout");
        }
        
        window.onload = function() {
            resizeWindow();
            if (timeoutSeconds > 0) {
                timeoutTimer = setTimeout(onTimeout, timeoutSeconds * 1000);
            }
        };
    </script>
</head>
<body>
    <div class="container">
        <div class="icon">$iconSymbol</div>
        <div class="message">$escapedMessage</div>
        <div class="buttons">
            <button onclick="onButton1()">$button1Text</button>
            <button class="cancel" onclick="onButton2()">$button2Text</button>
        </div>
    </div>
</body>
</html>
"@
        
        # Guardar el archivo HTA
        $htaPath = "$env:TEMP\UserPrompt_$(Get-Random).hta"
        Set-Content -Path $htaPath -Value $htaContent -Encoding UTF8 -Force
        
        # Ejecutar mshta
        $process = Start-Process -FilePath "mshta.exe" -ArgumentList "`"$htaPath`"" -PassThru
        
        # Esperar a que el usuario responda o se alcance el timeout
        $maxWait = if ($TimeoutSeconds -gt 0) { $TimeoutSeconds + 5 } else { 300 }
        $elapsed = 0
        $checkInterval = 1
        
        while ($elapsed -lt $maxWait) {
            Start-Sleep -Seconds $checkInterval
            $elapsed += $checkInterval
            
            if (Test-Path $resultPath) {
                break
            }
            
            # Verificar si el proceso sigue en ejecucion
            if ($process.HasExited) {
                Start-Sleep -Seconds 1
                break
            }
        }
        
        # Si se alcanzo el timeout maximo, matar el proceso
        if ($elapsed -ge $maxWait -and -not $process.HasExited) {
            $process.Kill()
            Write-Warning "Se alcanzo el timeout maximo esperando la respuesta del usuario."
        }
        
        # Leer el resultado
        $userResponse = "Cancel"
        
        if (Test-Path $resultPath) {
            $resultValue = (Get-Content $resultPath -Raw).Trim()
            
            if (-not [string]::IsNullOrWhiteSpace($resultValue)) {
                $userResponse = $resultValue
            }
        }
        
        # Limpiar archivos temporales
        try {
            if (Test-Path $htaPath) { Remove-Item $htaPath -Force -ErrorAction SilentlyContinue }
            if (Test-Path $resultPath) { Remove-Item $resultPath -Force -ErrorAction SilentlyContinue }
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
