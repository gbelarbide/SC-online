<#
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression; Start-GbDeploy -Name "Test"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression; Start-GbDeploy -Name "office64"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/gbdeploy.psm1') | Invoke-Expression; Get-DeploymentLog -AppName "office64"
#>

function Write-GbLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "VERBOSE", "SUCCESS")]
        [string]$Level = "INFO"
    )
    try {
        $logDir = "C:\temp"
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $logFile = Join-Path $logDir "gbdeploy.log"
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $fullMessage = "[$timestamp] [$Level] $Message"
        
        # Escribir al archivo
        $fullMessage | Out-File -FilePath $logFile -Append -Encoding UTF8
        
        # Escribir a la consola según el nivel
        switch ($Level) {
            "ERROR" { Write-Host "[$Level] $Message" -ForegroundColor Red }
            "WARNING" { Write-Warning $Message }
            "VERBOSE" { Write-Verbose $Message }
            "SUCCESS" { Write-Host $Message -ForegroundColor Green }
            "INFO" { Write-Host $Message }
        }
    }
    catch {
        # Fallback silencioso si no se puede escribir el log
    }
}

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

function Show-InstallationProgress {
    <#
    .SYNOPSIS
        Muestra ventana de instalación que permanece en primer plano.
    
    .DESCRIPTION
        Usa la técnica de ventana padre TopMost invisible + ShowDialog para
        garantizar que la ventana permanezca siempre en primer plano.
        Funciona en contexto de usuario y SYSTEM (con tareas programadas).
    
    .PARAMETER AppName
        Nombre de la aplicación que se está instalando.
    
    .EXAMPLE
        $progressWindow = Show-InstallationProgress -AppName "Office 64-bit"
        # ... realizar instalación ...
        Close-InstallationProgress -ProgressInfo $progressWindow
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )
    
    try {
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $isSystem = $currentUser.IsSystem
        Write-Verbose "Ejecutando como: $($currentUser.Name), IsSystem: $isSystem"
        
        # Obtener branding
        $brandingMessage = "DeployCnf"
        try {
            if (Get-Command Get-DeployCnf -ErrorAction SilentlyContinue) {
                $cnfResult = Get-DeployCnf
                if ($cnfResult) {
                    try {
                        $cnfObject = $cnfResult | ConvertFrom-Json
                        if ($cnfObject.Message) { $brandingMessage = $cnfObject.Message }
                        else { $brandingMessage = $cnfResult }
                    }
                    catch { $brandingMessage = $cnfResult }
                }
            }
        }
        catch { Write-Verbose "Usando branding por defecto" }
        
        # Escapar para script
        $escapedAppName = $AppName -replace "'", "''"
        $escapedBranding = $brandingMessage -replace "'", "''" -replace "`r`n", "`n" -replace "`n", " - "
        
        # Carpeta temporal y extracción del Executer
        $tempFolder = if ($isSystem) { "C:\ProgramData\Temp" } else { $env:TEMP }
        if (-not (Test-Path $tempFolder)) {
            New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
        }

        $executerPath = Join-Path $tempFolder "Executer.exe"
        if (-not (Test-Path $executerPath)) {
            Write-Verbose "Extrayendo Executer a $executerPath"
            $bytes = [Convert]::FromBase64String($Executer)
            [System.IO.File]::WriteAllBytes($executerPath, $bytes)
        }
        
        # Script que muestra la ventana con técnica TopMost + ShowDialog
        
        # Script que muestra la ventana con técnica TopMost + ShowDialog
        $scriptPath = "$tempFolder\ShowInstallProgress_$(Get-Random).ps1"
        $scriptContent = @"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# TÉCNICA CLAVE: Ventana padre invisible con TopMost
`$topWindow = New-Object System.Windows.Forms.Form
`$topWindow.TopMost = `$true
`$topWindow.WindowState = 'Minimized'
`$topWindow.ShowInTaskbar = `$false

# Ventana principal
`$form = New-Object System.Windows.Forms.Form
`$form.Text = 'Instalando'
`$form.Size = New-Object System.Drawing.Size(550, 400)
`$form.StartPosition = 'CenterScreen'
`$form.FormBorderStyle = 'FixedDialog'
`$form.MaximizeBox = `$false
`$form.MinimizeBox = `$false
`$form.TopMost = `$true
`$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
`$form.ShowInTaskbar = `$true

# Panel
`$panel = New-Object System.Windows.Forms.Panel
`$panel.Size = New-Object System.Drawing.Size(490, 320)
`$panel.Location = New-Object System.Drawing.Point(30, 30)
`$panel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
`$panel.BorderStyle = 'None'

# Branding
`$labelBranding = New-Object System.Windows.Forms.Label
`$labelBranding.Text = '$escapedBranding'
`$labelBranding.Font = New-Object System.Drawing.Font('Segoe UI', 10)
`$labelBranding.ForeColor = [System.Drawing.Color]::FromArgb(102, 102, 102)
`$labelBranding.Size = New-Object System.Drawing.Size(450, 60)
`$labelBranding.Location = New-Object System.Drawing.Point(20, 20)
`$labelBranding.TextAlign = 'MiddleCenter'

# Título
`$labelTitle = New-Object System.Windows.Forms.Label
`$labelTitle.Text = 'INSTALANDO'
`$labelTitle.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
`$labelTitle.ForeColor = [System.Drawing.Color]::Black
`$labelTitle.Size = New-Object System.Drawing.Size(450, 40)
`$labelTitle.Location = New-Object System.Drawing.Point(20, 90)
`$labelTitle.TextAlign = 'MiddleCenter'

# App Name
`$labelApp = New-Object System.Windows.Forms.Label
`$labelApp.Text = '$escapedAppName'
`$labelApp.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
`$labelApp.ForeColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
`$labelApp.Size = New-Object System.Drawing.Size(450, 40)
`$labelApp.Location = New-Object System.Drawing.Point(20, 140)
`$labelApp.TextAlign = 'MiddleCenter'

# Dots
`$labelDots = New-Object System.Windows.Forms.Label
`$labelDots.Text = ''
`$labelDots.Font = New-Object System.Drawing.Font('Segoe UI', 24, [System.Drawing.FontStyle]::Bold)
`$labelDots.ForeColor = [System.Drawing.Color]::Black
`$labelDots.Size = New-Object System.Drawing.Size(450, 50)
`$labelDots.Location = New-Object System.Drawing.Point(20, 180)
`$labelDots.TextAlign = 'MiddleCenter'

`$panel.Controls.AddRange(@(`$labelBranding, `$labelTitle, `$labelApp, `$labelDots))
`$form.Controls.Add(`$panel)

# Timer animación
`$dotCount = 0
`$timer = New-Object System.Windows.Forms.Timer
`$timer.Interval = 500
`$timer.Add_Tick({
    `$script:dotCount = (`$script:dotCount + 1) % 4
    `$labelDots.Text = '.' * `$script:dotCount
})
`$timer.Start()

# Prevenir cierre Alt+F4
`$form.Add_FormClosing({
    param(`$s, `$e)
    if (`$e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
        `$e.Cancel = `$true
    }
})

# Guardar referencias globales
`$global:InstallProgressForm = `$form
`$global:InstallProgressTimer = `$timer
`$global:InstallProgressTopWindow = `$topWindow

# CLAVE: ShowDialog con ventana padre TopMost
[void]`$form.ShowDialog(`$topWindow)

`$topWindow.Dispose()
"@
        
        $scriptContent | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
        
        $process = $null
        $taskName = $null
        
        if (-not $isSystem) {
            # Usuario normal - Usar Executer.exe
            Write-Verbose "Ejecutando en sesión de usuario usando Executer"
            $process = Start-Process -FilePath $executerPath `
                -ArgumentList "-File `"$scriptPath`"" `
                -PassThru -WindowStyle Hidden
            Write-Verbose "PID: $($process.Id)"
        }
        else {
            # SYSTEM - usar tarea programada con Executer.exe
            Write-Verbose "Ejecutando como SYSTEM usando Executer"
            $sessions = query user 2>$null | Select-Object -Skip 1
            
            if (-not $sessions) {
                Write-Warning "No hay sesiones activas"
                return $null
            }
            
            foreach ($session in $sessions) {
                if ([string]::IsNullOrWhiteSpace($session)) { continue }
                
                $sessionInfo = $session -split '\s+' | Where-Object { $_ -ne '' }
                $sessionUser = $sessionInfo[0]
                $sessionId = $null
                
                foreach ($item in $sessionInfo) {
                    if ($item -match '^\d+$') {
                        $sessionId = $item
                        break
                    }
                }
                
                if ($sessionId) {
                    Write-Verbose "Tarea para: $sessionUser (ID: $sessionId)"
                    
                    $taskName = "InstallProgress_$(Get-Random)"
                    $action = New-ScheduledTaskAction -Execute $executerPath `
                        -Argument "-File `"$scriptPath`""
                    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
                    $principal = New-ScheduledTaskPrincipal -UserId $sessionUser -LogonType Interactive
                    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
                    
                    Register-ScheduledTask -TaskName $taskName -TaskPath "\Temp\" -Action $action `
                        -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
                    Start-ScheduledTask -TaskName $taskName -TaskPath "\Temp\"
                    
                    Write-Verbose "Tarea creada: \Temp\$taskName"
                    Start-Sleep -Milliseconds 2000
                    
                    # Buscar proceso
                    for ($i = 0; $i -lt 15; $i++) {
                        $process = Get-Process -Name "Executer" -ErrorAction SilentlyContinue | 
                        Where-Object { 
                            try { $_.CommandLine -like "*$scriptPath*" } catch { $false }
                        } | 
                        Select-Object -First 1
                        if ($process) { break }
                        Start-Sleep -Milliseconds 1000
                    }
                    
                    if ($process) { Write-Verbose "Proceso encontrado: $($process.Id)" }
                    break
                }
            }
        }
        
        return @{
            Process    = $process
            ScriptPath = $scriptPath
            TaskName   = $taskName
            IsSystem   = $isSystem
        }
    }
    catch {
        Write-Error "Error: $_"
        return $null
    }
}



function Close-InstallationProgress {
    <#
    .SYNOPSIS
        Cierra la ventana de progreso.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ProgressInfo
    )
    
    try {
        Write-Verbose "Cerrando ventana..."
        
        # Cerrar proceso
        if ($ProgressInfo.Process) {
            Write-Verbose "Deteniendo PID: $($ProgressInfo.Process.Id)"
            Stop-Process -Id $ProgressInfo.Process.Id -Force -ErrorAction SilentlyContinue
        }
        
        # Limpiar script
        if ($ProgressInfo.ScriptPath -and (Test-Path $ProgressInfo.ScriptPath)) {
            Write-Verbose "Eliminando script: $($ProgressInfo.ScriptPath)"
            Remove-Item $ProgressInfo.ScriptPath -Force -ErrorAction SilentlyContinue
        }
        
        # Limpiar tarea
        if ($ProgressInfo.TaskName) {
            Write-Verbose "Eliminando tarea: $($ProgressInfo.TaskName)"
            Unregister-ScheduledTask -TaskName $ProgressInfo.TaskName -TaskPath "\Temp\" `
                -Confirm:$false -ErrorAction SilentlyContinue
        }
        
        Write-Verbose "Ventana cerrada"
    }
    catch {
        Write-Error "Error al cerrar: $_"
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
        [ValidateSet("OKCancel", "YesNo", "OK")]
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
        
        $executerPath = Join-Path $tempFolder "Executer.exe"
        if (-not (Test-Path $executerPath)) {
            $bytes = [Convert]::FromBase64String($Executer)
            [System.IO.File]::WriteAllBytes($executerPath, $bytes)
        }
        
        
        $resultPath = "$tempFolder\UserPrompt_Result_$(Get-Random).txt"
        
        # Obtener branding para coherencia visual
        $brandingMessage = "DeployCnf"
        try {
            if (Get-Command Get-DeployCnf -ErrorAction SilentlyContinue) {
                $cnfResult = Get-DeployCnf
                if ($cnfResult) {
                    try {
                        $cnfObject = $cnfResult | ConvertFrom-Json
                        if ($cnfObject.Message) { $brandingMessage = $cnfObject.Message }
                        else { $brandingMessage = $cnfResult }
                    }
                    catch { $brandingMessage = $cnfResult }
                }
            }
        }
        catch { }
        
        # Escapar variables para el script
        $escapedBranding = $brandingMessage -replace "'", "''" -replace "`r`n", "`n" -replace "`n", " - "
        $escapedMessage = $Message -replace "'", "''"
        $escapedTitle = $Title -replace "'", "''"
        
        $resultPath = "$tempFolder\UserPrompt_Result_$(Get-Random).txt"
        $scriptPath = "$tempFolder\UserPrompt_$(Get-Random).ps1"
        
        # Script que muestra la ventana con técnica TopMost + ShowDialog
        $scriptContent = @"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# TÉCNICA CLAVE: Ventana padre invisible con TopMost
`$topWindow = New-Object System.Windows.Forms.Form
`$topWindow.TopMost = `$true
`$topWindow.WindowState = 'Minimized'
`$topWindow.ShowInTaskbar = `$false

# Ventana principal
`$form = New-Object System.Windows.Forms.Form
`$form.Text = '$escapedTitle'
`$form.Size = New-Object System.Drawing.Size(550, 480)
`$form.StartPosition = 'CenterScreen'
`$form.FormBorderStyle = 'FixedDialog'
`$form.MaximizeBox = `$false
`$form.MinimizeBox = `$false
`$form.TopMost = `$true
`$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
`$form.ShowInTaskbar = `$true

# Branding
`$labelBranding = New-Object System.Windows.Forms.Label
`$labelBranding.Text = '$escapedBranding'
`$labelBranding.Font = New-Object System.Drawing.Font('Segoe UI', 10)
`$labelBranding.ForeColor = [System.Drawing.Color]::FromArgb(102, 102, 102)
`$labelBranding.Size = New-Object System.Drawing.Size(490, 50)
`$labelBranding.Location = New-Object System.Drawing.Point(30, 20)
`$labelBranding.TextAlign = 'MiddleCenter'

# Título de Acción
`$labelTitle = New-Object System.Windows.Forms.Label
`$labelTitle.Text = 'ATENCIÓN REQUERIDA'
`$labelTitle.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
`$labelTitle.ForeColor = [System.Drawing.Color]::Black
`$labelTitle.Size = New-Object System.Drawing.Size(490, 40)
`$labelTitle.Location = New-Object System.Drawing.Point(30, 80)
`$labelTitle.TextAlign = 'MiddleCenter'

# Mensaje
`$labelMsg = New-Object System.Windows.Forms.Label
`$labelMsg.Text = '$escapedMessage'
`$labelMsg.Font = New-Object System.Drawing.Font('Segoe UI', 11)
`$labelMsg.ForeColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
`$labelMsg.Size = New-Object System.Drawing.Size(450, 160)
`$labelMsg.Location = New-Object System.Drawing.Point(50, 130)
`$labelMsg.TextAlign = 'TopCenter'

# Countdown Label
`$labelCountdown = New-Object System.Windows.Forms.Label
`$labelCountdown.Text = ''
`$labelCountdown.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
`$labelCountdown.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
`$labelCountdown.Size = New-Object System.Drawing.Size(490, 30)
`$labelCountdown.Location = New-Object System.Drawing.Point(30, 300)
`$labelCountdown.TextAlign = 'MiddleCenter'

# Bandera para permitir cierre
`$script:allowClose = `$false

# Función para guardar resultado
function Save-Result([string]`$res) {
    Set-Content -Path '$($resultPath -replace '\\', '\\')' -Value `$res -Force
    `$script:allowClose = `$true
    `$form.Close()
}

# Panel de botones
`$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
`$btnPanel.FlowDirection = 'RightToLeft'
`$btnPanel.Size = New-Object System.Drawing.Size(490, 60)
`$btnPanel.Location = New-Object System.Drawing.Point(30, 350)

if ('$Buttons' -eq 'YesNo') {
    `$btnNo = New-Object System.Windows.Forms.Button
    `$btnNo.Text = 'No'
    `$btnNo.Size = New-Object System.Drawing.Size(120, 40)
    `$btnNo.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    `$btnNo.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    `$btnNo.FlatStyle = 'Flat'
    `$btnNo.Add_Click({ Save-Result 'No' })
    
    `$btnYes = New-Object System.Windows.Forms.Button
    `$btnYes.Text = 'Sí'
    `$btnYes.Size = New-Object System.Drawing.Size(120, 40)
    `$btnYes.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    `$btnYes.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    `$btnYes.ForeColor = [System.Drawing.Color]::White
    `$btnYes.FlatStyle = 'Flat'
    `$btnYes.Add_Click({ Save-Result 'Yes' })
    
    `$btnPanel.Controls.AddRange(@(`$btnNo, `$btnYes))
}
elseif ('$Buttons' -eq 'OK') {
    `$btnOk = New-Object System.Windows.Forms.Button
    `$btnOk.Text = 'Entendido'
    `$btnOk.Size = New-Object System.Drawing.Size(120, 40)
    `$btnOk.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    `$btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    `$btnOk.ForeColor = [System.Drawing.Color]::White
    `$btnOk.FlatStyle = 'Flat'
    `$btnOk.Add_Click({ Save-Result 'OK' })
    `$btnPanel.Controls.Add(`$btnOk)
}
else {
    `$btnCancel = New-Object System.Windows.Forms.Button
    `$btnCancel.Text = 'Más tarde'
    `$btnCancel.Size = New-Object System.Drawing.Size(140, 40)
    `$btnCancel.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    `$btnCancel.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    `$btnCancel.FlatStyle = 'Flat'
    `$btnCancel.Add_Click({ Save-Result 'Cancel' })
    
    `$btnOk = New-Object System.Windows.Forms.Button
    `$btnOk.Text = 'Instalar ahora'
    `$btnOk.Size = New-Object System.Drawing.Size(140, 40)
    `$btnOk.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    `$btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    `$btnOk.ForeColor = [System.Drawing.Color]::White
    `$btnOk.FlatStyle = 'Flat'
    `$btnOk.Add_Click({ Save-Result 'OK' })
    
    `$btnPanel.Controls.AddRange(@(`$btnCancel, `$btnOk))
}

`$form.Controls.AddRange(@(`$labelBranding, `$labelTitle, `$labelMsg, `$labelCountdown, `$btnPanel))

# Timer Animación / Timeout
if ($TimeoutSeconds -gt 0) {
    `$script:timeLeft = $TimeoutSeconds
    `$timer = New-Object System.Windows.Forms.Timer
    `$timer.Interval = 1000
    `$timer.Add_Tick({
        `$script:timeLeft--
        if (`$script:timeLeft -le 0) {
            `$timer.Stop()
            Save-Result 'OK'
        }
        else {
            `$m = [Math]::Floor(`$script:timeLeft / 60)
            `$s = `$script:timeLeft % 60
            `$labelCountdown.Text = "Se instalará automáticamente en: " + `$m + ":" + `$s.ToString("00")
        }
    })
    `$timer.Start()
}

# Prevenir cierre Alt+F4 a menos que hayamos pulsado un boton
`$form.Add_FormClosing({
    param(`$s, `$e)
    if (-not `$script:allowClose -and `$e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) { 
        `$e.Cancel = `$true 
    }
})

[void]`$form.ShowDialog(`$topWindow)
`$topWindow.Dispose()
"@
        $scriptContent | Out-File -FilePath $scriptPath -Encoding Unicode -Force
        
        $scriptExecutable = $executerPath
        
        if (-not $isSystem) {
            # Si NO estamos ejecutando como SYSTEM, ejecutar directamente usando Executer
            Write-Verbose "Ejecutando prompt directamente usando Executer"
            $null = Start-Process -FilePath $scriptExecutable -ArgumentList "-File `"$scriptPath`"" -Wait -PassThru -WindowStyle Hidden
        }
        else {
            # Si estamos ejecutando como SYSTEM, necesitamos ejecutar en la sesion del usuario
            Write-Verbose "Ejecutando como SYSTEM, buscando sesion interactiva para prompt"
            
            # Obtener la sesion interactiva
            $sessionId = (Get-Process -Name "explorer" -ErrorAction SilentlyContinue | Select-Object -First 1).SessionId
            
            if ($null -eq $sessionId) {
                Write-GbLog -Message "No se encontro sesion interactiva para mostrar prompt" -Level "WARNING"
                return "Cancel"
            }
            
            # Crear la tarea programada para la sesión del usuario
            $taskName = "UserPrompt_$(Get-Random)"
            $action = New-ScheduledTaskAction -Execute $scriptExecutable -Argument "-File `"$scriptPath`""
            $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(1)
            
            # Obtener el usuario de la sesion para la tarea
            $sessionUser = (query user | Select-String -Pattern "^\s*\S+\s+console\s+$sessionId" | ForEach-Object {
                    ($_ -split '\s+')[1]
                }) | Select-Object -First 1
            
            if (-not $sessionUser) { $sessionUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName }
            
            if ($sessionUser) {
                $principal = New-ScheduledTaskPrincipal -UserId $sessionUser -LogonType Interactive -RunLevel Highest
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
                Start-ScheduledTask -TaskName $taskName | Out-Null
            }
            else {
                Write-GbLog -Message "No se pudo determinar usuario de sesion $sessionId" -Level "WARNING"
                return "Cancel"
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
        Write-GbLog -Message "[$AppName] [$EventType] Attempt:$Attempt - $Details" -Level "INFO"
        
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
                Write-GbLog -Message "Error al obtener configuracion de Get-DeployCnf: $_" -Level "WARNING"
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
            
            # Mostrar mensaje al usuario (siempre en primer plano)
            Show-UserPrompt -Message "La aplicacion $Name se instalara en 5 minutos.`n`nPor favor, guarde su trabajo y cierre todas las aplicaciones." -Title "Instalacion Programada" -Buttons "OK" -Icon "Warning" -TimeoutSeconds 0
            
            # Esperar 5 minutos
            Write-Verbose "Esperando 5 minutos antes de la instalacion..."
            Start-Sleep -Seconds 300
            
            # PREINSTALACION: Descargar archivos antes de instalar
            Write-Host "Preparando archivos de instalacion..." -ForegroundColor Cyan
            try {
                # Descargar el modulo para ejecutar Start-Preinstall
                $moduleName = $Name.ToLower()
                $url = "https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/$moduleName.psm1"
                Write-Verbose "Descargando modulo desde: $url"
                $moduleContent = (new-object Net.WebClient).DownloadString($url)
                
                # Ejecutar el modulo
                Invoke-Expression $moduleContent
                
                # Ejecutar Start-Preinstall si existe
                if (Get-Command Start-Preinstall -ErrorAction SilentlyContinue) {
                    Write-Verbose "Ejecutando Start-Preinstall..."
                    $preinstallResult = Start-Preinstall
                    
                    if ($preinstallResult.Success) {
                        Write-Host "Archivos de instalacion preparados correctamente" -ForegroundColor Green
                    }
                    else {
                        Write-Warning "Error en la preparacion: $($preinstallResult.ErrorMessage)"
                    }
                }
                else {
                    Write-Verbose "El modulo $Name no tiene funcion Start-Preinstall"
                }
            }
            catch {
                Write-GbLog -Message "Error al preparar archivos: $_" -Level "WARNING"
                # Continuar de todos modos, el error se manejara en la instalacion
            }
            
            # Log: Instalacion forzada iniciada
            Add-DeploymentLog -AppName $Name -EventType "InstallationStarted" -Details "Instalacion forzada - ultimo intento" -Attempt $N
            
            # Ejecutar la instalacion
            Write-GbLog -Message "Ejecutando instalacion de $Name..." -Level "SUCCESS"
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
                Write-GbLog -Message "Despliegue de $Name completado exitosamente." -Level "SUCCESS"
            }
            else {
                Write-GbLog -Message "El despliegue de $Name finalizo con errores: $($deployResult.Message)" -Level "WARNING"
            }
        }
        else {
            # EJECUCIONES INTERMEDIAS: Preguntar al usuario
            
            # Detectar si es la primera ejecución (no existe tarea previa)
            $isFirstRun = ($currentAttempt -eq 1 -and -not $existingTask)
            
            if ($isFirstRun) {
                # PRIMERA EJECUCIÓN: Solo programar la tarea para 10 segundos después
                Write-GbLog -Message "=== PRIMERA EJECUCION ===" -Level "INFO"
                Write-GbLog -Message "Programando primera verificacion en 10 segundos..." -Level "WARNING"
                
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
                
                # Calcular hora de siguiente ejecucion (10 segundos)
                $nextRunTime = (Get-Date).AddSeconds(10)
                
                # Codificar script para powershell
                $scriptString = $scriptBlock.ToString()
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptString)
                $encodedCommand = [Convert]::ToBase64String($bytes)

                # Crear nueva tarea programada usando powershell (SYSTEM no necesita Executer para esto)
                Write-Verbose "Creando tarea para primera ejecucion real en $nextRunTime"
                $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"
                
                # Crear dos triggers: uno por tiempo y otro al logon
                $triggerTime = New-ScheduledTaskTrigger -Once -At $nextRunTime
                $triggerLogon = New-ScheduledTaskTrigger -AtLogOn
                
                $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                
                # Registrar con ambos triggers
                Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Trigger @($triggerTime, $triggerLogon) -Principal $principal -Settings $settings -Description $metadata -Force | Out-Null
                
                Write-GbLog -Message "Tarea programada creada. Primera verificacion en: $nextRunTime" -Level "SUCCESS"
                
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
                
                # PREINSTALACION: Descargar archivos antes de preguntar al usuario
                Write-Host "Preparando archivos de instalacion..." -ForegroundColor Cyan
                try {
                    # Descargar el modulo para ejecutar Start-Preinstall
                    $moduleName = $Name.ToLower()
                    $url = "https://raw.githubusercontent.com/gbelarbide/SC-online/refs/heads/main/Deploy/$moduleName.psm1"
                    Write-Verbose "Descargando modulo desde: $url"
                    $moduleContent = (new-object Net.WebClient).DownloadString($url)
                    
                    # Ejecutar el modulo
                    Invoke-Expression $moduleContent
                    
                    # Ejecutar Start-Preinstall si existe
                    if (Get-Command Start-Preinstall -ErrorAction SilentlyContinue) {
                        Write-Verbose "Ejecutando Start-Preinstall..."
                        $preinstallResult = Start-Preinstall
                        
                        if ($preinstallResult.Success) {
                            Write-Host "Archivos de instalacion preparados correctamente" -ForegroundColor Green
                        }
                        else {
                            Write-Warning "Error en la preparacion: $($preinstallResult.ErrorMessage)"
                        }
                    }
                    else {
                        Write-Verbose "El modulo $Name no tiene funcion Start-Preinstall"
                    }
                }
                catch {
                    Write-Warning "Error al preparar archivos: $_"
                    # Continuar de todos modos, el error se manejara en la instalacion
                }
                
                # DETECCION DE SESIONES: Verificar si hay usuarios conectados
                Write-Verbose "Verificando sesiones de usuario activas..."
                $activeSessions = $null
                try {
                    $activeSessions = query user 2>$null | Select-Object -Skip 1
                }
                catch {
                    Write-Verbose "No se pudieron obtener sesiones de usuario"
                }
                
                $hasActiveSessions = $false
                if ($activeSessions) {
                    foreach ($session in $activeSessions) {
                        if (-not [string]::IsNullOrWhiteSpace($session)) {
                            $hasActiveSessions = $true
                            break
                        }
                    }
                }
                
                if (-not $hasActiveSessions) {
                    # NO HAY SESIONES ACTIVAS: Instalar automáticamente
                    Write-GbLog -Message "=== NO SE DETECTARON SESIONES DE USUARIO ===" -Level "WARNING"
                    Write-GbLog -Message "Procediendo con instalacion automatica..." -Level "INFO"
                    
                    # Log: Instalación automática por falta de sesiones
                    Add-DeploymentLog -AppName $Name -EventType "InstallationStarted" -Details "Instalacion automatica - no hay sesiones activas (intento $currentAttempt)" -Attempt $currentAttempt
                    
                    # Ejecutar la instalacion
                    Write-GbLog -Message "Ejecutando instalacion de $Name..." -Level "SUCCESS"
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
                        Write-GbLog -Message "Despliegue de $Name completado exitosamente." -Level "SUCCESS"
                        
                        # Devolver JSON con resultado exitoso
                        $jsonResult = @{
                            result = "OK"
                            log    = @("Instalacion automatica completada (sin sesiones): $($deployResult.Message)")
                        }
                        return ($jsonResult | ConvertTo-Json -Compress)
                    }
                    else {
                        Write-GbLog -Message "El despliegue de $Name finalizo con errores: $($deployResult.Message)" -Level "WARNING"
                        
                        # Devolver JSON con error
                        $jsonResult = @{
                            result = "ERROR"
                            log    = @("Instalacion automatica fallida: $($deployResult.Message)")
                        }
                        return ($jsonResult | ConvertTo-Json -Compress)
                    }
                }
                
                # HAY SESIONES ACTIVAS: Preguntar al usuario
                Write-GbLog -Message "Sesiones de usuario detectadas. Mostrando dialogo..." -Level "INFO"
                
                # Log: Mensaje mostrado al usuario
                Add-DeploymentLog -AppName $Name -EventType "MessageShown" -Details "Intento $currentAttempt de $N" -Attempt $currentAttempt
                
                # Preguntar al usuario (timeout de 15 minutos)
                $response = Show-UserPrompt -Message $userMessage -Title "Instalacion de $Name" -Buttons "OKCancel" -Icon "Question" -TimeoutSeconds 900
                
                # Log: Respuesta del usuario
                Add-DeploymentLog -AppName $Name -EventType "UserResponse" -Details "Respuesta: $response" -Attempt $currentAttempt
                
                if ($response -eq "OK") {
                    # Usuario acepto: Ejecutar instalacion y eliminar tarea
                    Write-GbLog -Message "Usuario acepto la instalacion." -Level "SUCCESS"
                    
                    # Log: Instalacion iniciada
                    Add-DeploymentLog -AppName $Name -EventType "InstallationStarted" -Details "Usuario acepto en intento $currentAttempt" -Attempt $currentAttempt
                    
                    # Ejecutar la instalacion
                    Write-GbLog -Message "Ejecutando instalacion de $Name..." -Level "SUCCESS"
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
                        Write-GbLog -Message "Despliegue de $Name completado exitosamente." -Level "SUCCESS"
                        
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
                    Write-GbLog -Message "Usuario rechazo la instalacion. Programando siguiente intento..." -Level "WARNING"
                
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
                
                    Write-GbLog -Message "Siguiente intento programado para: $nextRunTime (o al iniciar sesion)" -Level "INFO"
                
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
        Write-GbLog -Message "Descargando modulo de despliegue para $Name..." -Level "INFO"
        
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
            Write-GbLog -Message "Ejecutando Start-Deploy..." -Level "SUCCESS"
            $deployResult = Start-Deploy
        }
        elseif (Get-Command Start-Install -ErrorAction SilentlyContinue) {
            Write-GbLog -Message "Ejecutando Start-Install..." -Level "SUCCESS"
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
        
        Write-GbLog -Message "Resultado guardado en el registro: $appPath" -Level "SUCCESS"
        Write-Verbose "JSON guardado: $jsonResult"
        
        return $true
    }
    catch {
        Write-Error "Error al guardar el resultado en el registro: $_"
        return $false
    }
}

#Exportar las funciones
#Export-ModuleMember -Function Show-UserMessage, Show-UserPrompt, New-GbScheduledTask, Remove-GbScheduledTask, Start-GbDeploy, Show-InstallationProgress
