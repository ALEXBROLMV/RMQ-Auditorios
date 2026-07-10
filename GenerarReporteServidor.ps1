<#
.SYNOPSIS
    Genera un reporte de actividad leyendo la base de datos SQLite.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Comprobar base de datos
$RutaBase = "C:\PROYECTOS\RMQ\Recolector"
$RutaDB = Join-Path -Path $RutaBase -ChildPath "Auditoria.db"
$RutaDLL = Join-Path -Path $RutaBase -ChildPath "System.Data.SQLite.dll"

if (-not (Test-Path $RutaDLL)) {
    [System.Windows.Forms.MessageBox]::Show("Falta el motor de Base de Datos SQLite en $RutaBase.", "Error Crítico", 0, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit
}

if (-not (Test-Path $RutaDB)) {
    [System.Windows.Forms.MessageBox]::Show("La base de datos Auditoria.db no existe aún. Ejecuta el script Recolector.ps1 primero.", "Advertencia", 0, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    exit
}

Add-Type -Path $RutaDLL

# Variables globales
$Global:DatosAccesos = @()
$Global:DatosApps = @()
$Global:DatosVivos = @()

# ----------------- UI -----------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Reporte de Actividad de Servidor (Modo: Base de Datos) - $env:COMPUTERNAME"
$Form.Size = New-Object System.Drawing.Size(1100, 750)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::White

$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Dock = "Fill"
$TabControl.Font = New-Object System.Drawing.Font("Arial", 10)

$Tab1 = New-Object System.Windows.Forms.TabPage
$Tab1.Text = "1. Historial de Conexiones"
$Grid1 = New-Object System.Windows.Forms.DataGridView
$Grid1.Dock = "Fill"
$Grid1.ReadOnly = $true
$Grid1.AllowUserToAddRows = $false
$Grid1.AutoSizeColumnsMode = "Fill"
$Grid1.BackgroundColor = [System.Drawing.Color]::White
$Tab1.Controls.Add($Grid1)

$Tab2 = New-Object System.Windows.Forms.TabPage
$Tab2.Text = "2. Historial de Programas"
$Grid2 = New-Object System.Windows.Forms.DataGridView
$Grid2.Dock = "Fill"
$Grid2.ReadOnly = $true
$Grid2.AllowUserToAddRows = $false
$Grid2.AutoSizeColumnsMode = "Fill"
$Grid2.BackgroundColor = [System.Drawing.Color]::White
$Tab2.Controls.Add($Grid2)

$Tab3 = New-Object System.Windows.Forms.TabPage
$Tab3.Text = "3. Programas en Vivo"
$Grid3 = New-Object System.Windows.Forms.DataGridView
$Grid3.Dock = "Fill"
$Grid3.ReadOnly = $true
$Grid3.AllowUserToAddRows = $false
$Grid3.AutoSizeColumnsMode = "Fill"
$Grid3.BackgroundColor = [System.Drawing.Color]::White
$Tab3.Controls.Add($Grid3)

$TabControl.Controls.Add($Tab1)
$TabControl.Controls.Add($Tab2)
$TabControl.Controls.Add($Tab3)
$Form.Controls.Add($TabControl)

$PanelBotones = New-Object System.Windows.Forms.Panel
$PanelBotones.Dock = "Bottom"
$PanelBotones.Height = 60
$PanelBotones.BackColor = [System.Drawing.Color]::FromArgb(26, 67, 80)

$BtnExcel = New-Object System.Windows.Forms.Button
$BtnExcel.Text = "Exportar a Excel (.csv)"
$BtnExcel.Location = New-Object System.Drawing.Point(20, 10)
$BtnExcel.Size = New-Object System.Drawing.Size(180, 40)
$BtnExcel.BackColor = [System.Drawing.Color]::FromArgb(140, 198, 63)
$BtnExcel.ForeColor = [System.Drawing.Color]::White
$BtnExcel.FlatStyle = "Flat"
$BtnExcel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

$BtnPDF = New-Object System.Windows.Forms.Button
$BtnPDF.Text = "Generar Reporte PDF"
$BtnPDF.Location = New-Object System.Drawing.Point(220, 10)
$BtnPDF.Size = New-Object System.Drawing.Size(180, 40)
$BtnPDF.BackColor = [System.Drawing.Color]::FromArgb(140, 198, 63)
$BtnPDF.ForeColor = [System.Drawing.Color]::White
$BtnPDF.FlatStyle = "Flat"
$BtnPDF.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

$LblCargando = New-Object System.Windows.Forms.Label
$LblCargando.Text = "Cargando historial desde Base de Datos..."
$LblCargando.Location = New-Object System.Drawing.Point(420, 20)
$LblCargando.AutoSize = $true
$LblCargando.ForeColor = [System.Drawing.Color]::White
$LblCargando.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Italic)

$PanelBotones.Controls.Add($BtnExcel)
$PanelBotones.Controls.Add($BtnPDF)
$PanelBotones.Controls.Add($LblCargando)
$Form.Controls.Add($PanelBotones)

# ----------------- LOGICA DE EXTRACCION DESDE SQLITE -----------------
function Cargar-Datos {
    $LblCargando.Text = "Conectando a la Base de Datos..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $ConnectionString = "Data Source=$RutaDB;Version=3;"
        $Conexion = New-Object System.Data.SQLite.SQLiteConnection($ConnectionString)
        $Conexion.Open()

        # --- MODULO 1: ACCESOS ---
        $LblCargando.Text = "Descargando Historial de Conexiones..."
        [System.Windows.Forms.Application]::DoEvents()

        $Cmd = $Conexion.CreateCommand()
        $Cmd.CommandText = "SELECT * FROM HistorialAccesos ORDER BY Inicio DESC LIMIT 10000"
        $Reader = $Cmd.ExecuteReader()
        while ($Reader.Read()) {
            $Inicio = [datetime]$Reader["Inicio"]
            $FinStr = ""; $HorasActivo = ""
            if ([string]::IsNullOrEmpty($Reader["Fin"])) {
                $Duracion = (Get-Date) - $Inicio
                if ($Duracion.TotalHours -gt 14) { $FinStr = "Cierre no registrado"; $HorasActivo = "Desconocido" }
                else { $FinStr = "Activo actualmente"; $HorasActivo = "{0:00}h {1:00}m (Hasta ahora)" -f $Duracion.Hours, $Duracion.Minutes }
            } else {
                $Fin = [datetime]$Reader["Fin"]
                $FinStr = $Fin.ToString("dd/MM/yyyy HH:mm:ss") + " (" + $Reader["FinRazon"] + ")"
                $Duracion = $Fin - $Inicio
                $HorasActivo = if ($Duracion.TotalSeconds -gt 0) { "{0:00}h {1:00}m {2:00}s" -f $Duracion.Hours, $Duracion.Minutes, $Duracion.Seconds } else { "< 1s" }
            }
            $Global:DatosAccesos += [PSCustomObject]@{
                Usuario = $Reader["Usuario"]
                Fecha = $Inicio.ToString("dd/MM/yyyy")
                'Hora Inicio' = $Inicio.ToString("HH:mm:ss")
                'Hora Cierre / Estado' = $FinStr
                'Tiempo Activo' = $HorasActivo
                'Tipo' = $Reader["Tipo"]
            }
        }
        $Reader.Close()

        # --- MODULO 2: APPS ---
        $LblCargando.Text = "Descargando Historial de Programas..."
        [System.Windows.Forms.Application]::DoEvents()

        $CmdApp = $Conexion.CreateCommand()
        $CmdApp.CommandText = "SELECT * FROM HistorialApps ORDER BY Inicio DESC LIMIT 15000"
        $ReaderApp = $CmdApp.ExecuteReader()
        while ($ReaderApp.Read()) {
            $InicioApp = [datetime]$ReaderApp["Inicio"]
            $FinAppStr = ""; $HorasAppActivo = ""
            if ([string]::IsNullOrEmpty($ReaderApp["Fin"])) {
                $Duracion = (Get-Date) - $InicioApp
                if ($Duracion.TotalHours -gt 14) { $FinAppStr = "Cierre no registrado"; $HorasAppActivo = "Desconocido" }
                else { $FinAppStr = "En Uso"; $HorasAppActivo = "{0:00}h {1:00}m (Hasta ahora)" -f $Duracion.Hours, $Duracion.Minutes }
            } else {
                $FinApp = [datetime]$ReaderApp["Fin"]
                $FinAppStr = $FinApp.ToString("dd/MM/yyyy HH:mm:ss")
                $Duracion = $FinApp - $InicioApp
                $HorasAppActivo = if ($Duracion.TotalSeconds -gt 0) { "{0:00}h {1:00}m {2:00}s" -f $Duracion.Hours, $Duracion.Minutes, $Duracion.Seconds } else { "< 1s" }
            }
            $Global:DatosApps += [PSCustomObject]@{
                Usuario = $ReaderApp["Usuario"]
                Programa = $ReaderApp["Programa"]
                Fecha = $InicioApp.ToString("dd/MM/yyyy")
                'Hora Apertura' = $InicioApp.ToString("HH:mm:ss")
                'Hora Cierre' = $FinAppStr
                'Tiempo de Uso' = $HorasAppActivo
            }
        }
        $ReaderApp.Close()
        $Conexion.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al leer la base de datos: $_", "Error", 0, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }

    # --- MODULO 3: PROGRAMAS EN VIVO ---
    $LblCargando.Text = "Consultando Programas en Vivo..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $Exclusiones = "svchost","System","Idle","smss","csrss","wininit","services","lsass","winlogon","spoolsv","taskhostw","explorer","conhost","dwm","fontdrvhost","sihost","SearchUI","RuntimeBroker","ShellExperienceHost","ctfmon","dllhost","jucheck","jusched","rdpclip","svcr","splwow64","vpprintshell","universalprinter"
        $Procesos = Get-Process -IncludeUserName -ErrorAction SilentlyContinue | Where-Object { $_.UserName -and $_.ProcessName -notin $Exclusiones -and $_.UserName -notmatch '(?i)(SYSTEM|LOCAL SERVICE|NETWORK SERVICE|UMFD-\d+|DWM-\d+|MSSQL.*|SQL.*|defaultuser0)' }
        foreach ($proc in $Procesos) {
            $UsuarioLimpio = $proc.UserName -replace ".*\\", ""
            $HoraApertura = "Desconocida"; $TiempoUso = "N/A"
            if ($proc.StartTime) {
                $HoraApertura = $proc.StartTime.ToString("dd/MM/yyyy HH:mm:ss")
                $DuracionProc = (Get-Date) - $proc.StartTime
                $TiempoUso = if ($DuracionProc.TotalHours -gt 24) { "{0:0} días, {1:00}h {2:00}m" -f $DuracionProc.Days, $DuracionProc.Hours, $DuracionProc.Minutes } else { "{0:00}h {1:00}m" -f $DuracionProc.Hours, $DuracionProc.Minutes }
            }
            $Global:DatosVivos += [PSCustomObject]@{
                Usuario = $UsuarioLimpio
                Programa = $proc.ProcessName
                'Hora de Apertura' = $HoraApertura
                'Estado' = 'Activo'
                'Tiempo de Uso' = $TiempoUso
            }
        }
    } catch {}

    if ($Global:DatosAccesos.Count -gt 0) { $Grid1.DataSource = [System.Collections.ArrayList]($Global:DatosAccesos) }
    if ($Global:DatosApps.Count -gt 0) { $Grid2.DataSource = [System.Collections.ArrayList]($Global:DatosApps) }
    if ($Global:DatosVivos.Count -gt 0) { $Grid3.DataSource = [System.Collections.ArrayList]($Global:DatosVivos | Sort-Object Usuario) }

    $LblCargando.Text = "¡Carga Completa! Datos obtenidos en milisegundos."
}

# ----------------- EXPORTACIONES -----------------
$BtnExcel.Add_Click({
    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Filter = "Archivo Excel CSV (*.csv)|*.csv"
    $SaveDialog.Title = "Guardar Reporte en Excel"
    $SaveDialog.FileName = "ReporteServidor_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
    if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Ruta = $SaveDialog.FileName
        "--- HISTORIAL DE CONEXIONES ---" | Out-File $Ruta -Encoding UTF8
        if ($Global:DatosAccesos) { $Global:DatosAccesos | Export-Csv -Path $Ruta -Append -NoTypeInformation -Encoding UTF8 }
        "`n--- HISTORIAL DE PROGRAMAS ---" | Out-File $Ruta -Encoding UTF8 -Append
        if ($Global:DatosApps) { $Global:DatosApps | Export-Csv -Path $Ruta -Append -NoTypeInformation -Encoding UTF8 }
        "`n--- PROGRAMAS EN VIVO ---" | Out-File $Ruta -Encoding UTF8 -Append
        if ($Global:DatosVivos) { $Global:DatosVivos | Export-Csv -Path $Ruta -Append -NoTypeInformation -Encoding UTF8 }
        
        [System.Windows.Forms.MessageBox]::Show("Exportado exitosamente a Excel (.csv)", "Éxito", 0, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        try { Start-Process $Ruta } catch {}
    }
})

$BtnPDF.Add_Click({
    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Filter = "Documento HTML para Imprimir a PDF (*.html)|*.html"
    $SaveDialog.Title = "Generar Reporte Imprimible"
    $SaveDialog.FileName = "ReporteServidor_$((Get-Date).ToString('yyyyMMdd_HHmm')).html"
    if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Ruta = $SaveDialog.FileName
        $CSS = @"
        <style>
            body { font-family: 'Arial', sans-serif; background-color: #ffffff; color: #1a4350; margin: 0; padding: 20px; }
            h1 { color: #1a4350; text-align: center; border-top: 15px solid #8cc63f; padding-top: 20px; margin-bottom: 30px; font-size: 2.2em; font-weight: bold; }
            h2 { color: #1a4350; margin-top: 30px; padding-bottom: 10px; font-size: 1.5em; text-transform: uppercase; border-bottom: 2px solid #1a4350; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 0.9em; border: 1px solid #1a4350; page-break-inside: avoid; }
            th, td { padding: 8px 10px; text-align: left; border-right: 1px solid #fff; }
            th { background-color: #1a4350; color: #fff; font-weight: bold; }
            tr:nth-child(even) { background-color: #edf5e1; }
            tr:nth-child(odd) { background-color: #ffffff; }
        </style>
"@
        $HTML = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>Reporte para PDF</title>$CSS</head><body>"
        $HTML += "<h1>Reporte de Actividad del Servidor</h1><p>Generado el: $((Get-Date).ToString('dd/MM/yyyy HH:mm:ss'))</p>"
        $HTML += "<h2>1. Historial de Conexiones</h2>"
        if ($Global:DatosAccesos) { $HTML += ($Global:DatosAccesos | ConvertTo-Html -Fragment) } else { $HTML += "<p>No hay datos.</p>" }
        $HTML += "<h2>2. Historial de Programas</h2>"
        if ($Global:DatosApps) { $HTML += ($Global:DatosApps | ConvertTo-Html -Fragment) } else { $HTML += "<p>No hay datos.</p>" }
        $HTML += "<h2>3. Programas en Vivo</h2>"
        if ($Global:DatosVivos) { $HTML += ($Global:DatosVivos | ConvertTo-Html -Fragment) } else { $HTML += "<p>No hay datos.</p>" }
        $HTML += "<script>window.onload = function() { window.print(); }</script></body></html>"
        
        $HTML | Out-File -FilePath $Ruta -Encoding UTF8
        try { Start-Process $Ruta } catch {}
    }
})

$Form.Add_Shown({ Cargar-Datos })
[void]$Form.ShowDialog()
