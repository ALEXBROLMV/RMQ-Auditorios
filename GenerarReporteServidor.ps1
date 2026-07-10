<#
.SYNOPSIS
    Genera un reporte de actividad leyendo la base de datos SQLite con filtros interactivos.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Configuración de base de datos dinámica
$DirectorioExe = [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
if (-not $DirectorioExe) { $DirectorioExe = $PWD.Path }
$RutaBase = Join-Path -Path $DirectorioExe -ChildPath "Recolector"
$RutaDB = Join-Path -Path $RutaBase -ChildPath "Auditoria.db"
$RutaDLL = Join-Path -Path $RutaBase -ChildPath "System.Data.SQLite.dll"

if (-not (Test-Path $RutaDLL)) {
    [System.Windows.Forms.MessageBox]::Show("Falta el motor SQLite en $RutaBase.", "Error Crítico", 0, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit
}
if (-not (Test-Path $RutaDB)) {
    [System.Windows.Forms.MessageBox]::Show("La base de datos Auditoria.db no existe. Corre Recolector.ps1 primero.", "Advertencia", 0, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    exit
}

Add-Type -Path $RutaDLL
$ConnectionString = "Data Source=$RutaDB;Version=3;"
$Global:Conexion = New-Object System.Data.SQLite.SQLiteConnection($ConnectionString)
$Global:Conexion.Open()

# Listas de exportación
$Global:DatosAccesos = @()
$Global:DatosApps = @()
$Global:DatosVivos = @()

# ----------------- UI -----------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Reporte de Actividad del Servidor - $env:COMPUTERNAME"
$Form.Size = New-Object System.Drawing.Size(1000, 750)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::White

# PANEL SUPERIOR (Filtros)
$PanelFiltros = New-Object System.Windows.Forms.Panel
$PanelFiltros.Dock = "Top"
$PanelFiltros.Height = 80
$PanelFiltros.BackColor = [System.Drawing.Color]::FromArgb(26, 67, 80)

$RutaLogo = Join-Path -Path $DirectorioExe -ChildPath "LOGORMQ.png"
if (Test-Path $RutaLogo) {
    $LogoBox = New-Object System.Windows.Forms.PictureBox
    $LogoBox.Image = [System.Drawing.Image]::FromFile($RutaLogo)
    $LogoBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $LogoBox.Size = New-Object System.Drawing.Size(120, 60)
    $LogoBox.Location = New-Object System.Drawing.Point(850, 10)
    $LogoBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $PanelFiltros.Controls.Add($LogoBox)
}

$LblFecha = New-Object System.Windows.Forms.Label
$LblFecha.Text = "1. Selecciona el Día:"
$LblFecha.ForeColor = [System.Drawing.Color]::White
$LblFecha.Location = New-Object System.Drawing.Point(20, 15)
$LblFecha.AutoSize = $true
$LblFecha.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

$Calendario = New-Object System.Windows.Forms.DateTimePicker
$Calendario.Location = New-Object System.Drawing.Point(20, 40)
$Calendario.Width = 220
$Calendario.Format = [System.Windows.Forms.DateTimePickerFormat]::Short

$LblUser = New-Object System.Windows.Forms.Label
$LblUser.Text = "2. Selecciona el Usuario:"
$LblUser.ForeColor = [System.Drawing.Color]::White
$LblUser.Location = New-Object System.Drawing.Point(260, 15)
$LblUser.AutoSize = $true
$LblUser.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

$ComboUsuarios = New-Object System.Windows.Forms.ComboBox
$ComboUsuarios.Location = New-Object System.Drawing.Point(260, 40)
$ComboUsuarios.Width = 250
$ComboUsuarios.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$LblCargandoFiltro = New-Object System.Windows.Forms.Label
$LblCargandoFiltro.Location = New-Object System.Drawing.Point(530, 43)
$LblCargandoFiltro.AutoSize = $true
$LblCargandoFiltro.ForeColor = [System.Drawing.Color]::FromArgb(140, 198, 63)
$LblCargandoFiltro.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Italic)

$PanelFiltros.Controls.Add($LblFecha)
$PanelFiltros.Controls.Add($Calendario)
$PanelFiltros.Controls.Add($LblUser)
$PanelFiltros.Controls.Add($ComboUsuarios)
$PanelFiltros.Controls.Add($LblCargandoFiltro)
$Form.Controls.Add($PanelFiltros)

# PESTAÑAS (Tablas)
$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Dock = "Fill"
$TabControl.Font = New-Object System.Drawing.Font("Arial", 10)

$Tab1 = New-Object System.Windows.Forms.TabPage
$Tab1.Text = "Historial de Conexiones"
$Grid1 = New-Object System.Windows.Forms.DataGridView
$Grid1.Dock = "Fill"
$Grid1.ReadOnly = $true
$Grid1.AllowUserToAddRows = $false
$Grid1.AutoSizeColumnsMode = "Fill"
$Grid1.BackgroundColor = [System.Drawing.Color]::White
$Tab1.Controls.Add($Grid1)

$Tab2 = New-Object System.Windows.Forms.TabPage
$Tab2.Text = "Programas Abiertos en ese día"
$Grid2 = New-Object System.Windows.Forms.DataGridView
$Grid2.Dock = "Fill"
$Grid2.ReadOnly = $true
$Grid2.AllowUserToAddRows = $false
$Grid2.AutoSizeColumnsMode = "Fill"
$Grid2.BackgroundColor = [System.Drawing.Color]::White
$Tab2.Controls.Add($Grid2)

$TabControl.Controls.Add($Tab1)
$TabControl.Controls.Add($Tab2)
$Form.Controls.Add($TabControl)

# PANEL INFERIOR (Botones)
$PanelBotones = New-Object System.Windows.Forms.Panel
$PanelBotones.Dock = "Bottom"
$PanelBotones.Height = 60
$PanelBotones.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

$BtnExcel = New-Object System.Windows.Forms.Button
$BtnExcel.Text = "Exportar Usuario a Excel"
$BtnExcel.Location = New-Object System.Drawing.Point(20, 10)
$BtnExcel.Size = New-Object System.Drawing.Size(200, 40)
$BtnExcel.BackColor = [System.Drawing.Color]::FromArgb(140, 198, 63)
$BtnExcel.ForeColor = [System.Drawing.Color]::White
$BtnExcel.FlatStyle = "Flat"
$BtnExcel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)

$BtnPDF = New-Object System.Windows.Forms.Button
$BtnPDF.Text = "Exportar a PDF / Imprimir"
$BtnPDF.Location = New-Object System.Drawing.Point(240, 10)
$BtnPDF.Size = New-Object System.Drawing.Size(200, 40)
$BtnPDF.BackColor = [System.Drawing.Color]::FromArgb(140, 198, 63)
$BtnPDF.ForeColor = [System.Drawing.Color]::White
$BtnPDF.FlatStyle = "Flat"
$BtnPDF.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)

$PanelBotones.Controls.Add($BtnExcel)
$PanelBotones.Controls.Add($BtnPDF)
$Form.Controls.Add($PanelBotones)

$PanelFiltros.SendToBack()
$PanelBotones.SendToBack()
$TabControl.BringToFront()

# ----------------- EVENTOS Y LÓGICA -----------------

function Buscar-UsuariosPorFecha {
    $FechaSeleccionada = $Calendario.Value.ToString("yyyy-MM-dd")
    $ComboUsuarios.Items.Clear()
    $Grid1.DataSource = $null
    $Grid2.DataSource = $null
    $LblCargandoFiltro.Text = "Buscando usuarios..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $Cmd = $Global:Conexion.CreateCommand()
        $Cmd.CommandText = "SELECT DISTINCT Usuario FROM HistorialAccesos WHERE date(Inicio) = @Fecha UNION SELECT DISTINCT Usuario FROM HistorialApps WHERE date(Inicio) = @Fecha ORDER BY Usuario"
        $Cmd.Parameters.AddWithValue("@Fecha", $FechaSeleccionada) | Out-Null
        $Reader = $Cmd.ExecuteReader()
        
        [void]$ComboUsuarios.Items.Add("-- TODOS --")
        
        while ($Reader.Read()) {
            [void]$ComboUsuarios.Items.Add($Reader["Usuario"].ToString())
        }
        $Reader.Close()

        if ($ComboUsuarios.Items.Count -gt 0) {
            $LblCargandoFiltro.Text = "Usuarios encontrados."
            $ComboUsuarios.SelectedIndex = 0 # Selecciona al primero y dispara el evento de las tablas
        } else {
            $LblCargandoFiltro.Text = "Nadie se conectó en esta fecha."
        }
    } catch {
        $LblCargandoFiltro.Text = "Error al buscar usuarios."
    }
}

function Cargar-Tablas {
    if ($ComboUsuarios.SelectedItem -eq $null) { return }
    $FechaSel = $Calendario.Value.ToString("yyyy-MM-dd")
    $UserSel = $ComboUsuarios.SelectedItem.ToString()
    $LblCargandoFiltro.Text = "Cargando historial de $UserSel..."
    [System.Windows.Forms.Application]::DoEvents()

    $Global:DatosAccesos = @()
    $Global:DatosApps = @()

    try {
        # Conexiones
        $FiltroUser = if ($UserSel -ne "-- TODOS --") { " AND Usuario = @User" } else { "" }
        
        $Cmd = $Global:Conexion.CreateCommand()
        $Cmd.CommandText = "SELECT * FROM HistorialAccesos WHERE date(Inicio) = @Fecha$FiltroUser ORDER BY Inicio DESC"
        $Cmd.Parameters.AddWithValue("@Fecha", $FechaSel) | Out-Null
        if ($UserSel -ne "-- TODOS --") { $Cmd.Parameters.AddWithValue("@User", $UserSel) | Out-Null }
        $Reader = $Cmd.ExecuteReader()
        while ($Reader.Read()) {
            $Inicio = [datetime]$Reader["Inicio"]
            $FinStr = ""; $HorasActivo = ""
            if ([string]::IsNullOrEmpty($Reader["Fin"])) {
                $Duracion = (Get-Date) - $Inicio
                if ($Duracion.TotalHours -gt 14) { $FinStr = "Cierre no registrado"; $HorasActivo = "N/A" }
                else { $FinStr = "Activo actualmente"; $HorasActivo = "{0:00}h {1:00}m" -f $Duracion.Hours, $Duracion.Minutes }
            } else {
                $Fin = [datetime]$Reader["Fin"]
                $FinStr = $Fin.ToString("HH:mm:ss") + " (" + $Reader["FinRazon"] + ")"
                $Duracion = $Fin - $Inicio
                $HorasActivo = if ($Duracion.TotalSeconds -gt 0) { "{0:00}h {1:00}m {2:00}s" -f $Duracion.Hours, $Duracion.Minutes, $Duracion.Seconds } else { "< 1s" }
            }
            $Global:DatosAccesos += [PSCustomObject]@{
                'Usuario' = $Reader["Usuario"]
                'Hora Conexión' = $Inicio.ToString("HH:mm:ss")
                'Hora Desconexión' = $FinStr
                'Tiempo de Sesión' = $HorasActivo
                'Tipo' = $Reader["Tipo"]
            }
        }
        $Reader.Close()

        # Apps
        $CmdApp = $Global:Conexion.CreateCommand()
        $CmdApp.CommandText = "SELECT * FROM HistorialApps WHERE date(Inicio) = @Fecha$FiltroUser ORDER BY Inicio DESC"
        $CmdApp.Parameters.AddWithValue("@Fecha", $FechaSel) | Out-Null
        if ($UserSel -ne "-- TODOS --") { $CmdApp.Parameters.AddWithValue("@User", $UserSel) | Out-Null }
        $ReaderApp = $CmdApp.ExecuteReader()
        while ($ReaderApp.Read()) {
            $InicioApp = [datetime]$ReaderApp["Inicio"]
            $FinAppStr = ""; $HorasAppActivo = ""
            if ([string]::IsNullOrEmpty($ReaderApp["Fin"])) {
                $Duracion = (Get-Date) - $InicioApp
                if ($Duracion.TotalHours -gt 14) { $FinAppStr = "Cierre no registrado"; $HorasAppActivo = "N/A" }
                else { $FinAppStr = "En Uso (Activo)"; $HorasAppActivo = "{0:00}h {1:00}m" -f $Duracion.Hours, $Duracion.Minutes }
            } else {
                $FinApp = [datetime]$ReaderApp["Fin"]
                $FinAppStr = $FinApp.ToString("HH:mm:ss")
                $Duracion = $FinApp - $InicioApp
                $HorasAppActivo = if ($Duracion.TotalSeconds -gt 0) { "{0:00}h {1:00}m {2:00}s" -f $Duracion.Hours, $Duracion.Minutes, $Duracion.Seconds } else { "< 1s" }
            }
            $Global:DatosApps += [PSCustomObject]@{
                'Usuario' = $ReaderApp["Usuario"]
                'Programa Abierto' = $ReaderApp["Programa"]
                'Hora Apertura' = $InicioApp.ToString("HH:mm:ss")
                'Hora Cierre' = $FinAppStr
                'Tiempo en Uso' = $HorasAppActivo
            }
        }
        $ReaderApp.Close()

        if ($Global:DatosAccesos.Count -gt 0) { $Grid1.DataSource = [System.Collections.ArrayList]($Global:DatosAccesos) } else { $Grid1.DataSource = $null }
        if ($Global:DatosApps.Count -gt 0) { $Grid2.DataSource = [System.Collections.ArrayList]($Global:DatosApps) } else { $Grid2.DataSource = $null }
        
        $LblCargandoFiltro.Text = "Mostrando registros de $UserSel"
    } catch {
        $LblCargandoFiltro.Text = "Error al cargar datos."
    }
}

$Calendario.Add_ValueChanged({ Buscar-UsuariosPorFecha })
$ComboUsuarios.Add_SelectedIndexChanged({ Cargar-Tablas })

# ----------------- EXPORTACIONES -----------------
$BtnExcel.Add_Click({
    if ($ComboUsuarios.SelectedItem -eq $null) { return }
    $UserSel = $ComboUsuarios.SelectedItem.ToString()
    $FechaSel = $Calendario.Value.ToString("yyyyMMdd")
    
    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Filter = "Archivo Excel CSV (*.csv)|*.csv"
    $SaveDialog.Title = "Guardar Reporte de Usuario en Excel"
    $SaveDialog.FileName = "Reporte_$( $UserSel )_$FechaSel.csv"
    
    if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Ruta = $SaveDialog.FileName
        "--- CONEXIONES DE $( $UserSel ) EL $( $Calendario.Value.ToString('dd/MM/yyyy') ) ---" | Out-File $Ruta -Encoding UTF8
        if ($Global:DatosAccesos) { $Global:DatosAccesos | Export-Csv -Path $Ruta -Append -NoTypeInformation -Encoding UTF8 }
        
        "`n--- PROGRAMAS UTILIZADOS ---" | Out-File $Ruta -Encoding UTF8 -Append
        if ($Global:DatosApps) { $Global:DatosApps | Export-Csv -Path $Ruta -Append -NoTypeInformation -Encoding UTF8 }
        
        [System.Windows.Forms.MessageBox]::Show("Exportado exitosamente a Excel (.csv)", "Éxito", 0, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        try { Start-Process $Ruta } catch {}
    }
})

$BtnPDF.Add_Click({
    if ($ComboUsuarios.SelectedItem -eq $null) { return }
    $UserSel = $ComboUsuarios.SelectedItem.ToString()
    $FechaSel = $Calendario.Value.ToString("yyyyMMdd")

    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Filter = "Documento HTML para Imprimir a PDF (*.html)|*.html"
    $SaveDialog.Title = "Generar Reporte Imprimible de Usuario"
    $SaveDialog.FileName = "Reporte_$( $UserSel )_$FechaSel.html"
    
    if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Ruta = $SaveDialog.FileName
        $CSS = @"
        <style>
            body { font-family: 'Arial', sans-serif; background-color: #ffffff; color: #1a4350; margin: 0; padding: 20px; }
            h1 { color: #1a4350; text-align: center; border-top: 15px solid #8cc63f; padding-top: 20px; margin-bottom: 5px; font-size: 2.2em; font-weight: bold; }
            h3 { color: #555; text-align: center; margin-bottom: 30px; }
            h2 { color: #1a4350; margin-top: 30px; padding-bottom: 10px; font-size: 1.5em; text-transform: uppercase; border-bottom: 2px solid #1a4350; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 0.9em; border: 1px solid #1a4350; page-break-inside: avoid; }
            th, td { padding: 8px 10px; text-align: left; border-right: 1px solid #fff; }
            th { background-color: #1a4350; color: #fff; font-weight: bold; }
            tr:nth-child(even) { background-color: #edf5e1; }
            tr:nth-child(odd) { background-color: #ffffff; }
        </style>
"@
        $HTML = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>Reporte de Usuario</title>$CSS</head><body>"
        $HTML += "<h1>Auditoría de Usuario: $UserSel</h1>"
        $HTML += "<h3>Fecha Auditada: $( $Calendario.Value.ToString('dd/MM/yyyy') )</h3>"
        
        $HTML += "<h2>1. Historial de Conexiones</h2>"
        if ($Global:DatosAccesos) { $HTML += ($Global:DatosAccesos | ConvertTo-Html -Fragment) } else { $HTML += "<p>No hay conexiones registradas este día.</p>" }
        
        $HTML += "<h2>2. Programas Específicos Utilizados</h2>"
        if ($Global:DatosApps) { $HTML += ($Global:DatosApps | ConvertTo-Html -Fragment) } else { $HTML += "<p>No hay programas registrados este día.</p>" }
        
        $HTML += "<script>window.onload = function() { window.print(); }</script></body></html>"
        
        $HTML | Out-File -FilePath $Ruta -Encoding UTF8
        try { Start-Process $Ruta } catch {}
    }
})

$Form.Add_Shown({ Buscar-UsuariosPorFecha })
$Form.Add_FormClosed({ $Global:Conexion.Close() })
[void]$Form.ShowDialog()
