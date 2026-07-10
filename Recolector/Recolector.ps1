<#
.SYNOPSIS
    Recolector Automático de Eventos para Base de Datos SQLite.
    Se recomienda programarlo en el Programador de Tareas para ejecutarse cada hora.
#>

$RutaBase = $PSScriptRoot
if (-not $RutaBase) { $RutaBase = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path) }
$RutaDB = Join-Path -Path $RutaBase -ChildPath "Auditoria.db"
$RutaDLL = Join-Path -Path $RutaBase -ChildPath "System.Data.SQLite.dll"

# Cargar el DLL de SQLite
Add-Type -Path $RutaDLL

# Crear string de conexión
$ConnectionString = "Data Source=$RutaDB;Version=3;"
$Conexion = New-Object System.Data.SQLite.SQLiteConnection($ConnectionString)
$Conexion.Open()

# 1. Crear las Tablas si no existen
$Cmd = $Conexion.CreateCommand()
$Cmd.CommandText = @"
    CREATE TABLE IF NOT EXISTS HistorialAccesos (
        RecordId INTEGER PRIMARY KEY,
        Usuario TEXT,
        Inicio DATETIME,
        Fin DATETIME,
        FinRazon TEXT,
        Tipo TEXT
    );
    CREATE TABLE IF NOT EXISTS HistorialApps (
        RecordId INTEGER PRIMARY KEY,
        Usuario TEXT,
        Programa TEXT,
        Inicio DATETIME,
        Fin DATETIME
    );
"@
$Cmd.ExecuteNonQuery() | Out-Null

# --- HELPER: Nombres Amigables ---
function Obtener-NombreAmigable ($procName) {
    $n = $procName.ToLower().Replace(".exe", "")
    switch -Regex ($n) {
        "contabilidad" { return "CONTPAQi Contabilidad" }
        "nomipaq" { return "CONTPAQi Nóminas" }
        "comercial" { return "CONTPAQi Comercial" }
        "adminpaq" { return "CONTPAQi AdminPAQ" }
        "factura" { return "CONTPAQi Facturación" }
        "chrome" { return "Google Chrome" }
        "msedge" { return "Microsoft Edge" }
        "anydesk" { return "AnyDesk" }
        "teamviewer" { return "TeamViewer" }
        "excel" { return "Microsoft Excel" }
        "winword" { return "Microsoft Word" }
        "outlook" { return "Microsoft Outlook" }
        default { return (Get-Culture).TextInfo.ToTitleCase($n) }
    }
}

# ----------------- RECOLECTAR ACCESOS -----------------
try {
    # Extraer de los últimos 2 días para actualizar sesiones que cerraron tarde
    $Eventos = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4624, 4634, 4778, 4779; StartTime=(Get-Date).AddDays(-2)} -ErrorAction SilentlyContinue
    $EventosPorUsuario = @{}
    
    foreach ($ev in $Eventos) {
        $xml = [xml]$ev.ToXml()
        $data = $xml.Event.EventData.Data
        $Usuario = ($data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
        
        if ($Usuario -and $Usuario -notmatch '(?i)(SYSTEM|LOCAL SERVICE|NETWORK SERVICE|UMFD-\d+|DWM-\d+|MSSQL.*|SQLTELEMETRY.*|defaultuser0|ANONYMOUS LOGON)' -and $Usuario -notlike "*$") {
            if (-not $EventosPorUsuario.ContainsKey($Usuario)) { $EventosPorUsuario[$Usuario] = @() }
            $Tipo = "Local"
            if ($ev.Id -in 4624, 4634) {
                $logonType = ($data | Where-Object { $_.Name -eq 'LogonType' }).'#text'
                if ($logonType -in '10', '3') { $Tipo = "Remoto (RDP/Red)" }
            } else {
                $Tipo = "Remoto (RDP)"
            }
            $EventosPorUsuario[$Usuario] += [PSCustomObject]@{ Id = $ev.Id; Time = $ev.TimeCreated; Tipo = $Tipo; RecordId = $ev.RecordId }
        }
    }

    $ListaFinalAccesos = @()
    foreach ($User in $EventosPorUsuario.Keys) {
        $Timeline = $EventosPorUsuario[$User] | Sort-Object Time
        $SesionActual = $null
        
        foreach ($ev in $Timeline) {
            if ($ev.Id -in 4624, 4778) {
                if ($SesionActual) {
                    $SesionActual.Fin = $ev.Time; $SesionActual.FinRazon = "Reconexion"; $ListaFinalAccesos += $SesionActual
                }
                $SesionActual = [PSCustomObject]@{ RecordId = $ev.RecordId; Usuario = $User; Inicio = $ev.Time; Fin = $null; FinRazon = ""; Tipo = $ev.Tipo }
            } elseif ($ev.Id -in 4634, 4779) {
                if ($SesionActual) {
                    $SesionActual.Fin = $ev.Time; $SesionActual.FinRazon = if ($ev.Id -eq 4634) { "Cierre Normal" } else { "Desconectado" }
                    $ListaFinalAccesos += $SesionActual
                    $SesionActual = $null
                }
            }
        }
        if ($SesionActual) { $ListaFinalAccesos += $SesionActual }
    }

    # Guardar en SQLite (UPSERT: Actualiza si ya existe el RecordId)
    $Tx = $Conexion.BeginTransaction()
    $CmdInsert = $Conexion.CreateCommand()
    $CmdInsert.CommandText = @"
        INSERT INTO HistorialAccesos (RecordId, Usuario, Inicio, Fin, FinRazon, Tipo)
        VALUES (@RecordId, @Usuario, @Inicio, @Fin, @FinRazon, @Tipo)
        ON CONFLICT(RecordId) DO UPDATE SET 
            Fin=excluded.Fin, FinRazon=excluded.FinRazon;
"@
    foreach ($v in $ListaFinalAccesos) {
        $CmdInsert.Parameters.Clear() | Out-Null
        $CmdInsert.Parameters.AddWithValue("@RecordId", $v.RecordId) | Out-Null
        $CmdInsert.Parameters.AddWithValue("@Usuario", $v.Usuario) | Out-Null
        $CmdInsert.Parameters.AddWithValue("@Inicio", $v.Inicio.ToString("yyyy-MM-dd HH:mm:ss")) | Out-Null
        $CmdInsert.Parameters.AddWithValue("@Fin", $(if ($v.Fin) { $v.Fin.ToString("yyyy-MM-dd HH:mm:ss") } else { [DBNull]::Value })) | Out-Null
        $CmdInsert.Parameters.AddWithValue("@FinRazon", $v.FinRazon) | Out-Null
        $CmdInsert.Parameters.AddWithValue("@Tipo", $v.Tipo) | Out-Null
        $CmdInsert.ExecuteNonQuery() | Out-Null
    }
    $Tx.Commit()
} catch {
    Write-Host "Error guardando Accesos: $_"
}

# ----------------- RECOLECTAR APPS -----------------
try {
    $EventosProcesos = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4688, 4689; StartTime=(Get-Date).AddDays(-2)} -ErrorAction SilentlyContinue
    $HistorialApp = @{}

    foreach ($ev in $EventosProcesos) {
        $xml = [xml]$ev.ToXml()
        $data = $xml.Event.EventData.Data
        $processId = ($data | Where-Object { $_.Name -in 'NewProcessId', 'ProcessId' } | Select-Object -First 1).'#text'
        $procNamePath = ($data | Where-Object { $_.Name -in 'NewProcessName', 'ProcessName' } | Select-Object -First 1).'#text'
        $user = ($data | Where-Object { $_.Name -eq 'SubjectUserName' }).'#text'

        if ($procNamePath -match '(?i)(contpaq|contabilidad|nomina|comercial|adminpaq|factura|chrome|msedge|anydesk|teamviewer|excel|winword|outlook)') {
            $procNameRaw = [System.IO.Path]::GetFileName($procNamePath)
            $procName = Obtener-NombreAmigable $procNameRaw
            
            $key = "$processId-$user-$procName"
            if (-not $HistorialApp.ContainsKey($key)) { 
                $HistorialApp[$key] = @{ RecordId = $ev.RecordId; Usuario = $user; Programa = $procName; Inicio = $null; Fin = $null } 
            }
            if ($ev.Id -eq 4688) {
                $HistorialApp[$key].Inicio = $ev.TimeCreated
                $HistorialApp[$key].RecordId = $ev.RecordId
            } elseif ($ev.Id -eq 4689) {
                if (-not $HistorialApp[$key].Fin) { $HistorialApp[$key].Fin = $ev.TimeCreated }
            }
        }
    }

    $Tx2 = $Conexion.BeginTransaction()
    $CmdInsertApp = $Conexion.CreateCommand()
    $CmdInsertApp.CommandText = @"
        INSERT INTO HistorialApps (RecordId, Usuario, Programa, Inicio, Fin)
        VALUES (@RecordId, @Usuario, @Programa, @Inicio, @Fin)
        ON CONFLICT(RecordId) DO UPDATE SET 
            Fin=excluded.Fin;
"@
    foreach ($Llave in $HistorialApp.Keys) {
        $App = $HistorialApp[$Llave]
        if ($App.Inicio) {
            $CmdInsertApp.Parameters.Clear() | Out-Null
            $CmdInsertApp.Parameters.AddWithValue("@RecordId", $App.RecordId) | Out-Null
            $CmdInsertApp.Parameters.AddWithValue("@Usuario", $App.Usuario) | Out-Null
            $CmdInsertApp.Parameters.AddWithValue("@Programa", $App.Programa) | Out-Null
            $CmdInsertApp.Parameters.AddWithValue("@Inicio", $App.Inicio.ToString("yyyy-MM-dd HH:mm:ss")) | Out-Null
            $CmdInsertApp.Parameters.AddWithValue("@Fin", $(if ($App.Fin) { $App.Fin.ToString("yyyy-MM-dd HH:mm:ss") } else { [DBNull]::Value })) | Out-Null
            $CmdInsertApp.ExecuteNonQuery() | Out-Null
        }
    }
    $Tx2.Commit()
} catch {
    Write-Host "Error guardando Programas: $_"
}

$Conexion.Close()
