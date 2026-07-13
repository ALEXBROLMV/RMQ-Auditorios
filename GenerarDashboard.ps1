<#
.SYNOPSIS
    Genera un Dashboard HTML interactivo a partir de la base de datos SQLite.
#>

$DirectorioActual = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($DirectorioActual)) { 
    $DirectorioActual = $PWD.Path 
}

$RutaBase = Join-Path -Path $DirectorioActual -ChildPath "Recolector"
$RutaDB = Join-Path -Path $RutaBase -ChildPath "Auditoria.db"
$RutaDLL = Join-Path -Path $RutaBase -ChildPath "System.Data.SQLite.dll"

if (-not (Test-Path $RutaDLL) -or -not (Test-Path $RutaDB)) {
    Write-Host "Error: Faltan archivos de la base de datos."
    exit
}

Add-Type -Path $RutaDLL
$Conexion = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$RutaDB;Version=3;")
$Conexion.Open()

# Extraer Datos
$Accesos = @()
$Cmd = $Conexion.CreateCommand()
$Cmd.CommandText = "SELECT * FROM HistorialAccesos ORDER BY Inicio DESC"
$Reader = $Cmd.ExecuteReader()
while ($Reader.Read()) {
    $Fin = if ([string]::IsNullOrEmpty($Reader["Fin"])) { $null } else { [datetime]$Reader["Fin"] }
    $Accesos += @{
        Usuario = $Reader["Usuario"].ToString()
        Inicio = ([datetime]$Reader["Inicio"]).ToString("yyyy-MM-ddTHH:mm:ss")
        Fin = if ($Fin) { $Fin.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
        FinRazon = $Reader["FinRazon"].ToString()
        Tipo = $Reader["Tipo"].ToString()
    }
}
$Reader.Close()

$Apps = @()
$CmdApp = $Conexion.CreateCommand()
$CmdApp.CommandText = "SELECT * FROM HistorialApps ORDER BY Inicio DESC"
$ReaderApp = $CmdApp.ExecuteReader()
while ($ReaderApp.Read()) {
    $Fin = if ([string]::IsNullOrEmpty($ReaderApp["Fin"])) { $null } else { [datetime]$ReaderApp["Fin"] }
    $Apps += @{
        Usuario = $ReaderApp["Usuario"].ToString()
        Programa = $ReaderApp["Programa"].ToString()
        Inicio = ([datetime]$ReaderApp["Inicio"]).ToString("yyyy-MM-ddTHH:mm:ss")
        Fin = if ($Fin) { $Fin.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
    }
}
$ReaderApp.Close()
$Conexion.Close()

$JSONData = @{ Accesos = $Accesos; Apps = $Apps } | ConvertTo-Json -Depth 5 -Compress

# --- HTML TEMPLATE ---
$HTML = @'
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Dashboard de Auditoría</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&display=swap');
        
        :root {
            --bg-main: #0f172a;
            --bg-panel: rgba(30, 41, 59, 0.7);
            --primary: #10b981;
            --secondary: #0ea5e9;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --danger: #ef4444;
        }

        body {
            margin: 0; padding: 0;
            font-family: 'Inter', sans-serif;
            background: var(--bg-main);
            color: var(--text-main);
            display: flex; height: 100vh; overflow: hidden;
        }

        /* Sidebar */
        .sidebar {
            width: 300px;
            background: var(--bg-panel);
            backdrop-filter: blur(10px);
            border-right: 1px solid rgba(255,255,255,0.1);
            display: flex; flex-direction: column;
        }
        
        .logo-area {
            padding: 20px; text-align: center; border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .logo-area h1 { margin: 0; font-size: 1.2rem; color: var(--primary); text-transform: uppercase; letter-spacing: 1px;}
        
        .filter-area { padding: 20px; }
        .filter-area label { display: block; font-size: 0.8rem; color: var(--text-muted); margin-bottom: 5px; }
        .filter-area input[type="date"] {
            width: 100%; padding: 10px; border-radius: 8px; border: none; background: #334155; color: white; font-family: 'Inter'; box-sizing: border-box;
        }

        .user-list { flex: 1; overflow-y: auto; padding: 10px; }
        .user-item {
            padding: 15px; margin-bottom: 10px; border-radius: 8px; background: rgba(255,255,255,0.03); cursor: pointer; transition: all 0.2s; border: 1px solid transparent;
        }
        .user-item:hover { background: rgba(255,255,255,0.1); }
        .user-item.active { border-color: var(--primary); background: rgba(16, 185, 129, 0.1); }
        .user-name { font-weight: 600; font-size: 1rem; }
        .user-meta { font-size: 0.8rem; color: var(--text-muted); margin-top: 5px; }

        /* Main Content */
        .main-content {
            flex: 1; padding: 30px; overflow-y: auto;
        }
        
        .header-dash { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; }
        .header-dash h2 { margin: 0; font-size: 2rem; font-weight: 800; }
        .header-dash p { margin: 5px 0 0 0; color: var(--text-muted); }

        .metrics-grid {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px;
        }
        .metric-card {
            background: var(--bg-panel); border: 1px solid rgba(255,255,255,0.1); padding: 20px; border-radius: 12px;
        }
        .metric-value { font-size: 2rem; font-weight: 800; color: var(--primary); margin-bottom: 5px;}
        .metric-label { font-size: 0.85rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 1px;}

        .section-title { font-size: 1.2rem; font-weight: 600; margin: 30px 0 15px 0; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 10px; }

        /* Timeline */
        .timeline-container { background: var(--bg-panel); padding: 20px; border-radius: 12px; margin-bottom: 30px; }
        .timeline-item { display: flex; margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px dashed rgba(255,255,255,0.1); align-items: center; }
        .timeline-item:last-child { border-bottom: none; margin-bottom: 0; padding-bottom: 0; }
        .time-box { width: 100px; font-weight: bold; color: var(--secondary); font-size: 0.9rem; }
        .event-box { flex: 1; }
        .event-title { font-weight: 600; font-size: 1rem; }
        .event-dur { font-size: 0.8rem; color: var(--text-muted); margin-top: 3px; display:inline-block; background: rgba(255,255,255,0.1); padding: 2px 8px; border-radius: 4px; margin-left: 10px;}
        .status-active { color: var(--primary); font-size: 0.8rem; font-weight: bold; margin-left: 10px; text-transform: uppercase; animation: pulse 2s infinite; }
        
        @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.5; } 100% { opacity: 1; } }

        /* App Cards */
        .app-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 15px; }
        .app-card { background: rgba(14, 165, 233, 0.1); border-left: 4px solid var(--secondary); padding: 15px; border-radius: 8px; }
        .app-name { font-weight: 600; margin-bottom: 5px; }
        .app-time { font-size: 0.85rem; color: var(--text-muted); }

        .empty-state { text-align: center; color: var(--text-muted); margin-top: 100px; }
        .empty-state h3 { font-size: 1.5rem; color: white; margin-bottom: 10px;}

        /* Scrollbar */
        ::-webkit-scrollbar { width: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.2); border-radius: 4px; }
    </style>
</head>
<body>

    <div class="sidebar">
        <div class="logo-area">
            <h1>Auditoría RMQ</h1>
        </div>
        <div class="filter-area">
            <label>Seleccionar Fecha:</label>
            <input type="date" id="datePicker">
        </div>
        <div class="user-list" id="userList">
            <!-- Users injected here -->
        </div>
    </div>

    <div class="main-content" id="mainContent">
        <div class="empty-state">
            <h3>Bienvenido al Dashboard</h3>
            <p>Selecciona una fecha en el panel izquierdo para ver la actividad del servidor.</p>
        </div>
    </div>

    <script>
        const DB = $JSONData;
        if (DB.Accesos && !Array.isArray(DB.Accesos)) DB.Accesos = [DB.Accesos];
        if (!DB.Accesos) DB.Accesos = [];
        if (DB.Apps && !Array.isArray(DB.Apps)) DB.Apps = [DB.Apps];
        if (!DB.Apps) DB.Apps = [];

        
        const datePicker = document.getElementById('datePicker');
        const userList = document.getElementById('userList');
        const mainContent = document.getElementById('mainContent');

        // Initialize with today
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        const todayStr = yyyy + "-" + mm + "-" + dd;
        
        datePicker.value = todayStr;

        datePicker.addEventListener('change', renderUsers);

        function renderUsers() {
            const selectedDate = datePicker.value;
            userList.innerHTML = '';
            
            // Find users active on this date
            const activeUsers = new Set();
            
            DB.Accesos.forEach(a => { if (a.Inicio.startsWith(selectedDate)) activeUsers.add(a.Usuario); });
            DB.Apps.forEach(a => { if (a.Inicio.startsWith(selectedDate)) activeUsers.add(a.Usuario); });

            const sortedUsers = Array.from(activeUsers).sort();

            if (sortedUsers.length === 0) {
                userList.innerHTML = '<div style="text-align:center; color: #94a3b8; margin-top: 20px;">No hay actividad este día.</div>';
                mainContent.innerHTML = '<div class="empty-state"><h3>Sin Registros</h3><p>Nadie ingresó al servidor en la fecha seleccionada.</p></div>';
                return;
            }

            sortedUsers.forEach((user, idx) => {
                const div = document.createElement('div');
                div.className = 'user-item';
                div.innerHTML = `<div class="user-name">${user}</div><div class="user-meta">Clic para expandir</div>`;
                div.onclick = () => {
                    document.querySelectorAll('.user-item').forEach(el => el.classList.remove('active'));
                    div.classList.add('active');
                    renderDashboard(user, selectedDate);
                };
                userList.appendChild(div);
                
                // Auto-select first user
                if (idx === 0) div.click();
            });
        }

        function formatDuration(start, end) {
            const s = new Date(start);
            const e = end ? new Date(end) : new Date();
            const diffMs = e - s;
            if (diffMs < 0) return "< 1m";
            const diffHrs = Math.floor(diffMs / 3600000);
            const diffMins = Math.floor((diffMs % 3600000) / 60000);
            return `${diffHrs}h ${diffMins}m`;
        }

        function formatTime(dateStr) {
            return new Date(dateStr).toLocaleTimeString('es-MX', { hour: '2-digit', minute:'2-digit' });
        }

        function renderDashboard(user, dateStr) {
            // Filter Data
            const userAccesos = DB.Accesos.filter(a => a.Usuario === user && a.Inicio.startsWith(dateStr));
            const userApps = DB.Apps.filter(a => a.Usuario === user && a.Inicio.startsWith(dateStr));

            let totalConexiones = userAccesos.length;
            let totalApps = userApps.length;
            
            // Render Header
            let html = `
                <div class="header-dash">
                    <div>
                        <h2>${user}</h2>
                        <p>Actividad registrada el ${dateStr}</p>
                    </div>
                </div>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <div class="metric-value">${totalConexiones}</div>
                        <div class="metric-label">Inicios de Sesión</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">${totalApps}</div>
                        <div class="metric-label">Programas Ejecutados</div>
                    </div>
                </div>
            `;

            // Render Accesos (Timeline)
            html += `<div class="section-title">Línea de Tiempo de Conexión</div>`;
            if (userAccesos.length > 0) {
                html += `<div class="timeline-container">`;
                userAccesos.forEach(a => {
                    const status = a.Fin ? `<span style="color:var(--text-muted); font-size:0.8rem; margin-left:10px;">Cerró sesión a las ${formatTime(a.Fin)}</span>` : `<span class="status-active">EN LÍNEA</span>`;
                    html += `
                        <div class="timeline-item">
                            <div class="time-box">${formatTime(a.Inicio)}</div>
                            <div class="event-box">
                                <span class="event-title">Sesión ${a.Tipo}</span>
                                <span class="event-dur">Duración: ${formatDuration(a.Inicio, a.Fin)}</span>
                                ${status}
                            </div>
                        </div>
                    `;
                });
                html += `</div>`;
            } else {
                html += `<p style="color: var(--text-muted)">No hay inicios de sesión directos registrados.</p>`;
            }

            // Render Apps
            html += `<div class="section-title">Programas Utilizados</div>`;
            if (userApps.length > 0) {
                html += `<div class="app-grid">`;
                userApps.forEach(a => {
                    const status = a.Fin ? `Cerrado a las ${formatTime(a.Fin)}` : `En Uso Actualmente`;
                    html += `
                        <div class="app-card">
                            <div class="app-name">${a.Programa}</div>
                            <div class="app-time">Abrió: ${formatTime(a.Inicio)} &bull; Duró ${formatDuration(a.Inicio, a.Fin)}</div>
                            <div class="app-time" style="margin-top:5px; color:var(--secondary)">${status}</div>
                        </div>
                    `;
                });
                html += `</div>`;
            } else {
                html += `<p style="color: var(--text-muted)">No ejecutó programas clave este día.</p>`;
            }

            mainContent.innerHTML = html;
        }

        // Init
        renderUsers();
    </script>
</body>
</html>
'@

$HTML = $HTML.Replace('const DB = $JSONData;', "const DB = $JSONData;")

$RutaHTML = Join-Path -Path $DirectorioActual -ChildPath "Dashboard.html"
$HTML | Out-File -FilePath $RutaHTML -Encoding UTF8

Write-Host "Generando Dashboard..."
Start-Process $RutaHTML
