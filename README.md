<!DOCTYPE html>

<html lang="es">

<head>

<meta charset="UTF-8">

<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Productivity Operations Dashboard</title>

<script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>

<style> *{ margin:0; padding:0; box-sizing:border-box; font-family:"Segoe UI",Tahoma,Geneva,Verdana,sans-serif; } :root{ --navy:#002b5c; --blue:#0066ff; --blue-dark:#004fd6; --bg:#f4f7fb; --white:#ffffff; --text:#1f2937; --muted:#6b7280; --border:#dce2ea; --green:#198754; --yellow:#e6a700; --red:#dc3545; --purple:#6f42c1; } body{ background:var(--bg); color:var(--text); min-height:100vh; padding:20px; } /* ========================================================= HEADER ========================================================= */ .header{ background: linear-gradient( 135deg, #002b5c, #00529c ); color:white; padding:28px; border-radius:15px; margin-bottom:20px; box-shadow: 0 5px 18px rgba(0,43,92,.16); } .header h1{ font-size:29px; margin-bottom:7px; } .header p{ font-size:14px; opacity:.9; } /* ========================================================= CONTROLS ========================================================= */ .controls{ background:white; padding:18px; border-radius:12px; margin-bottom:15px; display:flex; gap:10px; flex-wrap:wrap; align-items:center; box-shadow: 0 2px 8px rgba(0,0,0,.06); } .file-area{ flex:1; min-width:280px; } input[type=file]{ width:100%; padding:10px; border:1px solid var(--border); border-radius:8px; background:white; cursor:pointer; } button{ border:0; border-radius:8px; padding:11px 17px; background:var(--blue); color:white; font-weight:600; cursor:pointer; transition:.2s; } button:hover:not(:disabled){ background:var(--blue-dark); transform:translateY(-1px); } button:disabled{ background:#aeb8c7; cursor:not-allowed; } .btn-secondary{ background:#56687d; } .btn-secondary:hover:not(:disabled){ background:#3f5062; } .btn-purple{ background:var(--purple); } .btn-purple:hover:not(:disabled){ background:#59349a; } /* ========================================================= STATUS ========================================================= */ #status{ min-height:23px; margin-bottom:8px; font-weight:600; } .ok-msg{ color:var(--green); } .err-msg{ color:var(--red); } .info-msg{ color:var(--blue); } #lastUpdate{ color:var(--muted); font-size:13px; margin-bottom:18px; } /* ========================================================= FILE INFO ========================================================= */ .file-info{ display:none; background:#edf6ff; border:1px solid #c8e2ff; padding:12px 15px; border-radius:8px; margin-bottom:20px; font-size:13px; } .file-info strong{ color:var(--navy); } /* ========================================================= SECTION ========================================================= */ .section{ margin-top:25px; margin-bottom:25px; } .section-header{ display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap; margin-bottom:12px; } .section h2{ color:var(--navy); font-size:20px; } /* ========================================================= KPI CARDS ========================================================= */ .cards{ display:grid; grid-template-columns: repeat( auto-fit, minmax(190px,1fr) ); gap:15px; margin-bottom:25px; } .card{ background:white; border-radius:12px; padding:20px; box-shadow: 0 2px 8px rgba(0,0,0,.06); border-left:5px solid var(--blue); } .card.green{ border-left-color:var(--green); } .card.yellow{ border-left-color:var(--yellow); } .card.red{ border-left-color:var(--red); } .card.purple{ border-left-color:var(--purple); } .card h3{ color:var(--muted); font-size:13px; margin-bottom:10px; } .card .value{ font-size:29px; font-weight:700; } /* ========================================================= TWO COLUMNS ========================================================= */ .grid-2{ display:grid; grid-template-columns: repeat( 2, minmax(0,1fr) ); gap:20px; } .panel{ background:white; border-radius:12px; padding:18px; box-shadow: 0 2px 8px rgba(0,0,0,.06); } .panel h3{ color:var(--navy); margin-bottom:15px; font-size:17px; } /* ========================================================= BAR CHART ========================================================= */ .chart{ display:flex; flex-direction:column; gap:9px; } .bar-row{ display:grid; grid-template-columns: 120px 1fr 55px; gap:8px; align-items:center; } .bar-label{ font-size:12px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; } .bar-container{ height:22px; background:#edf0f5; border-radius:5px; overflow:hidden; } .bar{ height:100%; background:var(--blue); border-radius:5px; min-width:2px; transition:width .3s; } .bar-value{ text-align:right; font-weight:600; font-size:12px; } /* ========================================================= TABLES ========================================================= */ .table-wrap{ max-height:500px; overflow:auto; border-radius:10px; background:white; box-shadow: 0 2px 8px rgba(0,0,0,.05); } table{ width:100%; border-collapse:collapse; background:white; } th{ position:sticky; top:0; z-index:2; background:var(--navy); color:white; padding:11px; text-align:left; white-space:nowrap; font-size:13px; } td{ padding:10px 11px; border-bottom: 1px solid #e8ebef; white-space:nowrap; font-size:13px; } tbody tr:hover{ background:#f4f8ff; } /* ========================================================= BADGES ========================================================= */ .badge{ display:inline-block; padding:4px 8px; border-radius:20px; font-size:11px; font-weight:600; background:#edf2f7; } .badge-green{ background:#dff5e7; color:#176b38; } .badge-yellow{ background:#fff1c7; color:#806000; } .badge-red{ background:#ffe1e1; color:#a51e1e; } .badge-blue{ background:#dcecff; color:#0759a5; } .badge-purple{ background:#eadfff; color:#59349a; } /* ========================================================= CASE REPEAT ========================================================= */ .repeat-card{ background:white; border-radius:12px; box-shadow: 0 2px 8px rgba(0,0,0,.06); margin-bottom:15px; overflow:hidden; border-left:5px solid var(--purple); } .repeat-header{ padding:15px; background:#faf8ff; display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap; } .repeat-header strong{ color:var(--navy); } .repeat-summary{ display:flex; gap:8px; flex-wrap:wrap; } .timeline{ padding:10px 15px 15px; } .timeline-item{ position:relative; padding: 10px 10px 10px 25px; border-left: 2px solid #d9caff; } .timeline-item:last-child{ border-left-color:transparent; } .timeline-dot{ position:absolute; left:-6px; top:13px; width:10px; height:10px; border-radius:50%; background:var(--purple); } .timeline-date{ color:var(--muted); font-size:11px; margin-bottom:3px; } .timeline-user{ font-weight:600; color:var(--navy); font-size:13px; } .timeline-detail{ color:#4b5563; font-size:12px; margin-top:3px; } /* ========================================================= SEARCH ========================================================= */ .search{ width:100%; padding:12px 14px; border: 1px solid var(--border); border-radius:8px; background:white; outline:none; margin-bottom:12px; } .search:focus{ border-color:var(--blue); box-shadow: 0 0 0 3px rgba(0,102,255,.1); } /* ========================================================= FILTERS ========================================================= */ .filters{ display:grid; grid-template-columns: repeat( auto-fit, minmax(180px,1fr) ); gap:10px; margin-bottom:15px; } select{ width:100%; padding:10px; border: 1px solid var(--border); border-radius:8px; background:white; } /* ========================================================= EMPTY ========================================================= */ .empty{ padding:30px; text-align:center; color:var(--muted); } /* ========================================================= FOOTER ========================================================= */ .footer{ text-align:center; color:#7b8490; font-size:12px; margin-top:30px; padding:15px; } /* ========================================================= RESPONSIVE ========================================================= */ @media(max-width:900px){ .grid-2{ grid-template-columns:1fr; } } @media(max-width:650px){ body{ padding:10px; } .header h1{ font-size:22px; } .controls{ flex-direction:column; align-items:stretch; } button{ width:100%; } .bar-row{ grid-template-columns: 85px 1fr 45px; } th, td{ font-size:11px; padding:8px; } } </style>

</head>

<body>

<!-- ========================================================= HEADER ========================================================= -->

<header class="header">

<h1>📊 Productivity Operations Dashboard</h1>

<p>
    Análisis de productividad, actividades,
    distribución operacional y reincidencia de casos.
</p>

</header>

<!-- ========================================================= CONTROLES ========================================================= -->

<div class="controls">

<div class="file-area">

    <input
        type="file"
        id="fileInput"
        accept=".csv,.xlsx,.xls">

</div>

<button
    id="btnAnalizar"
    onclick="processFile()">

    📂 Cargar y analizar

</button>

<button
    id="btnActualizar"
    class="btn-secondary"
    onclick="renderAll()"
    disabled>

    🔄 Actualizar

</button>

<button
    id="btnExportar"
    class="btn-purple"
    onclick="exportResults()"
    disabled>

    📥 Exportar análisis

</button>

</div>

<div id="status"></div>

<div id="lastUpdate"></div>

<div id="fileInfo" class="file-info"> </div>

<!-- ========================================================= KPI PRINCIPALES ========================================================= -->

<div class="cards">

<div class="card">

    <h3>Actividades</h3>

    <div
        id="kpiActividades"
        class="value">

        0

    </div>

</div>


<div class="card green">

    <h3>Casos únicos</h3>

    <div
        id="kpiCasos"
        class="value">

        0

    </div>

</div>


<div class="card blue">

    <h3>Agentes activos</h3>

    <div
        id="kpiAgentes"
        class="value">

        0

    </div>

</div>


<div class="card yellow">

    <h3>Promedio actividades / caso</h3>

    <div
        id="kpiPromedioCaso"
        class="value">

        0

    </div>

</div>


<div class="card red">

    <h3>Casos repetidos</h3>

    <div
        id="kpiRepetidos"
        class="value">

        0

    </div>

</div>


<div class="card purple">

    <h3>% casos repetidos</h3>

    <div
        id="kpiPctRepetidos"
        class="value">

        0%

    </div>

</div>


<div class="card">

    <h3>Modificaciones totales</h3>

    <div
        id="kpiModificaciones"
        class="value">

        0

    </div>

</div>


<div class="card green">

    <h3>% cerrados</h3>

    <div
        id="kpiCerrados"
        class="value">

        0%

    </div>

</div>

</div>

<!-- ========================================================= PRODUCTIVIDAD ========================================================= -->

<div class="section">

<div class="section-header">

    <h2>👥 Productividad por agente</h2>

</div>


<div class="table-wrap">

    <table id="agentTable">

        <thead>

            <tr>

                <th>Agente</th>

                <th>Actividades</th>

                <th>Casos únicos</th>

                <th>Casos repetidos</th>

                <th>Modificaciones</th>

                <th>% cerrados</th>

                <th>Actividad principal</th>

            </tr>

        </thead>

        <tbody></tbody>

    </table>

</div>

</div>

<!-- ========================================================= DISTRIBUCIONES ========================================================= -->

<div class="section">

<div class="grid-2">


    <div class="panel">

        <h3>🛠️ Tipo de gestión</h3>

        <div
            id="gestionChart"
            class="chart">

        </div>

    </div>


    <div class="panel">

        <h3>📁 Tipo de caso</h3>

        <div
            id="tipoCasoChart"
            class="chart">

        </div>

    </div>


</div>

</div>

<!-- ========================================================= ESTADOS + HORAS ========================================================= -->

<div class="section">

<div class="grid-2">


    <div class="panel">

        <h3>📌 Estado del caso</h3>

        <div
            id="estadoChart"
            class="chart">

        </div>

    </div>


    <div class="panel">

        <h3>🕐 Actividades por hora</h3>

        <div
            id="horaChart"
            class="chart">

        </div>

    </div>


</div>

</div>

<!-- ========================================================= CASOS REPETIDOS ========================================================= -->

<div class="section">

<div class="section-header">

    <h2>🔁 Casos con múltiples modificaciones</h2>

</div>


<div class="panel">

    <div class="filters">

        <select id="repeatAgentFilter">

            <option value="">
                Todos los agentes
            </option>

        </select>


        <select id="repeatMinFilter">

            <option value="2">
                2+ modificaciones
            </option>

            <option value="3">
                3+ modificaciones
            </option>

            <option value="4">
                4+ modificaciones
            </option>

            <option value="5">
                5+ modificaciones
            </option>

        </select>

    </div>


    <div
        id="repeatSummary"
        class="empty">

        Carga una base para visualizar
        los casos repetidos.

    </div>

</div>

</div>

<!-- ========================================================= QUIÉN MODIFICA CASOS REPETIDOS ========================================================= -->

<div class="section">

<div class="section-header">

    <h2>👤 Intervención en casos repetidos</h2>

</div>


<div class="table-wrap">

    <table id="repeatAgentTable">

        <thead>

            <tr>

                <th>Agente</th>

                <th>Casos repetidos</th>

                <th>Modificaciones</th>

                <th>Tratamiento</th>

                <th>Actualización</th>

                <th>No requiere respuesta</th>

                <th>Otros</th>

            </tr>

        </thead>

        <tbody></tbody>

    </table>

</div>

</div>

<!-- ========================================================= DISTRIBUCIÓN DE REPETICIONES ========================================================= -->

<div class="section">

<div class="grid-2">


    <div class="panel">

        <h3>🔢 Distribución de intervenciones</h3>

        <div
            id="repeatDistribution"
            class="chart">

        </div>

    </div>


    <div class="panel">

        <h3>⚠️ Casos con más modificaciones</h3>

        <div class="table-wrap">

            <table id="topRepeatedTable">

                <thead>

                    <tr>

                        <th>Caso</th>

                        <th>Veces</th>

                        <th>Agentes</th>

                    </tr>

                </thead>

                <tbody></tbody>

            </table>

        </div>

    </div>


</div>

</div>

<!-- ========================================================= DETALLE GENERAL ========================================================= -->

<div class="section">

<div class="section-header">

    <h2>📋 Detalle de actividades</h2>

</div>


<input
    id="searchInput"
    class="search"
    type="search"
    placeholder="Buscar caso, agente, gestión, estado o tipo de caso...">


<div class="filters">

    <select id="agentFilter">

        <option value="">
            Todos los agentes
        </option>

    </select>


    <select id="gestionFilter">

        <option value="">
            Todos los tipos de gestión
        </option>

    </select>


    <select id="statusFilter">

        <option value="">
            Todos los estados
        </option>

    </select>


    <select id="typeFilter">

        <option value="">
            Todos los tipos de caso
        </option>

    </select>

</div>


<div class="table-wrap">

    <table id="activityTable">

        <thead>

            <tr>

                <th>Caso</th>

                <th>Tipo de caso</th>

                <th>Estado</th>

                <th>Fecha</th>

                <th>Hora</th>

                <th>Agente</th>

                <th>Correo</th>

                <th>Tipo de gestión</th>

                <th>Repeticiones</th>

            </tr>

        </thead>

        <tbody></tbody>

    </table>

</div>

</div>

<footer class="footer">

Productivity Operations Dashboard ·
Procesamiento local en el navegador ·
Compatible con GitHub Pages

</footer>

<script> /* ========================================================= ESTADO GLOBAL ========================================================= */ let actividades = []; let casosMap = {}; let autoRefreshTimer = null; /* ========================================================= NORMALIZAR TEXTO ========================================================= */ function norm(value){ if( value === null || value === undefined ){ return ""; } return String(value) .normalize("NFD") .replace(/[\u0300-\u036f]/g,"") .toLowerCase() .trim(); } /* ========================================================= ESCAPAR HTML ========================================================= */ function escapeHtml(value){ return String(value ?? "") .replace(/&/g,"&amp;") .replace(/</g,"&lt;") .replace(/>/g,"&gt;") .replace(/"/g,"&quot;") .replace(/'/g,"&#039;"); } /* ========================================================= DETECTAR COLUMNAS ========================================================= */ function detectarColumnas(headers){ const cols = {}; headers.forEach( (header,index)=>{ const h = norm(header); if( h === "numerodecaso" || h === "numero de caso" ){ cols.caso = index; } else if( h === "numerocaso" ){ /* * Se conserva como alternativa, * pero la base proporcionada utiliza * NumerodeCaso como identificador real. */ if(cols.caso === undefined){ cols.caso = index; } } else if( h === "tipocaso" || h === "tipo caso" ){ cols.tipoCaso = index; } else if( h === "estadocaso" || h === "estado caso" ){ cols.estado = index; } else if( h === "fecha" ){ cols.fecha = index; } else if( h === "hora" ){ cols.hora = index; } else if( h === "correo" || h === "email" ){ cols.correo = index; } else if( h === "usuario" || h === "agente" ){ cols.usuario = index; } else if( h === "fechahoraregistro" || h === "fecha hora registro" ){ cols.fechaHoraRegistro = index; } else if( h === "tipogestion" || h === "tipo gestion" ){ cols.tipoGestion = index; } } ); return cols; } /* ========================================================= LEER ARCHIVO ========================================================= */ function processFile(){ const file = document.getElementById("fileInput") .files[0]; if(!file){ setStatus( "Selecciona un archivo primero.", "err-msg" ); return; } setStatus( "⏳ Procesando archivo...", "info-msg" ); document.getElementById( "btnAnalizar" ).disabled = true; const reader = new FileReader(); reader.onload = function(event){ try{ const data = new Uint8Array( event.target.result ); let workbook; /* * CSV */ if( file.name .toLowerCase() .endsWith(".csv") ){ const text = new TextDecoder( "utf-8" ).decode(data); workbook = XLSX.read( text, { type:"string" } ); } /* * XLS / XLSX */ else{ workbook = XLSX.read( data, { type:"array", cellDates:false } ); } if( !workbook.SheetNames.length ){ throw new Error( "El archivo no contiene hojas." ); } const sheet = workbook.Sheets[ workbook.SheetNames[0] ]; const matrix = XLSX.utils.sheet_to_json( sheet, { header:1, defval:"", raw:true } ); if(matrix.length < 2){ throw new Error( "El archivo no contiene suficientes datos." ); } analizarBase( matrix, file ); } catch(error){ console.error(error); setStatus( "❌ Error procesando el archivo: " + error.message, "err-msg" ); } finally{ document.getElementById( "btnAnalizar" ).disabled = false; } }; reader.onerror = function(){ setStatus( "❌ No fue posible leer el archivo.", "err-msg" ); document.getElementById( "btnAnalizar" ).disabled = false; }; reader.readAsArrayBuffer(file); } /* ========================================================= ANALIZAR BASE ========================================================= */ function analizarBase(matrix,file){ const headers = matrix[0]; const cols = detectarColumnas( headers ); if(cols.caso === undefined){ throw new Error( "No se encontró la columna NumerodeCaso." ); } actividades = []; for( let i=1; i<matrix.length; i++ ){ const row = matrix[i] || []; const caso = String( row[cols.caso] ?? "" ).trim(); if(!caso){ continue; } const actividad = { index:i, caso, tipoCaso: cols.tipoCaso !== undefined ? String( row[cols.tipoCaso] ?? "" ).trim() : "", estado: cols.estado !== undefined ? String( row[cols.estado] ?? "" ).trim() : "", fecha: cols.fecha !== undefined ? String( row[cols.fecha] ?? "" ).trim() : "", hora: cols.hora !== undefined ? String( row[cols.hora] ?? "" ).trim() : "", correo: cols.correo !== undefined ? String( row[cols.correo] ?? "" ).trim() : "", usuario: cols.usuario !== undefined ? String( row[cols.usuario] ?? "" ).trim() : "Sin usuario", fechaHoraRegistro: cols.fechaHoraRegistro !== undefined ? String( row[cols.fechaHoraRegistro] ?? "" ).trim() : "", tipoGestion: cols.tipoGestion !== undefined ? String( row[cols.tipoGestion] ?? "" ).trim() : "Sin especificar" }; actividades.push( actividad ); } if(!actividades.length){ throw new Error( "No se encontraron registros válidos." ); } construirCasos(); llenarFiltros(); renderAll(); document.getElementById( "btnActualizar" ).disabled = false; document.getElementById( "btnExportar" ).disabled = false; const repetidos = Object.values(casosMap) .filter( c => c.actividades.length > 1 ) .length; setStatus( "✅ Base procesada: " + actividades.length + " actividades, " + Object.keys(casosMap).length + " casos únicos y " + repetidos + " casos repetidos.", "ok-msg" ); mostrarInfoArchivo( file ); } /* ========================================================= CONSTRUIR MAPA DE CASOS ========================================================= */ function construirCasos(){ casosMap = {}; actividades.forEach( actividad=>{ if( !casosMap[ actividad.caso ] ){ casosMap[ actividad.caso ] = { caso: actividad.caso, actividades:[] }; } casosMap[ actividad.caso ] .actividades .push( actividad ); } ); Object.values(casosMap) .forEach( caso=>{ caso.actividades.sort( compararFechaActividad ); } ); } /* ========================================================= FECHA/HORA PARA ORDENAMIENTO ========================================================= */ function parseFechaActividad(a){ /* * La base contiene Fecha y Hora separadas. * * Se intenta construir una fecha local. */ const texto = ( a.fecha + " " + a.hora ).trim(); const d = new Date(texto); if(!isNaN(d.getTime())){ return d.getTime(); } /* * Intento DD/MM/YYYY. */ const match = a.fecha.match( /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/ ); if(match){ const dia = Number(match[1]); const mes = Number(match[2])-1; const año = Number(match[3]); const horaMatch = a.hora.match( /(\d{1,2}):(\d{2})(?::(\d{2}))?/ ); const hora = horaMatch ? Number(horaMatch[1]) : 0; const minuto = horaMatch ? Number(horaMatch[2]) : 0; const segundo = horaMatch && horaMatch[3] ? Number(horaMatch[3]) : 0; return new Date( año, mes, dia, hora, minuto, segundo ).getTime(); } return 0; } function compararFechaActividad(a,b){ return ( parseFechaActividad(a) - parseFechaActividad(b) ); } /* ========================================================= RENDER PRINCIPAL ========================================================= */ function renderAll(){ if(!actividades.length){ return; } renderKPIs(); renderAgents(); renderCharts(); renderRepeatCases(); renderRepeatAgents(); renderRepeatDistribution(); renderTopRepeated(); renderActivityTable(); document.getElementById( "lastUpdate" ).textContent = "Última actualización: " + new Date().toLocaleString( "es-CO", { timeZone: "America/Bogota" } ); } /* ========================================================= KPIs ========================================================= */ function renderKPIs(){ const casos = Object.values(casosMap); const repetidos = casos.filter( c => c.actividades.length > 1 ); const cerrados = actividades.filter( a => norm(a.estado) .includes("cerrado") ); const promedio = casos.length ? actividades.length / casos.length : 0; const pctRepetidos = casos.length ? ( repetidos.length / casos.length * 100 ) .toFixed(1) : "0"; const pctCerrados = actividades.length ? ( cerrados.length / actividades.length * 100 ) .toFixed(1) : "0"; const agentes = new Set( actividades.map( a => a.usuario ) ); document.getElementById( "kpiActividades" ).textContent = actividades.length; document.getElementById( "kpiCasos" ).textContent = casos.length; document.getElementById( "kpiAgentes" ).textContent = agentes.size; document.getElementById( "kpiPromedioCaso" ).textContent = promedio.toFixed(2); document.getElementById( "kpiRepetidos" ).textContent = repetidos.length; document.getElementById( "kpiPctRepetidos" ).textContent = pctRepetidos + "%"; document.getElementById( "kpiModificaciones" ).textContent = actividades.length; document.getElementById( "kpiCerrados" ).textContent = pctCerrados + "%"; } /* ========================================================= AGENTES ========================================================= */ function renderAgents(){ const agentes = {}; actividades.forEach( a=>{ if(!agentes[a.usuario]){ agentes[a.usuario] = { actividades:0, casos:new Set(), repetidos:new Set(), modificaciones:0, cerrados:0, gestiones:{} }; } const x = agentes[a.usuario]; x.actividades++; x.casos.add( a.caso ); x.modificaciones++; if( norm(a.estado) .includes("cerrado") ){ x.cerrados++; } if( !x.gestiones[ a.tipoGestion ] ){ x.gestiones[ a.tipoGestion ] = 0; } x.gestiones[ a.tipoGestion ]++; } ); Object.values(casosMap) .forEach( caso=>{ if( caso.actividades.length <= 1 ){ return; } const agentesCaso = new Set( caso.actividades.map( a => a.usuario ) ); agentesCaso.forEach( agente=>{ if( agentes[agente] ){ agentes[ agente ] .repetidos .add( caso.caso ); } } ); } ); const tbody = document.querySelector( "#agentTable tbody" ); tbody.innerHTML = ""; Object.entries(agentes) .sort( (a,b)=> b[1].actividades - a[1].actividades ) .forEach( ([nombre,x])=>{ const porcentaje = x.actividades ? ( x.cerrados / x.actividades * 100 ).toFixed(1) : "0"; const principal = Object.entries( x.gestiones ) .sort( (a,b)=> b[1]-a[1] )[0]; const tr = document.createElement( "tr" ); tr.innerHTML = ` <td> <strong> ${escapeHtml(nombre)} </strong> </td> <td> ${x.actividades} </td> <td> ${x.casos.size} </td> <td> ${x.repetidos.size} </td> <td> ${x.modificaciones} </td> <td> ${porcentaje}% </td> <td> ${ principal ? escapeHtml( principal[0] ) : "-" } </td> `; tbody.appendChild(tr); } ); } /* ========================================================= CHARTS ========================================================= */ function renderCharts(){ renderSimpleChart( "gestionChart", contar( actividades, a=>a.tipoGestion ) ); renderSimpleChart( "tipoCasoChart", contar( actividades, a=>a.tipoCaso ) ); renderSimpleChart( "estadoChart", contar( actividades, a=>a.estado ) ); const horas = {}; actividades.forEach( a=>{ let h = extraerHora( a.hora ); if(h === null){ const d = parseFechaActividad(a); if(d){ h = new Date(d) .getHours(); } } if(h !== null){ const key = String(h) .padStart(2,"0") + ":00"; horas[key] = (horas[key] || 0) + 1; } } ); renderSimpleChart( "horaChart", horas, true ); } function contar(array,fn){ const result = {}; array.forEach( item=>{ const key = fn(item) || "Sin especificar"; result[key] = (result[key] || 0) + 1; } ); return result; } function renderSimpleChart( containerId, data, sortNumeric=false ){ const container = document.getElementById( containerId ); container.innerHTML = ""; const entries = Object.entries(data); if(!entries.length){ container.innerHTML = `<div class="empty"> Sin datos </div>`; return; } entries.sort( (a,b)=> b[1]-a[1] ); const max = Math.max( ...entries.map( x=>x[1] ) ); entries.forEach( ([label,value])=>{ const percentage = max ? value/max*100 : 0; const row = document.createElement( "div" ); row.className = "bar-row"; row.innerHTML = ` <div class="bar-label" title="${escapeHtml(label)}"> ${escapeHtml(label)} </div> <div class="bar-container"> <div class="bar" style="width:${percentage}%"> </div> </div> <div class="bar-value"> ${value} </div> `; container.appendChild(row); } ); } /* ========================================================= HORA ========================================================= */ function extraerHora(value){ const match = String(value || "") .match( /(\d{1,2}):(\d{2})/ ); if(!match){ return null; } const hora = Number(match[1]); if( hora < 0 || hora > 23 ){ return null; } return hora; } /* ========================================================= CASOS REPETIDOS ========================================================= */ function getRepeatedCases(){ return Object.values( casosMap ) .filter( c => c.actividades.length > 1 ); } function renderRepeatCases(){ const container = document.getElementById( "repeatSummary" ); const min = Number( document.getElementById( "repeatMinFilter" ).value ); const agente = document.getElementById( "repeatAgentFilter" ).value; let casos = getRepeatedCases() .filter( c => c.actividades.length >= min ); if(agente){ casos = casos.filter( c => c.actividades.some( a => a.usuario === agente ) ); } casos.sort( (a,b)=> b.actividades.length - a.actividades.length ); if(!casos.length){ container.innerHTML = ` <div class="empty"> No hay casos que cumplan los filtros seleccionados. </div> `; return; } container.innerHTML = ""; casos.forEach( caso=>{ const card = document.createElement( "div" ); card.className = "repeat-card"; const agentes = [ ...new Set( caso.actividades.map( a => a.usuario ) ) ]; const gestiones = [ ...new Set( caso.actividades.map( a => a.tipoGestion ) ) ]; const header = document.createElement( "div" ); header.className = "repeat-header"; header.innerHTML = ` <div> <strong> Caso #${escapeHtml(caso.caso)} </strong> </div> <div class="repeat-summary"> <span class="badge badge-purple"> ${caso.actividades.length} modificaciones </span> <span class="badge badge-blue"> ${agentes.length} agente(s) </span> <span class="badge"> ${gestiones.length} gestión(es) </span> </div> `; card.appendChild(header); const timeline = document.createElement( "div" ); timeline.className = "timeline"; caso.actividades.forEach( actividad=>{ const item = document.createElement( "div" ); item.className = "timeline-item"; const fecha = ( actividad.fecha + " " + actividad.hora ).trim(); item.innerHTML = ` <div class="timeline-dot"> </div> <div class="timeline-date"> ${escapeHtml(fecha)} </div> <div class="timeline-user"> ${escapeHtml( actividad.usuario )} </div> <div class="timeline-detail"> Gestión: <strong> ${escapeHtml( actividad.tipoGestion )} </strong> &nbsp; · &nbsp; Estado: <strong> ${escapeHtml( actividad.estado )} </strong> ${ actividad.tipoCaso ? ` &nbsp; · &nbsp; Tipo: ${escapeHtml( actividad.tipoCaso )} ` : "" } </div> `; timeline.appendChild( item ); } ); card.appendChild( timeline ); container.appendChild( card ); } ); } /* ========================================================= AGENTES EN CASOS REPETIDOS ========================================================= */ function renderRepeatAgents(){ const data = {}; getRepeatedCases() .forEach( caso=>{ const agentes = new Set( caso.actividades.map( a => a.usuario ) ); agentes.forEach( agente=>{ if(!data[agente]){ data[agente] = { casos:new Set(), modificaciones:0, tratamiento:0, actualizacion:0, noRespuesta:0, otros:0 }; } data[agente] .casos .add( caso.caso ); } ); caso.actividades.forEach( actividad=>{ const agente = actividad.usuario; if(!data[agente]){ data[agente] = { casos:new Set(), modificaciones:0, tratamiento:0, actualizacion:0, noRespuesta:0, otros:0 }; } data[agente] .modificaciones++; const gestion = norm( actividad.tipoGestion ); if( gestion.includes( "tratamiento" ) ){ data[agente] .tratamiento++; } else if( gestion.includes( "actualizacion" ) ){ data[agente] .actualizacion++; } else if( gestion.includes( "no requiere respuesta" ) || gestion.includes( "no requiere" ) ){ data[agente] .noRespuesta++; } else{ data[agente] .otros++; } } ); } ); const tbody = document.querySelector( "#repeatAgentTable tbody" ); tbody.innerHTML = ""; Object.entries(data) .sort( (a,b)=> b[1].modificaciones - a[1].modificaciones ) .forEach( ([agente,x])=>{ const tr = document.createElement( "tr" ); tr.innerHTML = ` <td> <strong> ${escapeHtml(agente)} </strong> </td> <td> ${x.casos.size} </td> <td> ${x.modificaciones} </td> <td> ${x.tratamiento} </td> <td> ${x.actualizacion} </td> <td> ${x.noRespuesta} </td> <td> ${x.otros} </td> `; tbody.appendChild(tr); } ); } /* ========================================================= DISTRIBUCIÓN DE REPETICIONES ========================================================= */ function renderRepeatDistribution(){ const distribucion = {}; getRepeatedCases() .forEach( caso=>{ let categoria; if( caso.actividades.length >= 4 ){ categoria = "4+"; } else{ categoria = String( caso.actividades.length ); } distribucion[categoria] = ( distribucion[categoria] || 0 ) + 1; } ); const orden = { "2":0, "3":1, "4+":2 }; const container = document.getElementById( "repeatDistribution" ); container.innerHTML = ""; Object.entries( distribucion ) .sort( (a,b)=> (orden[a[0]] ?? 99) - (orden[b[0]] ?? 99) ) .forEach( ([label,value])=>{ const row = document.createElement( "div" ); row.className = "bar-row"; const max = Math.max( ...Object.values( distribucion ) ); const width = value/max*100; row.innerHTML = ` <div class="bar-label"> ${label} ${ label === "4+" ? " modificaciones" : " modificaciones" } </div> <div class="bar-container"> <div class="bar" style="width:${width}%"> </div> </div> <div class="bar-value"> ${value} </div> `; container.appendChild(row); } ); } /* ========================================================= TOP CASOS REPETIDOS ========================================================= */ function renderTopRepeated(){ const tbody = document.querySelector( "#topRepeatedTable tbody" ); tbody.innerHTML = ""; getRepeatedCases() .sort( (a,b)=> b.actividades.length - a.actividades.length ) .slice(0,15) .forEach( caso=>{ const agentes = new Set( caso.actividades.map( a => a.usuario ) ); const tr = document.createElement( "tr" ); tr.innerHTML = ` <td> <strong> ${escapeHtml( caso.caso )} </strong> </td> <td> <span class="badge badge-purple"> ${caso.actividades.length} </span> </td> <td> ${agentes.size} </td> `; tbody.appendChild(tr); } ); } /* ========================================================= TABLA DE ACTIVIDADES ========================================================= */ function renderActivityTable(){ const search = norm( document.getElementById( "searchInput" ).value ); const agent = document.getElementById( "agentFilter" ).value; const gestion = document.getElementById( "gestionFilter" ).value; const status = document.getElementById( "statusFilter" ).value; const type = document.getElementById( "typeFilter" ).value; const tbody = document.querySelector( "#activityTable tbody" ); tbody.innerHTML = ""; const datos = actividades .filter( a=>{ const texto = norm( [ a.caso, a.tipoCaso, a.estado, a.fecha, a.hora, a.usuario, a.correo, a.tipoGestion ].join(" ") ); if( search && !texto.includes(search) ){ return false; } if( agent && a.usuario !== agent ){ return false; } if( gestion && a.tipoGestion !== gestion ){ return false; } if( status && a.estado !== status ){ return false; } if( type && a.tipoCaso !== type ){ return false; } return true; } ); datos .sort( compararFechaActividad ) .reverse(); datos.forEach( actividad=>{ const cantidad = casosMap[ actividad.caso ] ? casosMap[ actividad.caso ] .actividades .length : 1; const tr = document.createElement( "tr" ); tr.innerHTML = ` <td> <strong> ${escapeHtml( actividad.caso )} </strong> </td> <td> ${escapeHtml( actividad.tipoCaso )} </td> <td> ${escapeHtml( actividad.estado )} </td> <td> ${escapeHtml( actividad.fecha )} </td> <td> ${escapeHtml( actividad.hora )} </td> <td> ${escapeHtml( actividad.usuario )} </td> <td> ${escapeHtml( actividad.correo )} </td> <td> ${escapeHtml( actividad.tipoGestion )} </td> <td> ${ cantidad > 1 ? ` <span class="badge badge-purple"> ${cantidad} </span> ` : ` <span class="badge"> 1 </span> ` } </td> `; tbody.appendChild(tr); } ); if(!datos.length){ tbody.innerHTML = ` <tr> <td colspan="9" class="empty"> No se encontraron actividades. </td> </tr> `; } } /* ========================================================= FILTROS ========================================================= */ function llenarFiltros(){ llenarSelect( "agentFilter", unique( actividades.map( a=>a.usuario ) ) ); llenarSelect( "gestionFilter", unique( actividades.map( a=>a.tipoGestion ) ) ); llenarSelect( "statusFilter", unique( actividades.map( a=>a.estado ) ) ); llenarSelect( "typeFilter", unique( actividades.map( a=>a.tipoCaso ) ) ); llenarSelect( "repeatAgentFilter", unique( actividades.map( a=>a.usuario ) ) ); } function unique(array){ return [ ...new Set( array.filter( Boolean ) ) ].sort( (a,b)=> a.localeCompare( b, "es" ) ); } function llenarSelect( id, values ){ const select = document.getElementById(id); while( select.options.length > 1 ){ select.remove( 1 ); } values.forEach( value=>{ const option = document.createElement( "option" ); option.value = value; option.textContent = value; select.appendChild( option ); } ); } /* ========================================================= EXPORTACIÓN ========================================================= */ function exportResults(){ if(!actividades.length){ return; } const wb = XLSX.utils.book_new(); /* * HOJA 1: * ACTIVIDADES */ const actividadesExport = actividades.map( a=>({ "Caso": a.caso, "Tipo de caso": a.tipoCaso, "Estado": a.estado, "Fecha": a.fecha, "Hora": a.hora, "Agente": a.usuario, "Correo": a.correo, "Fecha/Hora Registro": a.fechaHoraRegistro, "Tipo de Gestión": a.tipoGestion, "Número de modificaciones": casosMap[ a.caso ] .actividades .length }) ); const wsActividades = XLSX.utils.json_to_sheet( actividadesExport ); XLSX.utils.book_append_sheet( wb, wsActividades, "Actividades" ); /* * HOJA 2: * RESUMEN DE CASOS */ const casosExport = Object.values(casosMap) .map( caso=>{ const agentes = [ ...new Set( caso.activities ? caso.activities.map( a=>a.usuario ) : caso.actividades.map( a=>a.usuario ) ) ]; return { "Caso": caso.caso, "Número de actividades": caso.actividades.length, "Número de agentes": agentes.length, "Agentes": agentes.join( " | " ), "Primera actividad": caso.actividades.length ? ( caso.actividades[0] .fecha + " " + caso.actividades[0] .hora ) : "", "Última actividad": caso.actividades.length ? ( caso.actividades[ caso.actividades.length-1 ].fecha + " " + caso.actividades[ caso.actividades.length-1 ].hora ) : "" }; } ); const wsCasos = XLSX.utils.json_to_sheet( casosExport ); XLSX.utils.book_append_sheet( wb, wsCasos, "Casos" ); /* * HOJA 3: * INTERVENCIONES */ const intervenciones = []; getRepeatedCases() .forEach( caso=>{ caso.actividades.forEach( actividad=>{ intervenciones.push({ "Caso": caso.caso, "Fecha": actividad.fecha, "Hora": actividad.hora, "Agente": actividad.usuario, "Correo": actividad.correo, "Tipo de gestión": actividad.tipoGestion, "Estado": actividad.estado, "Tipo de caso": actividad.tipoCaso, "Total modificaciones caso": caso.actividades.length }); } ); } ); const wsIntervenciones = XLSX.utils.json_to_sheet( intervenciones ); XLSX.utils.book_append_sheet( wb, wsIntervenciones, "Casos Repetidos" ); /* * HOJA 4: * PRODUCTIVIDAD */ const productividad = {}; actividades.forEach( a=>{ if( !productividad[a.usuario] ){ productividad[a.usuario] = { actividades:0, casos:new Set(), cerrados:0 }; } productividad[a.usuario] .actividades++; productividad[a.usuario] .casos .add( a.caso ); if( norm(a.estado) .includes("cerrado") ){ productividad[a.usuario] .cerrados++; } } ); const productividadExport = Object.entries( productividad ) .map( ([agente,x])=>({ "Agente": agente, "Actividades": x.actividades, "Casos únicos": x.casos.size, "Casos cerrados": x.cerrados, "% cerrados": x.actividades ? Number( ( x.cerrados / x.actividades * 100 ).toFixed(1) ) : 0 }) ); const wsProductividad = XLSX.utils.json_to_sheet( productividadExport ); XLSX.utils.book_append_sheet( wb, wsProductividad, "Productividad" ); XLSX.writeFile( wb, "Productivity_Operations_Analysis.xlsx" ); } /* ========================================================= INFO ARCHIVO ========================================================= */ function mostrarInfoArchivo(file){ const info = document.getElementById( "fileInfo" ); const mb = ( file.size / 1024 / 1024 ).toFixed(2); info.innerHTML = ` <strong>Archivo:</strong> ${escapeHtml(file.name)} &nbsp; | &nbsp; <strong>Tamaño:</strong> ${mb} MB &nbsp; | &nbsp; <strong>Registros:</strong> ${actividades.length} `; info.style.display = "block"; } /* ========================================================= STATUS ========================================================= */ function setStatus( message, className="" ){ const el = document.getElementById( "status" ); el.textContent = message; el.className = className; } /* ========================================================= EVENTOS ========================================================= */ document .getElementById( "fileInput" ) .addEventListener( "change", function(){ if(this.files[0]){ processFile(); } } ); document .getElementById( "searchInput" ) .addEventListener( "input", renderActivityTable ); document .getElementById( "agentFilter" ) .addEventListener( "change", renderActivityTable ); document .getElementById( "gestionFilter" ) .addEventListener( "change", renderActivityTable ); document .getElementById( "statusFilter" ) .addEventListener( "change", renderActivityTable ); document .getElementById( "typeFilter" ) .addEventListener( "change", renderActivityTable ); document .getElementById( "repeatAgentFilter" ) .addEventListener( "change", renderRepeatCases ); document .getElementById( "repeatMinFilter" ) .addEventListener( "change", renderRepeatCases ); /* ========================================================= FIN ========================================================= */ </script>

</body>

</html>
