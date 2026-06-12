<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Consulta - Acceso Privado</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f4f7f6;
            color: #333;
            margin: 0;
            padding: 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 90vh;
        }
        .container {
            background-color: #fff;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
            max-width: 800px;
            width: 100%;
            box-sizing: border-box;
        }
        /* Estilos del Login */
        .login-card {
            max-width: 400px;
            margin: 0 auto;
            text-align: center;
        }
        .login-card h2 {
            color: #2c3e50;
            margin-bottom: 20px;
        }
        .form-group {
            margin-bottom: 15px;
            text-align: left;
        }
        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .form-group input {
            width: 100%;
            padding: 10px;
            border: 1px solid #ccc;
            border-radius: 4px;
            box-sizing: border-box;
            font-size: 16px;
        }
        .btn-login {
            width: 100%;
            padding: 12px;
            background-color: #3498db;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            cursor: pointer;
            font-weight: bold;
            margin-top: 10px;
        }
        .btn-login:hover {
            background-color: #2980b9;
        }
        .error-msg {
            color: #e74c3c;
            margin-top: 15px;
            font-weight: bold;
            display: none;
        }
        /* Estilos de la aplicación principal */
        h1 {
            color: #2c3e50;
            margin-top: 0;
            text-align: center;
        }
        .file-section, .search-section {
            margin-bottom: 20px;
            padding: 15px;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            background-color: #fafafa;
        }
        label {
            font-weight: bold;
            display: block;
            margin-bottom: 8px;
        }
        input[type="text"] {
            width: 100%;
            padding: 10px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 16px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }
        th {
            background-color: #27ae60;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        .no-results {
            text-align: center;
            color: #7f8c8d;
            margin-top: 20px;
        }
        .status-msg {
            font-style: italic;
            color: #2980b9;
            margin-top: 5px;
        }
        .header-app {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            border-bottom: 2px solid #eee;
            padding-bottom: 10px;
        }
        .btn-logout {
            padding: 8px 15px;
            background-color: #e74c3c;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .btn-logout:hover {
            background-color: #c0392b;
        }
    </style>
</head>
<body>

<div class="container login-card" id="login-screen">
    <h2>Iniciar Sesión</h2>
    <form id="login-form" onsubmit="event.preventDefault(); validarLogin();">
        <div class="form-group">
            <label for="username">Usuario:</label>
            <input type="text" id="username" required placeholder="Ingresa usuario">
        </div>
        <div class="form-group">
            <label for="password">Contraseña:</label>
            <input type="password" id="password" required placeholder="Ingresa contraseña">
        </div>
        <button type="submit" class="btn-login">Ingresar</button>
        <div id="error-message" class="error-msg">Usuario o contraseña incorrectos.</div>
    </form>
</div>

<div class="container" id="main-app" style="display: none;">
    <div class="header-app">
        <h1>Consultar Datos de Excel</h1>
        <button class="btn-logout" onclick="cerrarSesion()">Cerrar Sesión</button>
    </div>
    
    <div class="file-section">
        <label for="excel-file">1. Selecciona tu archivo de Excel (.xlsx o .xls):</label>
        <input type="file" id="excel-file" accept=".xlsx, .xls" />
        <div id="status" class="status-msg"></div>
    </div>

    <div class="search-section">
        <label for="search-input">2. Ingresa tu búsqueda (busca en cualquier columna):</label>
        <input type="text" id="search-input" placeholder="Escribe para filtrar datos..." disabled />
    </div>

    <div id="results-container">
        <table id="data-table" style="display: none;">
            <thead id="table-head"></thead>
            <tbody id="table-body"></tbody>
        </table>
        <div id="no-results-msg" class="no-results" style="display: none;">No se encontraron modificaciones o coincidencias.</div>
    </div>
</div>

<script>
    // --- LÓGICA DE CONTROL DE ACCESO ---
    function validarLogin() {
        const userIn = document.getElementById('username').value;
        const passIn = document.getElementById('password').value;
        const errorMsg = document.getElementById('error-message');

        // Validación estricta con las credenciales solicitadas
        if (userIn === "admin" && passIn === "1234") {
            errorMsg.style.display = "none";
            document.getElementById('login-screen').style.display = "none";
            document.getElementById('main-app').style.display = "block";
            
            // Limpiar el formulario de login por seguridad
            document.getElementById('username').value = "";
            document.getElementById('password').value = "";
        } else {
            errorMsg.style.display = "block";
        }
    }

    function cerrarSesion() {
        // Ocultar app, borrar datos cargados y mostrar login
        document.getElementById('main-app').style.display = "none";
        document.getElementById('login-screen').style.display = "block";
        
        // Resetear el buscador y el archivo cargado
        excelData = [];
        document.getElementById('excel-file').value = "";
        document.getElementById('search-input').value = "";
        document.getElementById('search-input').disabled = true;
        document.getElementById('status').innerText = "";
        document.getElementById('data-table').style.style = "none";
        document.getElementById('table-head').innerHTML = "";
        document.getElementById('table-body').innerHTML = "";
    }


    // --- LÓGICA DE LECTURA DE EXCEL (SHEETJS) ---
    let excelData = []; 

    document.getElementById('excel-file').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        document.getElementById('status').innerText = "Cargando y procesando archivo...";
        const reader = new FileReader();

        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            excelData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            document.getElementById('status').innerText = `Archivo cargado. Se encontraron ${excelData.length} filas.`;
            document.getElementById('search-input').disabled = false; 
            
            mostrarDatos(excelData);
        };

        reader.readAsArrayBuffer(file);
    });

    document.getElementById('search-input').addEventListener('input', function(e) {
        const query = e.target.value.toLowerCase().trim();
        
        if (query === "") {
            mostrarDatos(excelData);
            return;
        }

        const filteredData = excelData.filter(row => {
            return Object.values(row).some(value => 
                String(value).toLowerCase().includes(query)
            );
        });

        mostrarDatos(filteredData);
    });

    function mostrarDatos(data) {
        const table = document.getElementById('data-table');
        const thead = document.getElementById('table-head');
        const tbody = document.getElementById('table-body');
        const noResults = document.getElementById('no-results-msg');

        thead.innerHTML = "";
        tbody.innerHTML = "";

        if (data.length === 0) {
            table.style.display = "none";
            noResults.style.display = "block";
            return;
        }

        table.style.display = "table";
        noResults.style.display = "none";

        const columnas = Object.keys(excelData[0]);

        const headerRow = document.createElement('tr');
        columnas.forEach(col => {
            const th = document.createElement('th');
            th.innerText = col;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        data.forEach(row => {
            const tr = document.createElement('tr');
            columnas.forEach(col => {
                const td = document.createElement('td');
                td.innerText = row[col] !== undefined ? row[col] : "";
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }
</script>

</body>
</html>
