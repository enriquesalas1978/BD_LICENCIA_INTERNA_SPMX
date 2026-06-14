<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Verificación SSOMA - Licencias Internas</title>
    <style>
        :root {
            --primary-color: #003366;
            --secondary-color: #f4f4f4;
            --accent-color: #ff9900;
            --text-color: #333333;
            --border-color: #cccccc;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: var(--secondary-color);
            color: var(--text-color);
            margin: 0;
            padding: 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .container {
            background-color: #ffffff;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            width: 100%;
            max-width: 600px;
        }
        h2 {
            color: var(--primary-color);
            text-align: center;
            border-bottom: 2px solid var(--accent-color);
            padding-bottom: 10px;
        }
        .search-section {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
        }
        input[type="text"] {
            flex: 1;
            padding: 10px;
            border: 1px solid var(--border-color);
            border-radius: 4px;
            font-size: 16px;
        }
        button {
            background-color: var(--primary-color);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
            transition: background-color 0.3s;
        }
        button:hover {
            background-color: #002244;
        }
        .ficha-trabajador {
            display: none; /* Oculto por defecto */
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 20px;
            background-color: #fafafa;
        }
        .ficha-header {
            background-color: var(--primary-color);
            color: white;
            padding: 10px;
            text-align: center;
            border-radius: 4px 4px 0 0;
            margin: -20px -20px 20px -20px;
        }
        .data-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
        }
        .data-item {
            display: flex;
            flex-direction: column;
        }
        .data-label {
            font-size: 12px;
            color: #666;
            font-weight: bold;
            text-transform: uppercase;
        }
        .data-value {
            font-size: 15px;
            font-weight: 500;
            color: var(--text-color);
            padding: 5px 0;
            border-bottom: 1px solid #eee;
        }
        .status-vigente {
            color: green;
            font-weight: bold;
        }
        .status-vencido {
            color: red;
            font-weight: bold;
        }
        #error-message {
            color: red;
            text-align: center;
            display: none;
            font-weight: bold;
            margin-top: 15px;
        }
    </style>
</head>
<body>

    <div class="container">
        <h2>Verificación de Licencia de Equipo Móvil</h2>
        
        <div class="search-section">
            <input type="text" id="dni-input" placeholder="Ingrese el DNI del trabajador" maxlength="8" onkeypress="return isNumberKey(event)">
            <button onclick="buscarLicencia()">Consultar</button>
        </div>

        <div id="error-message">DNI no registrado en la base de datos o sin autorización vigente.</div>

        <div class="ficha-trabajador" id="ficha-result">
            <div class="ficha-header">
                <h3 style="margin:0;">Ficha de Operador Autorizado</h3>
            </div>
            <div class="data-grid">
                <div class="data-item">
                    <span class="data-label">Licencia N° (Interna)</span>
                    <span class="data-value" id="res-lic-interna">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">DNI</span>
                    <span class="data-value" id="res-dni">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Apellidos</span>
                    <span class="data-value" id="res-apellidos">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Nombres</span>
                    <span class="data-value" id="res-nombres">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Cargo</span>
                    <span class="data-value" id="res-cargo">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Equipo Autorizado</span>
                    <span class="data-value" id="res-equipo">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Licencia Conducir MTC</span>
                    <span class="data-value" id="res-lic-mtc">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Clase / Categoría</span>
                    <span class="data-value" id="res-clase">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">F. Emisión MTC</span>
                    <span class="data-value" id="res-emision">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">F. Caducidad MTC</span>
                    <span class="data-value" id="res-caducidad">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Situación Licencia MTC</span>
                    <span class="data-value" id="res-situacion">-</span>
                </div>
                <div class="data-item">
                    <span class="data-label">Restricciones</span>
                    <span class="data-value" id="res-restricciones">-</span>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Restricción para ingresar solo números en el input del DNI
        function isNumberKey(evt) {
            var charCode = (evt.which) ? evt.which : evt.keyCode
            if (charCode > 31 && (charCode < 48 || charCode > 57))
                return false;
            return true;
        }

        // Simulación de la Matriz de Excel alojada en OneDrive convertida a formato JSON
        // Esta estructura deberá ser reemplazada por un endpoint Fetch() en producción.
        const matrizExcelData = [
            {
                "LICENCIA N°": "LI-00125",
                "APELLIDOS": "PÉREZ MAMANI",
                "NOMBRES": "JUAN CARLOS",
                "DNI": "70123456",
                "LIC CONDUCIR": "Q70123456",
                "CLASE": "A-IIIB",
                "F. EMISIÓN MTC": "15/03/2022",
                "F. CADUCIDAD": "15/03/2026",
                "SITUACION LICENCIA MTC": "VIGENTE",
                "EQUIPO AUTORIZADO": "MIXER / BOMBA TELESCÓPICA",
                "RESTRICCIONES": "USO DE LENTES CORRECTORES",
                "CARGO": "OPERADOR DE BOMBA"
            },
            {
                "LICENCIA N°": "LI-00126",
                "APELLIDOS": "QUISPE CONDORI",
                "NOMBRES": "MIGUEL ANGEL",
                "DNI": "40654321",
                "LIC CONDUCIR": "Q40654321",
                "CLASE": "A-IIIC",
                "F. EMISIÓN MTC": "10/01/2020",
                "F. CADUCIDAD": "10/01/2025",
                "SITUACION LICENCIA MTC": "VENCIDA",
                "EQUIPO AUTORIZADO": "CARGADOR FRONTAL",
                "RESTRICCIONES": "NINGUNA",
                "CARGO": "OPERADOR DE MAQUINARIA PESADA"
            }
        ];

        function buscarLicencia() {
            const inputDNI = document.getElementById("dni-input").value.trim();
            const errorMsg = document.getElementById("error-message");
            const ficha = document.getElementById("ficha-result");

            // Resetear visibilidad
            errorMsg.style.display = "none";
            ficha.style.display = "none";

            if(inputDNI.length === 0) {
                alert("Por favor, ingrese un DNI válido.");
                return;
            }

            // Lógica de búsqueda en la matriz (En producción, aquí va un fetch() asíncrono al JSON exportado de OneDrive)
            const resultado = matrizExcelData.find(trabajador => trabajador.DNI === inputDNI);

            if (resultado) {
                // Inyectar los datos en el DOM
                document.getElementById("res-lic-interna").innerText = resultado["LICENCIA N°"];
                document.getElementById("res-dni").innerText = resultado["DNI"];
                document.getElementById("res-apellidos").innerText = resultado["APELLIDOS"];
                document.getElementById("res-nombres").innerText = resultado["NOMBRES"];
                document.getElementById("res-cargo").innerText = resultado["CARGO"];
                document.getElementById("res-equipo").innerText = resultado["EQUIPO AUTORIZADO"];
                document.getElementById("res-lic-mtc").innerText = resultado["LIC CONDUCIR"];
                document.getElementById("res-clase").innerText = resultado["CLASE"];
                document.getElementById("res-emision").innerText = resultado["F. EMISIÓN MTC"];
                document.getElementById("res-caducidad").innerText = resultado["F. CADUCIDAD"];
                
                // Formateo condicional para situación de licencia
                const situacionElem = document.getElementById("res-situacion");
                situacionElem.innerText = resultado["SITUACION LICENCIA MTC"];
                if(resultado["SITUACION LICENCIA MTC"].toUpperCase() === "VIGENTE") {
                    situacionElem.className = "data-value status-vigente";
                } else {
                    situacionElem.className = "data-value status-vencido";
                }

                document.getElementById("res-restricciones").innerText = resultado["RESTRICCIONES"];

                // Mostrar la ficha técnica
                ficha.style.display = "block";
            } else {
                // Mostrar error si no se encuentra
                errorMsg.style.display = "block";
            }
        }
    </script>
</body>
</html>
