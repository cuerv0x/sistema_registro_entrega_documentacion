#!/usr/bin/env python3

import pandas as pd
import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import logging

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class GestorNotas:
    def __init__(self, archivo_excel: str):
        self.archivo_excel = archivo_excel
        self.datos = None
        self.cargar_datos()

    def cargar_datos(self) -> None:
        try:
            if not os.path.exists(self.archivo_excel):
                raise FileNotFoundError(f"El archivo {self.archivo_excel} no existe")

            primera_fila = pd.read_excel(self.archivo_excel, nrows=1)

            palabras_encabezado = [
                "columna",
                "atinencia",
                "nombre",
                "modulos",
                "módulos",
            ]

            primera_fila_texto = " ".join(
                str(valor).lower() for valor in primera_fila.values.flatten()
            )
            contiene_palabras_encabezado = any(
                palabra in primera_fila_texto for palabra in palabras_encabezado
            )

            if contiene_palabras_encabezado:
                logger.info(
                    "Primera fila contiene palabras de encabezado. Saltando primera fila..."
                )
                self.datos = pd.read_excel(self.archivo_excel, skiprows=1)
                logger.info("Datos cargados saltando la primera fila")
            else:
                self.datos = pd.read_excel(self.archivo_excel)
                logger.info(
                    "Datos cargados normalmente (primera fila como encabezados)"
                )

            logger.info(f"Datos cargados exitosamente desde {self.archivo_excel}")
            logger.info(f"Columnas disponibles: {list(self.datos.columns)}")
            logger.info(f"Número de registros: {len(self.datos)}")

        except Exception as e:
            logger.error(f"Error al cargar el archivo Excel: {e}")
            raise

    def buscar_por_cedula(self, cedula: str) -> Optional[Dict]:
        try:
            cedula_limpia = cedula.replace("-", "").replace(" ", "")

            columnas_cedula = []
            for col in self.datos.columns:
                if any(
                    palabra in col.lower()
                    for palabra in [
                        "cedula",
                        "cédula",
                        "identificacion",
                        "identificación",
                        "dni",
                        "id",
                    ]
                ):
                    columnas_cedula.append(col)

            if not columnas_cedula:
                columnas_cedula = self.datos.columns

            for columna in columnas_cedula:
                self.datos[columna] = self.datos[columna].astype(str)

                coincidencias = self.datos[
                    self.datos[columna].str.contains(
                        cedula_limpia, na=False, regex=False
                    )
                ]

                if not coincidencias.empty:
                    estudiante = coincidencias.iloc[0].to_dict()
                    logger.info(f"Estudiante encontrado en columna '{columna}'")
                    return estudiante

            logger.warning(f"No se encontró estudiante con cédula {cedula}")
            return None

        except Exception as e:
            logger.error(f"Error al buscar por cédula: {e}")
            return None

    def mostrar_informacion_estudiante(self, cedula: str) -> None:
        estudiante = self.buscar_por_cedula(cedula)

        if estudiante is None:
            print(f"\nNo se encontró información para la cédula: {cedula}")
            return

        print(f"\n{'='*60}")
        print(f"INFORMACIÓN DEL ESTUDIANTE")
        print(f"{'='*60}")

        for campo, valor in estudiante.items():
            if pd.notna(valor) and str(valor).strip() != "":
                print(f"{campo}: {valor}")

        print(f"{'='*60}")

    def generar_excel_con_timestamp(
        self, nombre_base: str = "reporte_estudiantes"
    ) -> str:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_archivo = f"{nombre_base}_{timestamp}.xlsx"

            info_general = pd.DataFrame(
                {
                    "Fecha_Generacion": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                    "Archivo_Origen": [self.archivo_excel],
                    "Total_Registros": [len(self.datos)],
                    "Columnas_Disponibles": [", ".join(self.datos.columns)],
                }
            )

            with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:
                info_general.to_excel(
                    writer, sheet_name="Informacion_General", index=False
                )

                self.datos.to_excel(writer, sheet_name="Datos_Completos", index=False)

                resumen_columnas = pd.DataFrame(
                    {
                        "Columna": self.datos.columns,
                        "Tipo_Dato": self.datos.dtypes.astype(str),
                        "Valores_No_Nulos": self.datos.count(),
                        "Valores_Unicos": [
                            self.datos[col].nunique() for col in self.datos.columns
                        ],
                    }
                )
                resumen_columnas.to_excel(
                    writer, sheet_name="Resumen_Columnas", index=False
                )

            logger.info(f"Archivo Excel generado: {nombre_archivo}")
            return nombre_archivo

        except Exception as e:
            logger.error(f"Error al generar Excel: {e}")
            raise

    def generar_html5_interactivo(self, nombre_archivo: str = None) -> str:
        try:
            if nombre_archivo is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"tabla_estudiantes_{timestamp}.html"

            columnas_especificas = [
                "Cédula",
                "Nombre",
                "Módulos",
                "Temario",
                "Tabla de Especificaciones",
                "Prueba Escrita",
            ]

            df_html = pd.DataFrame()

            for i, col_especifica in enumerate(columnas_especificas):
                if i < len(self.datos.columns):
                    df_html[col_especifica] = self.datos.iloc[:, i]
                else:
                    df_html[col_especifica] = ""

            datos_json = df_html.to_json(orient="records", force_ascii=False)

            html_content = f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Búsqueda de Estudiantes</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        
        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .search-section {{
            padding: 30px;
            background: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }}
        
        .search-box {{
            display: flex;
            gap: 15px;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
        }}
        
        .search-input {{
            padding: 15px 20px;
            border: 2px solid #ddd;
            border-radius: 25px;
            font-size: 16px;
            width: 300px;
            transition: all 0.3s ease;
        }}
        
        .search-input:focus {{
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 15px rgba(102, 126, 234, 0.3);
        }}
        
        .search-btn {{
            padding: 15px 30px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 25px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s ease;
        }}
        
        .search-btn:hover {{
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }}
        
        .clear-btn {{
            padding: 15px 30px;
            background: #6c757d;
            color: white;
            border: none;
            border-radius: 25px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s ease;
        }}
        
        .clear-btn:hover {{
            background: #5a6268;
            transform: translateY(-2px);
        }}
        
        .table-container {{
            padding: 30px;
            overflow-x: auto;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }}
        
        .data-table th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
            white-space: nowrap;
        }}
        
        .data-table td {{
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
            transition: background-color 0.3s ease;
            max-width: 200px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        
        .data-table tr:hover {{
            background-color: #f8f9fa;
        }}
        
        .data-table tr:hover td {{
            white-space: normal;
            word-wrap: break-word;
        }}
        
        .no-results {{
            text-align: center;
            padding: 50px;
            color: #6c757d;
            font-size: 1.2em;
        }}
        
        .stats {{
            padding: 20px 30px;
            background: #f8f9fa;
            border-top: 1px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
        }}
        
        .stat-item {{
            text-align: center;
        }}
        
        .stat-number {{
            font-size: 1.5em;
            font-weight: bold;
            color: #667eea;
        }}
        
        .stat-label {{
            color: #6c757d;
            font-size: 0.9em;
        }}
        
        .highlight {{
            background-color: #fff3cd !important;
            font-weight: bold;
        }}
        
        @media (max-width: 768px) {{
            .search-box {{
                flex-direction: column;
            }}
            
            .search-input {{
                width: 100%;
                max-width: 300px;
            }}
            
            .data-table {{
                font-size: 14px;
            }}
            
            .data-table th,
            .data-table td {{
                padding: 8px 10px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Sistema de Búsqueda de Estudiantes</h1>
            <p>Busque información de estudiantes por número de cédula</p>
            <div class="info">
                Datos combinados de {len(self.archivos_origen)} archivos Excel
            </div>
        </div>
        
        <div class="search-section">
            <div class="search-box">
                <input type="text" id="searchInput" class="search-input" 
                       placeholder="Ingrese número de cédula exacto..." 
                       onkeyup="searchTable()">
                <button onclick="searchTable()" class="search-btn">Buscar</button>
                <button onclick="clearSearch()" class="clear-btn">Limpiar</button>
            </div>
        </div>
        
        <div class="table-container">
            <table id="dataTable" class="data-table">
                <thead>
                    <tr>
                        <th>Cédula</th>
                        <th>Nombre</th>
                        <th>Módulos</th>
                        <th>Temario</th>
                        <th>Tabla de Especificaciones</th>
                        <th>Prueba Escrita</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                </tbody>
            </table>
            <div id="initialMessage" class="no-results">
                <p>Ingrese un número de cédula para buscar estudiantes</p>
            </div>
            <div id="noResults" class="no-results" style="display: none;">
                <p>No se encontraron resultados para la búsqueda</p>
            </div>
            <div id="errorMessage" class="no-results" style="display: none;">
                <p>Error: Solo se permiten números. No use guiones (-) ni otros caracteres.</p>
            </div>
        </div>
        
        <div class="stats">
            <div class="stat-item">
                <div class="stat-number" id="totalRecords">0</div>
                <div class="stat-label">Total de Registros</div>
            </div>
            <div class="stat-item">
                <div class="stat-number" id="filteredRecords">0</div>
                <div class="stat-label">Registros Filtrados</div>
            </div>
            <div class="stat-item">
                <div class="stat-number">6</div>
                <div class="stat-label">Columnas</div>
            </div>
            <div class="stat-item">
                <div class="stat-number">{len(df_html)}</div>
                <div class="stat-label">Registros Originales</div>
            </div>
        </div>
    </div>

    <script>
        const tableData = {datos_json};
        let filteredData = [...tableData];
        
        function renderTable(data) {{
            const tbody = document.getElementById('tableBody');
            const dataTable = document.getElementById('dataTable');
            const noResults = document.getElementById('noResults');
            const initialMessage = document.getElementById('initialMessage');
            const errorMessage = document.getElementById('errorMessage');
            
            tbody.innerHTML = '';
            
            dataTable.style.display = 'none';
            noResults.style.display = 'none';
            initialMessage.style.display = 'none';
            errorMessage.style.display = 'none';
            
            if (data.length === 0) {{
                if (document.getElementById('searchInput').value.trim() === '') {{
                    initialMessage.style.display = 'block';
                }} else {{
                    noResults.style.display = 'block';
                }}
            }} else {{
                dataTable.style.display = 'table';
                
                data.forEach(row => {{
                    const tr = document.createElement('tr');
                    const columns = ['Cédula', 'Nombre', 'Módulos', 'Temario', 'Tabla de Especificaciones', 'Prueba Escrita'];
                    
                    columns.forEach(col => {{
                        const td = document.createElement('td');
                        td.textContent = row[col] || '';
                        tr.appendChild(td);
                    }});
                    tbody.appendChild(tr);
                }});
            }}
            
            updateStats(data.length);
        }}
        
        function searchTable() {{
            const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
            
            if (searchTerm === '') {{
                filteredData = [];
            }} else {{
                filteredData = tableData.filter(row => {{
                    const cedula = String(row['Cédula'] || '').toLowerCase();
                    return cedula === searchTerm;
                }});
            }}
            
            renderTable(filteredData);
        }}
        
        function clearSearch() {{
            document.getElementById('searchInput').value = '';
            filteredData = [];
            renderTable(filteredData);
        }}
        
        function updateStats(filteredCount) {{
            document.getElementById('totalRecords').textContent = tableData.length;
            document.getElementById('filteredRecords').textContent = filteredCount;
        }}
        
        document.addEventListener('DOMContentLoaded', function() {{
            renderTable([]);
        }});
        
        document.getElementById('searchInput').addEventListener('keypress', function(e) {{
            if (e.key === 'Enter') {{
                searchTable();
            }}
        }});
        
        document.getElementById('searchInput').addEventListener('input', function(e) {{
            const value = e.target.value;
            const cleanValue = value.replace(/[^0-9]/g, '');
            if (value !== cleanValue) {{
                e.target.value = cleanValue;
            }}
        }});
    </script>
</body>
</html>
"""

            with open(nombre_archivo, "w", encoding="utf-8") as f:
                f.write(html_content)

            logger.info(f"Archivo HTML5 generado: {nombre_archivo}")
            return nombre_archivo

        except Exception as e:
            logger.error(f"Error al generar HTML5: {e}")
            raise


def main() -> None:
    print("\033[1;36mSistema de registro de entrega de documentación en evaluación\033[0m")
    print("\033[1;36m" + "=" * 50 + "\033[0m")

    archivos_excel = [
        f
        for f in os.listdir(".")
        if f.endswith(".xlsx") and not f.startswith("~$") and os.path.isfile(f)
    ]

    if not archivos_excel:
        print("\033[1;31mNo se encontraron archivos Excel (.xlsx) en el directorio actual\033[0m")
        print("\033[1;31mAsegúrese de que los archivos tengan extensión .xlsx\033[0m")
        return

    print("\033[1;32mArchivos Excel (.xlsx) encontrados:\033[0m")
    for i, archivo in enumerate(archivos_excel, 1):
        print(f"  \033[1;33m{i}.\033[0m {archivo}")

    print(f"\n\033[1;34mCombinando {len(archivos_excel)} archivos Excel...\033[0m")

    dataframes = []

    for archivo in archivos_excel:
        try:
            print(f"\033[1;35mLeyendo {archivo}...\033[0m")

            primera_fila = pd.read_excel(archivo, nrows=1)
            palabras_encabezado = [
                "columna",
                "atinencia",
                "nombre",
                "modulos",
                "módulos",
            ]
            primera_fila_texto = " ".join(
                str(valor).lower() for valor in primera_fila.values.flatten()
            )
            contiene_palabras_encabezado = any(
                palabra in primera_fila_texto for palabra in palabras_encabezado
            )

            if contiene_palabras_encabezado:
                df = pd.read_excel(archivo, skiprows=1)
                print(f"  \033[1;33mSaltando primera fila (encabezado detectado)\033[0m")
            else:
                df = pd.read_excel(archivo)
                print(f"  \033[1;32mLeyendo normalmente\033[0m")

            print(f"  \033[1;36mRegistros: {len(df)}, Columnas: {len(df.columns)}\033[0m")
            dataframes.append(df)

        except Exception as e:
            print(f"  \033[1;31mError al leer {archivo}: {str(e)[:50]}...\033[0m")
            continue

    if not dataframes:
        print("\033[1;31mNo se pudieron leer archivos Excel válidos\033[0m")
        return

    print(f"\n\033[1;34mCombinando {len(dataframes)} DataFrames...\033[0m")
    
    columnas_por_archivo = [len(df.columns) for df in dataframes]
    if len(set(columnas_por_archivo)) > 1:
        print("\033[1;31mERROR: Los archivos Excel tienen diferentes números de columnas:\033[0m")
        for i, (archivo, num_cols) in enumerate(zip(archivos_excel, columnas_por_archivo), 1):
            print(f"  \033[1;33m{i}.\033[0m {archivo}: \033[1;31m{num_cols} columnas\033[0m")
        print("\033[1;31mTodos los archivos deben tener el mismo número de columnas.\033[0m")
        return
    
    datos_combinados = pd.concat(dataframes, ignore_index=True)

    print(f"\033[1;32mDatos combinados exitosamente:\033[0m")
    print(f"   \033[1;36mTotal de registros: {len(datos_combinados)}\033[0m")

    gestor = GestorNotasCombinado(datos_combinados, archivos_excel)

    archivo_generado = gestor.generar_html5_interactivo()
    print(f"\033[1;32mArchivo HTML5 generado exitosamente:\033[0m \033[1;36m{archivo_generado}\033[0m")


class GestorNotasCombinado:
    def __init__(self, datos_combinados: pd.DataFrame, archivos_origen: List[str]):
        self.datos = datos_combinados
        self.archivos_origen = archivos_origen
        logger.info(f"Datos combinados cargados desde {len(archivos_origen)} archivos")
        logger.info(f"Total de registros: {len(self.datos)}")

    def generar_html5_interactivo(self, nombre_archivo: str = None) -> str:
        try:
            if nombre_archivo is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"tabla_estudiantes_combinada_{timestamp}.html"

            columnas_especificas = [
                "Cédula",
                "Nombre",
                "Módulos",
                "Temario",
                "Tabla de Especificaciones",
                "Prueba Escrita",
            ]

            df_html = pd.DataFrame()

            for i, col_especifica in enumerate(columnas_especificas):
                if i < len(self.datos.columns):
                    df_html[col_especifica] = self.datos.iloc[:, i]
                else:
                    df_html[col_especifica] = ""

            datos_json = df_html.to_json(orient="records", force_ascii=False)

            html_content = f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Búsqueda de Estudiantes - Datos Combinados</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        
        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
            margin-bottom: 10px;
        }}
        
        .header .info {{
            font-size: 0.9em;
            opacity: 0.8;
            background: rgba(255,255,255,0.1);
            padding: 10px;
            border-radius: 5px;
            margin-top: 10px;
        }}
        
        .search-section {{
            padding: 30px;
            background: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }}
        
        .search-box {{
            display: flex;
            gap: 15px;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
        }}
        
        .search-input {{
            padding: 15px 20px;
            border: 2px solid #ddd;
            border-radius: 25px;
            font-size: 16px;
            width: 300px;
            transition: all 0.3s ease;
        }}
        
        .search-input:focus {{
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 15px rgba(102, 126, 234, 0.3);
        }}
        
        .search-btn {{
            padding: 15px 30px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 25px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s ease;
        }}
        
        .search-btn:hover {{
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }}
        
        .clear-btn {{
            padding: 15px 30px;
            background: #6c757d;
            color: white;
            border: none;
            border-radius: 25px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s ease;
        }}
        
        .clear-btn:hover {{
            background: #5a6268;
            transform: translateY(-2px);
        }}
        
        .table-container {{
            padding: 30px;
            overflow-x: auto;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }}
        
        .data-table th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
            white-space: nowrap;
        }}
        
        .data-table td {{
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
            transition: background-color 0.3s ease;
            max-width: 200px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        
        .data-table tr:hover {{
            background-color: #f8f9fa;
        }}
        
        .data-table tr:hover td {{
            white-space: normal;
            word-wrap: break-word;
        }}
        
        .no-results {{
            text-align: center;
            padding: 50px;
            color: #6c757d;
            font-size: 1.2em;
        }}
        
        .stats {{
            padding: 20px 30px;
            background: #f8f9fa;
            border-top: 1px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
        }}
        
        .stat-item {{
            text-align: center;
        }}
        
        .stat-number {{
            font-size: 1.5em;
            font-weight: bold;
            color: #667eea;
        }}
        
        .stat-label {{
            color: #6c757d;
            font-size: 0.9em;
        }}
        
        .highlight {{
            background-color: #fff3cd !important;
            font-weight: bold;
        }}
        
        @media (max-width: 768px) {{
            .search-box {{
                flex-direction: column;
            }}
            
            .search-input {{
                width: 100%;
                max-width: 300px;
            }}
            
            .data-table {{
                font-size: 14px;
            }}
            
            .data-table th,
            .data-table td {{
                padding: 8px 10px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎓 Sistema de registro de entrega de documentación en evaluación</h1>
            <div class="info">
                📁 Datos combinados de {len(self.archivos_origen)} archivos Excel
            </div>
        </div>
        
        <div class="search-section">
            <div class="search-box">
                <input type="text" id="searchInput" class="search-input" 
                       placeholder="Ingrese número de cédula exacto..." 
                       onkeyup="searchTable()">
                <button onclick="searchTable()" class="search-btn">🔍 Buscar</button>
                <button onclick="clearSearch()" class="clear-btn">🗑️ Limpiar</button>
            </div>
        </div>
        
        <div class="table-container">
            <table id="dataTable" class="data-table">
                <thead>
                    <tr>
                        <th>Cédula</th>
                        <th>Nombre</th>
                        <th>Módulos</th>
                        <th>Temario</th>
                        <th>Tabla de Especificaciones</th>
                        <th>Prueba Escrita</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                </tbody>
            </table>
            <div id="initialMessage" class="no-results">
                <p>🔍 Ingrese un número de cédula para buscar estudiantes</p>
            </div>
            <div id="noResults" class="no-results" style="display: none;">
                <p>❌ No se encontraron resultados para la búsqueda</p>
            </div>
            <div id="errorMessage" class="no-results" style="display: none;">
                <p>⚠️ Error: Solo se permiten números. No use guiones (-) ni otros caracteres.</p>
            </div>
        </div>
        </div>
    </div>

    <script>
        const tableData = {datos_json};
        let filteredData = [...tableData];
        
        function renderTable(data) {{
            const tbody = document.getElementById('tableBody');
            const dataTable = document.getElementById('dataTable');
            const noResults = document.getElementById('noResults');
            const initialMessage = document.getElementById('initialMessage');
            const errorMessage = document.getElementById('errorMessage');
            
            tbody.innerHTML = '';
            
            dataTable.style.display = 'none';
            noResults.style.display = 'none';
            initialMessage.style.display = 'none';
            errorMessage.style.display = 'none';
            
            if (data.length === 0) {{
                if (document.getElementById('searchInput').value.trim() === '') {{
                    initialMessage.style.display = 'block';
                }} else {{
                    noResults.style.display = 'block';
                }}
            }} else {{
                dataTable.style.display = 'table';
                
                data.forEach(row => {{
                    const tr = document.createElement('tr');
                    const columns = ['Cédula', 'Nombre', 'Módulos', 'Temario', 'Tabla de Especificaciones', 'Prueba Escrita'];
                    
                    columns.forEach(col => {{
                        const td = document.createElement('td');
                        td.textContent = row[col] || '';
                        tr.appendChild(td);
                    }});
                    tbody.appendChild(tr);
                }});
            }}
            
            updateStats(data.length);
        }}
        
        function searchTable() {{
            const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
            
            if (searchTerm === '') {{
                filteredData = [];
            }} else {{
                filteredData = tableData.filter(row => {{
                    const cedula = String(row['Cédula'] || '').toLowerCase();
                    return cedula === searchTerm;
                }});
            }}
            
            renderTable(filteredData);
        }}
        
        function clearSearch() {{
            document.getElementById('searchInput').value = '';
            filteredData = [];
            renderTable(filteredData);
        }}
        
        function updateStats(filteredCount) {{
            document.getElementById('totalRecords').textContent = tableData.length;
            document.getElementById('filteredRecords').textContent = filteredCount;
        }}
        
        document.addEventListener('DOMContentLoaded', function() {{
            renderTable([]);
        }});
        
        document.getElementById('searchInput').addEventListener('keypress', function(e) {{
            if (e.key === 'Enter') {{
                searchTable();
            }}
        }});
        
        document.getElementById('searchInput').addEventListener('input', function(e) {{
            const value = e.target.value;
            const cleanValue = value.replace(/[^0-9]/g, '');
            if (value !== cleanValue) {{
                e.target.value = cleanValue;
            }}
        }});
    </script>
</body>
</html>
"""

            with open(nombre_archivo, "w", encoding="utf-8") as f:
                f.write(html_content)

            logger.info(f"Archivo HTML5 generado: {nombre_archivo}")
            return nombre_archivo

        except Exception as e:
            logger.error(f"Error al generar HTML5: {e}")
            raise


if __name__ == "__main__":
    main()
