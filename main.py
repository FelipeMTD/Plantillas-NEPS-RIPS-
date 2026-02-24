import csv
import zipfile
import shutil
from pathlib import Path
from excel_com import ExcelCOM
from texto_en_col import normalizar_carpeta_csv

BASE_DIR = Path(__file__).parent
ZIP_DIR = BASE_DIR / "zip"
WORK_DIR = BASE_DIR / "_work"
PLANTILLA = BASE_DIR / "COPIA_LIMPIA.xlsm"

# --- MATRIZ DE MAPEO CORREGIDA (IGNORANDO TIPO DE DOC) ---
# Origen CSV (Índice) -> Destino Excel (Letra)
# Los CSV tienen 2 columnas vacías al inicio. Índice 3 es Documento.
MAPEO_CONFIG = {
    "AC": {
        3: "C",  # Num Doc  -> Col D
        4: "D",
        9: "G",  # Código   -> Col G
        17: "H",  # Finalidad-> Col H
        18: "J"   # Causa    -> Col J
    },
    "AP": {
        3: "C",  # Num Doc  -> Col D
        4: "D",
        10: "G",  # Código   -> Col G
        15: "H",  # Finalidad-> Col H
        16: "J"   # Causa    -> Col J
    },
    "AT": {
        3: "C",  # Num Doc  -> Col D
        7: "H",  # Código   -> Col G
        11: "J"   # Nombre   -> Col H
        # Columna C (Excel) se mantiene vacía por diseño
    }
}

def extraer_zip(zip_path: Path) -> Path:
    destino = WORK_DIR / zip_path.stem
    if destino.exists(): shutil.rmtree(destino)
    destino.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(destino)
    return destino

def detectar_delimitador(filepath):
    with open(filepath, "r", encoding="utf-8-sig") as f:
        linea = f.readline()
        if ";" in linea and "," not in linea: return ";"
    return "," 

def letra_a_indice(letra):
    # 'C' es el índice 0 en nuestro buffer de 8 columnas (C a J)
    return ord(letra.upper()) - ord('C')

def procesar_zips():
    print("\n" + "="*60)
    print("🚀 PROCESANDO RIPS NUEVA EPS - MATRIZ C, D, G, H, J")
    print("="*60)

    excel = ExcelCOM(PLANTILLA)
    try:
        excel.abrir()
        fila_actual_est = 3 # Inicio forzado en fila 3
        
        zips = sorted(ZIP_DIR.glob("*.zip"))
        for i, zip_file in enumerate(zips, 1):
            print(f"[{i}/{len(zips)}] 📂 Carpeta: {zip_file.name}")
            carpeta = extraer_zip(zip_file)
            normalizar_carpeta_csv(carpeta)

            datos_est = []
            for tipo, mapeo in MAPEO_CONFIG.items():
                csv_path = next(carpeta.glob(f"{tipo}*.CSV"), None)
                if csv_path:
                    delim = detectar_delimitador(csv_path)
                    with open(csv_path, newline="", encoding="utf-8-sig") as f:
                        reader = csv.reader(f, delimiter=delim)
                        next(reader, None)
                        for r in reader:
                            if not r or len(r) < 5: continue
                            
                            # Buffer de 8 columnas (C, D, E, F, G, H, I, J)
                            row_buffer = [""] * 8 
                            for csv_idx, letra_excel in mapeo.items():
                                idx_buf = letra_a_indice(letra_excel)
                                if 0 <= idx_buf < 8:
                                    row_buffer[idx_buf] = r[csv_idx].strip()
                            datos_est.append(row_buffer)

            if datos_est:
                print(f"   💾 Insertando {len(datos_est)} filas en ESTRUCTURA...")
                excel.pegar_estructura_matriz(datos_est, fila_actual_est)
                fila_actual_est += len(datos_est)

            # Procesar US normalmente
            csv_us = next(carpeta.glob("US*.CSV"), None)
            if csv_us:
                delim = detectar_delimitador(csv_us)
                with open(csv_us, newline="", encoding="utf-8-sig") as f:
                    datos_us = [(r + [""] * 14)[:14] for r in csv.reader(f, delimiter=delim)][1:]
                if datos_us:
                    excel.pegar_us_rango(datos_us, excel.siguiente_fila(excel.ws_us, 2))

    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        excel.cerrar()
        print("✨ Proceso finalizado.")

if __name__ == "__main__":
    procesar_zips()