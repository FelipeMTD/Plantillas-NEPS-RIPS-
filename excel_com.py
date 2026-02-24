import win32com.client as win32
from pathlib import Path
import re
import math

XL_UP = -4162
CONTROL_SHEET = "__RIPS_CONTROL__"
_re_non_digits = re.compile(r"\D+")

def norm_doc(v):
    if v is None or isinstance(v, bool): return ""
    if isinstance(v, (int, float)):
        return str(int(round(v))) if not (isinstance(v, float) and not math.isfinite(v)) else ""
    s = str(v).strip()
    if not s: return ""
    m = re.fullmatch(r"(\d+)\.0+", s)
    if m: return m.group(1)
    return _re_non_digits.sub("", s)

class ExcelCOM:
    def __init__(self, path_xlsm: Path):
        self.path = str(path_xlsm.resolve())
        self.excel = None
        self.wb = None
        self.ws_estructura = None
        self.ws_us = None
        self.ws_control = None
        self.seen_us = set()

    def abrir(self):
        print(f"⏳ Abriendo Excel (Modo Oculto)...")
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False  # <--- OCULTO
        self.excel.DisplayAlerts = False
        self.wb = self.excel.Workbooks.Open(self.path)
        self.ws_estructura = self.wb.Worksheets("ESTRUCTURA")
        self.ws_us = self.wb.Worksheets("US")
        self._init_control()
        self._load_seen_us()

    def _init_control(self):
        try:
            self.ws_control = self.wb.Worksheets(CONTROL_SHEET)
        except Exception:
            self.ws_control = self.wb.Worksheets.Add()
            self.ws_control.Name = CONTROL_SHEET
            self.ws_control.Visible = 2 

    def _load_seen_us(self):
        try:
            row = 2
            while True:
                val = self.ws_control.Cells(row, 2).Value
                if not val: break
                self.seen_us.add(str(val))
                row += 1
        except:
            pass

    def append_us_control_batch(self, docs):
        if not docs: return
        try:
            start = self.ws_control.Cells(self.ws_control.Rows.Count, 1).End(XL_UP).Row + 1
            data = [["U", d] for d in docs]
            self.ws_control.Range(f"A{start}:B{start+len(data)-1}").Value = data
        except:
            pass

    def cerrar(self):
        if self.wb:
            self.wb.Save()
            self.wb.Close()
        if self.excel:
            self.excel.Quit()
        print("🚪 Excel guardado y cerrado.")

    def siguiente_fila(self, ws, col):
        last = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
        # Estructura empieza en 3 para respetar fórmulas de fila 2
        start = 3 if ws.Name == "ESTRUCTURA" else 2
        return max(start, last + 1)

    def pegar_estructura_matriz(self, filas, fila_inicio):
        """
        Pega los datos en ESTRUCTURA. 
        La primera posición del buffer (Columna C) ahora llegará vacía desde el main.
        """
        if not filas: return 0
        cant_filas = len(filas)
        end_row = fila_inicio + cant_filas - 1
        
        # Rango de pegado C:J (C será sobreescrita con vacío si no se mapea)
        self.ws_estructura.Range(f"C{fila_inicio}:J{end_row}").Value = filas
        return cant_filas
    
    def pegar_us_rango(self, filas, fila_inicio):
        nuevos = []
        keys = []
        for r in filas:
            if len(r) < 2: continue
            doc_limpio = norm_doc(r[1])
            key = f"{r[0]}|{doc_limpio}"
            
            if key not in self.seen_us:
                r[1] = doc_limpio
                nuevos.append(r)
                self.seen_us.add(key)
                keys.append(key)
        
        if nuevos:
            end = fila_inicio + len(nuevos) - 1
            self.ws_us.Range(f"A{fila_inicio}:N{end}").Value = nuevos
            self.append_us_control_batch(keys)
            print(f"      ✅ {len(nuevos)} usuarios nuevos insertados.")
            return end + 1
        else:
            return fila_inicio