from pathlib import Path
import csv

def normalizar_carpeta_csv(carpeta):
    for csv_file in Path(carpeta).glob("*.CSV"):
        with csv_file.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        with csv_file.open("w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rows)