import shutil
import os

def reset_excel():
    # Definimos las rutas exactas
    plantilla = r"C:\Users\FELIPE SISTEMAS\Documents\GOYE\NUEVAEPS\PLANTILLA\COPIA_LIMPIA.xlsm"
    destino = r"C:\Users\FELIPE SISTEMAS\Documents\GOYE\NUEVAEPS\COPIA_LIMPIA.xlsm"

    try:
        # shutil.copy2 copia el archivo y mantiene los metadatos (fecha de creación, etc.)
        # Si el archivo de destino ya existe, lo sobrescribe automáticamente.
        shutil.copy2(plantilla, destino)
        print("✅ Reset completado: El Excel ha sido restaurado con la copia limpia.")
    except FileNotFoundError:
        print("❌ Error: No se encontró el archivo en la carpeta PLANTILLA.")
    except Exception as e:
        print(f"❌ Ocurrió un error inesperado: {e}")

# Llamar a la función
reset_excel()
