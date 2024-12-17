import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

def seleccionar_carpeta():
    """Abre un diálogo para que el usuario seleccione una carpeta"""
    root = tk.Tk()
    root.withdraw()
    carpeta_seleccionada = filedialog.askdirectory(title="Selecciona la carpeta a analizar")
    return carpeta_seleccionada

def copiar_fechas(original, nuevo):
    """Copia la fecha de creación y modificación del archivo original al archivo nuevo"""
    stat_info = os.stat(original)
    os.utime(nuevo, (stat_info.st_atime, stat_info.st_mtime))
    if os.name == 'nt':  # Solo para Windows
        import pywintypes
        import win32file
        import win32con
        
        handle = win32file.CreateFile(nuevo, win32con.GENERIC_WRITE, 0, None, win32con.OPEN_EXISTING, 0, None)
        win32file.SetFileTime(handle, pywintypes.Time(stat_info.st_ctime), pywintypes.Time(stat_info.st_atime), pywintypes.Time(stat_info.st_mtime))
        handle.close()

def convertir_archivos(carpeta):
    """Busca archivos .xlsm en la carpeta principal y subcarpetas, los convierte a .xlsx y elimina el original"""
    archivos_convertidos = []

    # Recorrer todas las carpetas y subcarpetas
    for ruta_raiz, _, archivos in os.walk(carpeta):
        for archivo in archivos:
            if archivo.lower().endswith(".xlsm"):  # Asegura que sea .xlsm
                ruta_completa = os.path.join(ruta_raiz, archivo)
                nombre_base = os.path.splitext(archivo)[0]
                ruta_nueva = os.path.join(ruta_raiz, f"{nombre_base}.xlsx")

                try:
                    print(f"Procesando archivo: {ruta_completa}")
                    # Convertir archivo
                    workbook = load_workbook(ruta_completa, keep_vba=False)
                    workbook.save(ruta_nueva)
                    
                    # Copiar fechas del archivo original
                    copiar_fechas(ruta_completa, ruta_nueva)
                    
                    # Eliminar archivo original
                    os.remove(ruta_completa)

                    # Agregar detalles al reporte
                    archivos_convertidos.append({
                        "Nombre del archivo": archivo,
                        "Ruta completa": ruta_nueva
                    })

                    print(f"Convertido y eliminado: {ruta_completa}")

                except Exception as e:
                    print(f"Error al procesar {ruta_completa}: {e}")

    return archivos_convertidos

def generar_reporte(archivos_convertidos):
    """Genera un reporte en formato Excel con la información de los archivos convertidos"""
    if not archivos_convertidos:
        print("No se encontraron archivos para convertir.")
        return

    # Guardar reporte en la carpeta actual donde se ejecuta el script
    ruta_actual = os.getcwd()
    ruta_reporte = os.path.join(ruta_actual, "reporte_conversion.xlsx")
    df = pd.DataFrame(archivos_convertidos)
    df.to_excel(ruta_reporte, index=False)
    print(f"Reporte generado: {ruta_reporte}")

def main():
    print("Programa para convertir archivos .xlsm a .xlsx, mantener fechas y generar reporte.")
    carpeta = seleccionar_carpeta()

    if not carpeta:
        print("No se seleccionó ninguna carpeta. Finalizando...")
        return

    print(f"Analizando carpeta: {carpeta}")
    archivos_convertidos = convertir_archivos(carpeta)
    generar_reporte(archivos_convertidos)

if __name__ == "__main__":
    main()
