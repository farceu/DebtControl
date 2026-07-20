import os
import sys
import pyodbc
from dotenv import load_dotenv
from PIL import Image, ImageChops
from playwright.sync_api import sync_playwright

load_dotenv()


def recortar_margenes_blancos(path, padding=10):
    """Recorta el espacio en blanco sobrante alrededor del contenido real de la imagen."""
    img = Image.open(path).convert("RGB")
    fondo = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, fondo)
    bbox = diff.getbbox()
    if bbox is None:
        return
    izq, arriba, der, abajo = bbox
    izq = max(izq - padding, 0)
    arriba = max(arriba - padding, 0)
    der = min(der + padding, img.width)
    abajo = min(abajo + padding, img.height)
    img.crop((izq, arriba, der, abajo)).save(path)


# 2. Configuración de conexión a SQL Server (se toma de variables de entorno)
def procesar_pendientes():
    conn_str = (
        "Driver={ODBC Driver 17 for SQL Server};"
        f"Server={os.environ['DB_SERVER']};"
        f"Database={os.environ['DB_NAME']};"
        f"UID={os.environ['DB_USER']};PWD={os.environ.get('DB_PASSWORD', '')};"
    )

    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        # Reemplaza 'CampoHTML', 'TuTabla' e 'Id' con tus nombres reales
        cursor.execute("SELECT nkey_mail, isnull(textBodycompleto,' ') FROM mail WHERE enviado = 'X' and esnuevomail='S' and isnull(generoImagen,'N') ='N'")
        registros = cursor.fetchall()
        
        if not registros:
            print("No se encontró registros")
            sys.exit(1)

        with sync_playwright() as p:
            browser = p.chromium.launch()
            # device_scale_factor no cambia el layout (mismo ancho/alto en CSS),
            # solo aumenta la densidad de píxeles de la captura -> más nitidez
            page = browser.new_page(device_scale_factor=2)

            for row in registros:
                reg_id = row[0]
                html_content = row[1]
                output_path = f"Z:/Ejecutables/salida/temp/{reg_id}.png"

                try:

                    # Generar imagen
                    page.set_content(html_content)
                    page.screenshot(path=output_path, full_page=True)

                    # El viewport es más ancho que el contenido real -> recortar el
                    # espacio en blanco sobrante para que no se vea diminuto en el mail
                    recortar_margenes_blancos(output_path)

                    # 3. Marcar como procesado en la base de datos
                    cursor.execute("UPDATE mail SET generoImagen = 'S' WHERE nkey_mail = ?", reg_id)
                    conn.commit()

                except Exception as err:
                    print(f"❌ Error procesando ID {reg_id}: {err}")

            browser.close()
        
        conn.close()
        print("Proceso finalizado.")

    except Exception as e:
        print(f"Error crítico: {e}")

if __name__ == "__main__":
    procesar_pendientes()
sys.exit(0)