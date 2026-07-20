import os
import sys
import pyodbc
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()


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
            # device_scale_factor alto = captura más nítida (equivalente a pantalla retina)
            page = browser.new_page(viewport={"width": 1024, "height": 800}, device_scale_factor=2)

            for row in registros:
                reg_id = row[0]
                html_content = row[1]
                output_path = f"Z:/Ejecutables/salida/temp/{reg_id}.png"

                try:

                    # Generar imagen
                    page.set_content(html_content, wait_until="networkidle")

                    # Medir el tamaño real del contenido renderizado
                    ancho = page.evaluate("document.documentElement.scrollWidth")
                    alto = page.evaluate("document.documentElement.scrollHeight")

                    # Ajustar el viewport al tamaño exacto del contenido (tamaño dinámico)
                    page.set_viewport_size({"width": ancho, "height": alto})

                    page.screenshot(path=output_path)

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