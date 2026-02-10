import os
from pathlib import Path

def actualizar_index(datos):
    # 1. Recuperamos la ruta de destino que mandamos desde el script principal
    # Si no existe, usamos la carpeta actual por defecto
    ruta_carpeta = datos.get('ruta_destino', os.getcwd())
    
    # 2. El archivo que vamos a GUARDAR (en la carpeta de prueba)
    ruta_html_destino = os.path.join(ruta_carpeta, "index.html")
    
    # 3. El archivo que usamos como PLANTILLA (ahora en la carpeta /web)
    # Subimos un nivel desde 'scripts' para entrar a 'web'
    BASE_DIR = Path(__file__).resolve().parent.parent
    ruta_plantilla = os.path.join(BASE_DIR, "web", "index.html")

    if not os.path.exists(ruta_plantilla): 
        print(f"❌ Error: No se encontró la plantilla en {ruta_plantilla}")
        return

    with open(ruta_plantilla, "r", encoding="utf-8") as f:
        html = f.read()

    # --- Reemplazos (Esto se queda igual porque ya funciona bien) ---
    html = html.replace("{{ fecha }}", str(datos['fecha']))
    html = html.replace("{{ v_fisico }}", f"{datos['v_fisico']:,.0f}")
    html = html.replace("{{ v_transito }}", f"{datos['v_transito']:,.0f}")
    html = html.replace("{{ v_total }}", f"{datos['v_total']:,.0f}")
    html = html.replace("{{ doh }}", f"{datos['doh']:,.2f}" if isinstance(datos['doh'], (int,float)) else str(datos['doh']))
    html = html.replace("{{ v_diaria }}", f"{datos['v_diaria']:,.0f}")

    html = html.replace("{{ m1_n }}", str(datos['m1_n']))
    html = html.replace("{{ m1_v }}", f"{datos['m1_v']:,.0f}")
    html = html.replace("{{ m2_n }}", str(datos['m2_n']))
    html = html.replace("{{ m2_v }}", f"{datos['m2_v']:,.0f}")
    html = html.replace("{{ m3_n }}", str(datos['m3_n']))
    html = html.replace("{{ m3_v }}", f"{datos['m3_v']:,.0f}")

    # 1. Reemplazo para las tarjetas de Entradas
    html = html.replace("{{ e_ayer }}", f"{datos['e_ayer']:,.0f}")
    html = html.replace("{{ e_mes }}", f"{datos['e_mes']:,.0f}")

    # 2. Generación de las filas de la tabla Top 10
    filas_tabla = ""
    for item in datos['top_10']:
        # Convertimos el importe a número por si acaso viene como string
        valor = float(item['Importe2']) if isinstance(item['Importe2'], (str, float, int)) else 0
        filas_tabla += f"""
        <tr>
            <td>{str(item['Nombre']).upper()}</td>
            <td style='text-align: right; font-weight: bold;'>${valor:,.0f}</td>
        </tr>
        """
    
    # 3. Reemplazamos el hueco de la tabla en el HTML
    html = html.replace("{{ filas_top_10 }}", filas_tabla)

    # 4. GUARDAMOS el resultado en la carpeta de prueba
    with open(ruta_html_destino, "w", encoding="utf-8") as f:
        f.write(html)