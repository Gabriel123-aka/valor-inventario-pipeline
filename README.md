# 🚀 Pipeline de Automatización: Valor de Inventario y Portal Ejecutivo

Este proyecto resuelve la necesidad crítica de visualizar el valor financiero del stock en tiempo real, consolidando datos heterogéneos en una solución automatizada y portátil.

## 💡 Valor del Proyecto
Como **Analytics Engineer**, desarrollé este sistema para eliminar el procesamiento manual de reportes diarios. El pipeline reduce el tiempo de consolidación de horas a **segundos**, garantizando la integridad referencial entre 5 fuentes de datos distintas.

## 🛠️ Stack Tecnológico
* **Python (Pandas):** Motor ETL para limpieza, transformación y cálculos financieros de DOH (Days On Hand).
* **Matplotlib:** Generación dinámica de gráficas de comportamiento semanal y cumplimiento de objetivos.
* **HTML5 / Bootstrap:** Frontend interactivo y responsivo para visualización gerencial.
* **Openpyxl:** Automatización y formateo profesional de reportes maestros en Excel.

## ⚙️ Inteligencia de Rutas y Portabilidad
El sistema cuenta con una **lógica de detección de entorno**. Si detecta la red corporativa, opera en modo producción sincronizando con los servidores; de lo contrario, utiliza la carpeta `data_samples` para demostraciones funcionales, permitiendo que este portafolio sea 100% ejecutable en cualquier entorno local.

## 📂 Estructura del Repositorio
* **`data_samples/`**: Archivos fuente anonimizados para pruebas del pipeline.
* **`scripts/`**: Código fuente en Python (`valor_inventario.py` y `actualizar_portal.py`).
* **`web/`**: Plantilla HTML y recursos visuales del portal ejecutivo.
* **`output/`**: Directorio donde el sistema genera el Excel consolidado y el portal web final.

## 🚀 Cómo ejecutarlo
1. Clona el repositorio.
2. Asegúrate de tener instaladas las dependencias: `pip install pandas openpyxl matplotlib`.
3. Ejecuta el script principal: `python scripts/valor_inventario.py`.
4. Visualiza los resultados generados en la carpeta `output/`.

---
> **Nota de Privacidad:** Los datos en `data_samples/` han sido anonimizados y los valores numéricos alterados para proteger la confidencialidad de la empresa original, manteniendo intacta la lógica funcional del sistema.
