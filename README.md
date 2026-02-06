# 🚀 Pipeline de Automatización: Valor de Inventario y Portal Ejecutivo

Este proyecto resuelve la necesidad crítica de visualizar el valor financiero del stock en tiempo real, consolidando datos heterogéneos en una solución automatizada, interactiva y portátil.

## 💡 Valor del Proyecto
Como **Analytics Engineer**, desarrollé este sistema para eliminar el procesamiento manual de reportes diarios. El pipeline transforma datos crudos en insights ejecutivos, reduciendo el tiempo de consolidación de horas a **segundos** y garantizando la integridad referencial entre 5 fuentes de datos distintas:
* Inventarios (Existencias y Costos)
* Tránsitos (Movimientos en camino)
* DOH (Days On Hand / Días de Inventario)
* OC (Órdenes de Compra Pendientes)
* Entradas (Recepciones de mercancía)

## 🛠️ Stack Tecnológico
* **Python (Pandas):** Motor ETL avanzado para limpieza, transformación y cálculos de métricas financieras.
* **Matplotlib:** Generación automatizada de visualizaciones de datos (comportamiento semanal vs. objetivos).
* **Openpyxl:** Engine de automatización para la generación de reportes maestros en Excel con formato contable e inserción de branding corporativo.
* **HTML5, CSS3 & JS:** Desarrollo de un portal web ejecutivo con diseño responsivo (Bootstrap 5) e interactividad para visualización de KPIs.

## ⚙️ Inteligencia de Rutas y Portabilidad
El sistema integra una **lógica de detección de entorno (Environment Awareness)**. Mediante el uso de la librería `pathlib`, el script identifica si tiene acceso a la red corporativa. De lo contrario, activa automáticamente el **"Modo Demo"**, utilizando el directorio `data_samples` y protegiendo los resultados en una carpeta local de salida. Esto permite que el portafolio sea **100% ejecutable** en cualquier entorno local de forma inmediata.



## 📂 Arquitectura del Proyecto
La solución se organiza de forma modular para garantizar la escalabilidad y el orden profesional:

* **`pipeline_valor_inventario_github/`**: Raíz del proyecto.
* **`├── requirements.txt`**: Lista de dependencias para la reproducción exacta del entorno.
* **`├── data_samples/`**: Datasets anonimizados para pruebas del pipeline.
* **`├── scripts/`**: Lógica de procesamiento (`valor_inventario.py`) y motor de renderizado web (`actualizar_portal.py`).
* **`├── web/`**: Plantilla base (`index.html`) y recursos visuales (logos e imágenes).
* **`└── output/`**: Directorio de salida generado automáticamente con el reporte Excel y el Portal Web.

## 🚀 Guía de Ejecución
1.  **Clonar el repositorio** en tu máquina local.
2.  **Instalar dependencias**:
    ```bash
    pip install -r pipeline_valor_inventario_github/requirements.txt
    ```
3.  **Ejecutar el pipeline**:
    ```bash
    python pipeline_valor_inventario_github/scripts/valor_inventario.py
    ```
4.  **Consultar resultados** en el directorio `output/` recién creado.

---
> **Nota de Privacidad:** Los datos en `data_samples/` han sido anonimizados y los valores numéricos alterados para proteger la confidencialidad de la información original, manteniendo intacta la lógica funcional y financiera del sistema.
