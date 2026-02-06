# 🚀 Pipeline de Automatización: Valor de Inventario y Portal Ejecutivo

Este proyecto resuelve la necesidad crítica de visualizar el valor financiero del stock en tiempo real, consolidando datos heterogéneos en una solución automatizada, interactiva y portátil.

## 💡 Valor del Proyecto
Como **Analytics Engineer**, desarrollé este sistema para eliminar el procesamiento manual de reportes diarios. El pipeline transforma datos crudos en insights ejecutivos, reduciendo el tiempo de consolidación de horas a **segundos** y garantizando la integridad referencial entre 5 fuentes de datos distintas (Inventarios, Tránsitos, DOH, OC y Entradas).

## 🛠️ Stack Tecnológico
* **Python (Pandas):** Motor ETL avanzado para limpieza, transformación y cálculos de métricas financieras como DOH (Days On Hand) y variaciones diarias.
* **Matplotlib:** Generación automatizada de visualizaciones de datos (comportamiento semanal vs. objetivos).
* **HTML5 & CSS3:** Estructura y diseño de interfaz con branding corporativo y animaciones personalizadas.
* **JavaScript:** Lógica de interactividad para la navegación del portal y dinamismo en la presentación de métricas y animaciones.
* **Bootstrap 5:** Framework para garantizar un diseño responsivo y moderno.
* **Openpyxl:** Engine de automatización para la generación de reportes maestros en Excel con formato contable.

## ⚙️ Inteligencia de Rutas y Portabilidad
El sistema integra una **lógica de detección de entorno (Environment Awareness)**. Mediante el uso de la librería `pathlib`, el script identifica si tiene acceso a la red corporativa. De lo contrario, se autoconfigura para utilizar el directorio `data_samples`, permitiendo que este portafolio sea **100% ejecutable** en cualquier entorno local de forma inmediata.

## 📂 Arquitectura del Proyecto
La solución se organiza bajo la carpeta raíz `pipeline_valor_inventario_github` para mantener una estructura modular y profesional:

* **`pipeline_valor_inventario_github/data_samples/`**: Datasets anonimizados para pruebas del pipeline.
* **`pipeline_valor_inventario_github/scripts/`**: Lógica de procesamiento de datos (`valor_inventario.py`) y motor de renderizado web (`actualizar_portal.py`).
* **`pipeline_valor_inventario_github/web/`**: Plantilla base (`index.html`), activos visuales y lógica de estilos.
* **`output/`**: Directorio de salida generado automáticamente donde reside el Excel final y el portal web dinámico.

## 🚀 Guía de Ejecución
1.  Clonar el repositorio.
2.  Instalar dependencias: `pip install pandas openpyxl matplotlib`.
3.  Ejecutar el pipeline:
    ```bash
    python pipeline_valor_inventario_github/scripts/valor_inventario.py
    ```
4.  Consultar resultados en el directorio `output/` recién creado.

---
> **Nota de Privacidad:** Los datos en `data_samples/` han sido anonimizados y los valores numéricos alterados para proteger la confidencialidad de la información original, manteniendo intacta la lógica funcional y financiera del sistema.
