# Conversor de presentaciones BdE: azul → blanco

**Autor:** Rubén Veiga Duarte ([ruben.veiga@bde.es](mailto:[ruben.veiga@bde.es])) - Mandame un correo si ves algún error o tienes alguna sugerencia.

**Descripción:** Esta herramienta convierte automáticamente presentaciones PowerPoint con fondo oscuro (azul) a fondo blanco, adaptando también los colores del texto, gráficos y bordes para que sean visibles sobre fondo claro.

**Última versión disponible:** [https://github.com/rubo88/slides_white](https://github.com/rubo88/slides_white)

**Fecha de última actualización:** 05/03/2026




---

## ¿Para qué sirve?

Convierte presentaciones con la **Plantilla azul BdE** al formato de **fondo blanco**, realizando los siguientes cambios automáticamente:

- Fondo de diapositivas, masters y layouts → blanco
- Texto blanco → negro
- Gráficos (títulos, ejes, leyendas, etiquetas) → negro
- Barras con patrón gris oscuro → negro
- Líneas blancas en series → negro
- Rellenos azules en masters y layouts → eliminados
- Colores del esquema de tema → corregidos para fondo claro

---

## Requisitos

- Tener **Python** instalado (versión 3.10 o superior).
- PPTs en formato BDE azul

---

## Cómo usarlo

1. **Copia** las presentaciones `.pptx` que quieras convertir en esta carpeta.
2. **Haz doble clic** en el archivo `run_fix_chart_colors.bat`.
3. La primera vez, el programa instalará las dependencias necesarias automáticamente en la propia carpeta (no se instala nada en tu Python).
4. El script procesará todos los `.pptx` de la carpeta y generará un nuevo archivo por cada uno con el sufijo `_white.pptx`.
5. Al terminar, pulsa cualquier tecla para cerrar la ventana.

> **Ejemplo:** `250305_MTBEyPUX.pptx` → `250305_MTBEyPUX_white.pptx`

---

## Notas

- Los archivos originales **no se modifican**. Siempre se genera un archivo nuevo con `_white` al final.
- Los archivos que ya terminen en `_white.pptx` se omiten automáticamente para evitar dobles conversiones.
