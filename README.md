# � Excel Sheet Unlocker (XLSX/XLSM)

Este proyecto permite **eliminar la protección de hojas y celdas** en archivos de Excel (`.xlsx` y `.xlsm`) cuando se ha olvidado la contraseña.  
� **Nota importante**: esto **no rompe contraseñas de apertura/cifrado**, solo elimina la restricción de edición de hojas protegidas.

---

## ✨ Características

- Funciona con **Excel moderno**: `.xlsx` y `.xlsm` (archivos con macros).
- Mantiene intactas fórmulas, datos y macros.
- No usa fuerza bruta → simplemente elimina la etiqueta de protección en el XML interno.
- Compatible con:
  - **Office 2016**
  - **Office 2019**
  - **Office 2021**
  - **Office 2024**
  - **Office 365 (Microsoft 365)**
- Compatible con **Windows, Linux y macOS** (requiere Python 3).

---

## � Requisitos

- Python 3.7+
- Librerías estándar (`zipfile`, `shutil`, `os`, `re`)

No requiere instalar paquetes externos.

---

## � Uso

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/tuusuario/excel-sheet-unlocker.git
   cd excel-sheet-unlocker
