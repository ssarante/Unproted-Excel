import zipfile
import shutil
import os

def remove_sheet_protection(filepath, output):
    temp_dir = "temp_excel"

    # Extraer contenido del xlsm (que es un zip)
    with zipfile.ZipFile(filepath, 'r') as z:
        z.extractall(temp_dir)

    # Recorrer hojas y quitar etiquetas de protección
    worksheets = os.path.join(temp_dir, "xl", "worksheets")
    for file in os.listdir(worksheets):
        if file.endswith(".xml"):
            xml_path = os.path.join(worksheets, file)
            with open(xml_path, "r", encoding="utf-8") as f:
                content = f.read()
            # Borrar completamente la etiqueta sheetProtection
            import re
            content = re.sub(r'<sheetProtection[^>]*/>', '', content)
            with open(xml_path, "w", encoding="utf-8") as f:
                f.write(content)

    # Crear un nuevo .xlsm
    shutil.make_archive("unprotected", 'zip', temp_dir)
    shutil.move("unprotected.zip", output)

    # Limpiar
    shutil.rmtree(temp_dir)

if __name__ == "__main__":
    remove_sheet_protection("book1.xlsm", "book1_unprotected.xlsm")
    print("✅ Archivo generado: book1_unprotected.xlsm (sin bloqueo de hoja)")
