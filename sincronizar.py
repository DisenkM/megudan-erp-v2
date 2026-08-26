import os
import json
import subprocess
from google.oauth2 import service_account
from googleapiclient.discovery import build

def main():
    # 1. Fusionar contenido de los archivos legibles
    consolidado = "=== CONSOLIDADO COMPLETO DEL REPOSITORIO ===\n"
    
    # CORREGIDO: Buscamos de forma directa en el directorio actual sin rompernos con rutas del servidor
    for root, dirs, files in os.walk("."):
        # Filtrar de forma segura para ignorar carpetas ocultas del sistema de GitHub (.git, .github)
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        
        for file in files:
            # Evitar leer archivos de configuración y el propio script de Python
            if file in ["sincronizar.py", "package-lock.json", "package.json"]:
                continue
                
            filepath = os.path.join(root, file)
            # Limpiamos la ruta estética para el documento final (ej: ./src/codigo.gs)
            clean_path = os.path.relpath(filepath, ".")
            
            try:
                # Validar si el archivo es de texto legible o códigos .gs / .js / .md
                result = subprocess.run(["file", filepath], capture_output=True, text=True)
                es_texto = "text" in result.stdout or "empty" in result.stdout or "JSON" in result.stdout
                es_codigo = file.endswith(('.gs', '.js', '.md', '.txt', '.json', '.html'))
                
                if es_texto or es_codigo:
                    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                        contenido = f.read().strip()
                        # Solo agregamos el archivo si realmente tiene código o letras adentro
                        if contenido:
                            consolidado += f"\n\n=========================================\n"
                            consolidado += f" RUTA DEL ARCHIVO: ./{clean_path}\n"
                            consolidado += f"=========================================\n\n"
                            consolidado += contenido + "\n"
            except Exception as e:
                print(f"No se pudo leer el archivo {clean_path}: {str(e)}")
                continue

    # 2. Autenticarse en Google Drive y Docs
    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = service_account.Credentials.from_service_account_info(creds_dict)
    drive_service = build("drive", "v3", credentials=creds)
    docs_service = build("docs", "v1", credentials=creds)

    # 3. Buscar el Google Doc existente creado por el usuario en la carpeta
    folder_id = os.environ["FOLDER_ID"]
    doc_name = os.environ["DOCUMENT_NAME"]
    
    query = f"name = '{doc_name}' and '{folder_id}' in parents and mimeType = 'application/vnd.google-apps.document' and trashed = false"
    results = drive_service.files().list(q=query, spaces="drive", fields="files(id)").execute()
    items = results.get("files", [])

    if not items:
        print(f"Error: No se encontró el documento de Google Docs llamado '{doc_name}' en la carpeta.")
        exit(1)
        
    doc_id = items[0]["id"]
    print(f"Documento encontrado con ID: {doc_id}")

    # 4. Vaciar el contenido viejo e inyectar el nuevo consolidado
    doc = docs_service.documents().get(documentId=doc_id).execute()
    end_index = doc.get("body").get("content")[-1].get("endIndex")

    requests = []
    # Si el documento tiene texto previo, lo borramos (dejando el carácter mínimo 1)
    if end_index > 2:
        requests.append({"deleteContentRange": {"range": {"startIndex": 1, "endIndex": end_index - 1}}})
        
    # Insertamos todo el repositorio unificado
    requests.append({"insertText": {"location": {"index": 1}, "text": consolidado}})

    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
    print("¡Sincronización completada con éxito!")

if __name__ == "__main__":
    main()
