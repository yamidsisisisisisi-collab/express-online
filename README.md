# Cartas y correos personalizados

Aplicación local para importar empresas y correos desde Excel y completar el nombre de cada empresa en una plantilla Word. Genera un ZIP de borradores individuales `.eml`, cada uno con su carta `.docx` adjunta, para que se revisen y se envíen manualmente desde el correo del usuario. Los contactos y la plantilla se guardan en `programa.sqlite3`, junto a este programa.

El nombre de la empresa se coloca después de «Señores» y en el campo «Empresa», en MAYÚSCULAS, negrita y Arial de 12 puntos.

## Iniciar en Windows

1. Instala Python 3.10 o posterior.
2. Abre PowerShell en esta carpeta y ejecuta:

   ```powershell
   py -m pip install -r requirements.txt
   py app.py
   ```

3. Abre `http://127.0.0.1:5000` en el navegador.
4. Carga la carta `.docx` y el Excel `.xlsx`. El Excel debe tener columnas llamadas `empresa` y `correo`. Importar otro Excel reemplaza la lista de contactos actual.
5. Selecciona las empresas y pulsa **Descargar correos seleccionados**.
6. Extrae el ZIP, abre cada archivo `.eml` con tu aplicación de correo, verifica el destinatario y la carta adjunta y pulsa **Enviar**. El programa nunca envía los correos por sí mismo.

## Envío manual

No necesitas Outlook clásico, ID ni permisos especiales para preparar los mensajes. El archivo `.eml` incluye el destinatario, el asunto, el texto y la carta adjunta. Si tu aplicación de correo no permite editar y enviar un `.eml` abierto, utiliza el `.docx` de la carpeta `Cartas` dentro del ZIP, copia el destinatario y el asunto del mensaje y adjunta la carta manualmente.

## Contenido del ZIP

- `Correos`: borradores `.eml`, uno por empresa, con la carta personalizada adjunta.
- `Cartas`: copias `.docx` individuales, disponibles para revisar o adjuntar manualmente.
- `LEEME.txt`: instrucciones breves para abrir y enviar los borradores.
- Para detener la aplicación, vuelve a PowerShell y presiona `Ctrl+C`.

La aplicación solo escucha en `127.0.0.1` y no se debe exponer a Internet.
