# HotelConecta

Aplicación para enviar mensajes personalizados por WhatsApp y Email a contactos cargados desde un Excel.

## Funcionalidades

- Carga de contactos desde Excel.
- Envío de WhatsApp personalizado con `pywhatkit`.
- Envío de Emails usando `yagmail`.
- Interfaz gráfica con `customtkinter`.

## Requisitos

- Python 3.10+
- Paquetes: `customtkinter`, `pandas`, `pywhatkit`, `yagmail`, `Pillow`, etc.

## Uso

1. Ejecutar el script principal.
2. Configurar email de Gmail y clave de aplicación.
3. Iniciar sesión con la clave almacenada en Google Sheets.
4. Cargar un Excel con columnas `Nombre`, `Telefono`, y `Correo`.
5. Enviar mensajes desde las pestañas de WhatsApp o Email.
