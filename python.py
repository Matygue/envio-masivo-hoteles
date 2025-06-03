import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import pywhatkit
import yagmail
import time
import requests
import csv
import io
from PIL import Image, ImageTk
import os
import json

CONFIG_FILE = "config.json"

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    else:
        return None

def guardar_config(email, clave):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"email": email, "clave": clave}, f)

# --- Configuración dinámica de email ---
config = cargar_config()

if config:
    EMAIL = config["email"]
    APP_PASSWORD = config["clave"]
else:
    ventana_login = ctk.CTk()
    ventana_login.title("Configuración Inicial")
    ventana_login.geometry("400x250")
    ventana_login.resizable(False, False)

    ctk.CTkLabel(ventana_login, text="Configurá tu email de Gmail", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=15)

    entry_email = ctk.CTkEntry(ventana_login, placeholder_text="Correo Gmail")
    entry_email.pack(pady=10)

    entry_clave = ctk.CTkEntry(ventana_login, placeholder_text="Contraseña de aplicación", show="*")
    entry_clave.pack(pady=10)

    def guardar_datos():
        global EMAIL, APP_PASSWORD
        EMAIL = entry_email.get()
        APP_PASSWORD = entry_clave.get()

        if EMAIL and APP_PASSWORD:
            guardar_config(EMAIL, APP_PASSWORD)
            ventana_login.destroy()
        else:
            messagebox.showwarning("Campos vacíos", "Completá los campos antes de continuar.")

    ctk.CTkButton(ventana_login, text="Guardar y continuar", command=guardar_datos).pack(pady=20)
    ventana_login.mainloop()

# --- Inicializar Yagmail ---
yag = yagmail.SMTP(EMAIL, APP_PASSWORD)

# --- Configuración Google Sheets ---
SHEET_ID = "171WOV_rKZNOkTGYxKypAKiJsOur4JALfqDWRZnwA4Tc"  # ← Reemplazá por tu ID
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv"

# --- Inicializar CustomTkinter ---
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")

# --- Verificación de clave ---
def verificar_clave(clave_ingresada):
    try:
        response = requests.get(SHEET_URL)
        if response.status_code == 200:
            contenido = response.content.decode("utf-8")
            reader = csv.reader(io.StringIO(contenido))
            next(reader)  # saltar encabezado
            claves = [fila[0].strip() for fila in reader]
            return clave_ingresada in claves
    except Exception as e:
        print(f"Error: {e}")
    return False

# --- Ventana de Login ---
def mostrar_login():
    login = ctk.CTk()
    login.title("Acceso HotelConecta")
    login.geometry("400x250")
    login.resizable(False, False)

    ctk.CTkLabel(login, text="Ingresá tu clave", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=30)
    entry_clave = ctk.CTkEntry(login, placeholder_text="Clave", show="*", width=250)
    entry_clave.pack(pady=10)

    def acceder():
        clave = entry_clave.get().strip()
        if verificar_clave(clave):
            login.destroy()
            lanzar_app()
        else:
            messagebox.showerror("Clave incorrecta", "La clave ingresada no es válida.")

    ctk.CTkButton(login, text="Ingresar", command=acceder).pack(pady=20)
    login.mainloop()

# --- Aplicación Principal ---
def lanzar_app():
    yag = yagmail.SMTP(EMAIL, APP_PASSWORD)
    app = ctk.CTk()
    app.title("HotelConecta – Captación Inteligente")
    app.geometry("700x700")
    app.resizable(False, False)

    df_contactos = None

    def cargar_excel():
        nonlocal df_contactos
        ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                df_contactos = pd.read_excel(ruta)
                label_archivo.configure(text=f"📄 Archivo cargado: {ruta.split('/')[-1]}", text_color="green")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")

    def enviar_whatsapp():
        if df_contactos is None:
            messagebox.showwarning("Error", "Primero cargá un archivo Excel.")
            return

        mensaje = txt_mensaje_wsp.get("1.0", "end").strip()
        if not mensaje:
            messagebox.showwarning("Error", "Ingresá un mensaje para enviar.")
            return

        for _, row in df_contactos.iterrows():
            try:
                nombre = row.get("Nombre", "amigo")
                telefono = str(row.get("Telefono", "")).strip()

                if not telefono.startswith("+54"):
                    telefono = "+54" + telefono.lstrip("0+ ")

                personalizado = mensaje.replace("{nombre}", nombre)
                pywhatkit.sendwhatmsg_instantly(telefono, personalizado, wait_time=15, tab_close=True)
                time.sleep(20)
            except Exception as e:
                print(f"❌ Error enviando a {nombre}: {e}")

        messagebox.showinfo("WhatsApp", "Mensajes enviados exitosamente.")

    def enviar_email():
        if df_contactos is None:
            messagebox.showwarning("Error", "Primero cargá un archivo Excel.")
            return

        mensaje = txt_mensaje_email.get("1.0", "end").strip()
        if not mensaje:
            messagebox.showwarning("Error", "Ingresá un mensaje para enviar.")
            return

        asunto = "¡Tenemos una propuesta para vos!"
        enviados, errores = 0, 0

        for _, row in df_contactos.iterrows():
            try:
                nombre = row.get("Nombre", "amigo")
                correo = row.get("Correo", "")
                if correo:
                    personalizado = mensaje.replace("{nombre}", nombre)
                    yag.send(to=correo, subject=asunto, contents=personalizado)
                    enviados += 1
                else:
                    errores += 1
            except Exception as e:
                print(f"❌ Error con {nombre}: {e}")
                errores += 1

        messagebox.showinfo("Email", f"Correos enviados: {enviados}\nErrores: {errores}")

    # --- Header ---
    header = ctk.CTkFrame(app, fg_color="#1E90FF", height=60, corner_radius=0)
    header.pack(fill="x")
    label_header = ctk.CTkLabel(header, text=" HotelConecta – Captación Inteligente", font=ctk.CTkFont(size=22, weight="bold"), text_color="white")
    label_header.pack(pady=15)

    # --- Frame Excel ---
    frame_excel = ctk.CTkFrame(app, corner_radius=12)
    frame_excel.pack(padx=20, pady=(20, 10), fill="x")
    btn_excel = ctk.CTkButton(frame_excel, text=" Cargar Excel", command=cargar_excel, width=180, height=40, fg_color="#1E90FF", hover_color="#3742fa")
    btn_excel.pack(side="left", padx=(10,0), pady=15)
    label_archivo = ctk.CTkLabel(frame_excel, text="Ningún archivo cargado", font=ctk.CTkFont(size=12))
    label_archivo.pack(side="left", padx=15)

    # --- Tabs WhatsApp y Email ---
    tabs = ctk.CTkTabview(app, width=650, height=400)
    tabs.pack(pady=10)
    tabs.add("WhatsApp")
    tabs.add("Email")

    # --- WhatsApp ---
    frame_wsp = tabs.tab("WhatsApp")
    ctk.CTkLabel(frame_wsp, text=" Mensaje de WhatsApp (usa {nombre})", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(10,5))
    txt_mensaje_wsp = ctk.CTkTextbox(frame_wsp, width=600, height=150, corner_radius=10)
    txt_mensaje_wsp.pack(padx=15)
    txt_mensaje_wsp.insert("1.0", "Hola {nombre}, ¿cómo estás? Te escribo para contarte algo que puede interesarte.")
    ctk.CTkButton(frame_wsp, text=" Enviar WhatsApps", command=enviar_whatsapp, width=200, height=45, fg_color="#25D366", hover_color="#1ebe57", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)

    # --- Email ---
    frame_email = tabs.tab("Email")
    ctk.CTkLabel(frame_email, text=" Mensaje de Email (usa {nombre})", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(10,5))
    txt_mensaje_email = ctk.CTkTextbox(frame_email, width=600, height=180, corner_radius=10)
    txt_mensaje_email.pack(padx=15)
    txt_mensaje_email.insert("1.0", "Hola {nombre},\n\nTenemos una propuesta interesante para vos.\n\nSaludos,\nMatías")
    ctk.CTkButton(frame_email, text=" Enviar Emails", command=enviar_email, width=200, height=45, fg_color="#0073e6", hover_color="#005bb5", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)

    # --- Footer ---
    ctk.CTkLabel(app, text="Versión 1.0  •  Desarrollado por Matías Guerrero", font=ctk.CTkFont(size=10), text_color="gray").pack(pady=10)

    # --- Lanzar app ---
    app.mainloop()

# --- Lanzar login primero ---
mostrar_login()
