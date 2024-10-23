from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time

class WatcherHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and (event.src_path.endswith('.xls') or event.src_path.endswith('.xlsx')):
            print(f"Nuevo archivo detectado: {event.src_path}")
            contenido = leer_archivo_excel(event.src_path)
            if contenido:
                enviar_correo(contenido, "cacua666@hotmail.com")

def leer_archivo_excel(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo)
        contenido_texto = df.to_string(index=False)
        return contenido_texto
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return None

def enviar_correo(contenido, destinatario):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = os.getenv('SENDER_EMAIL')
    sender_password = os.getenv('SENDER_PASSWORD')
    
    if not sender_email or not sender_password:
        print("Error: Las credenciales de correo no están configuradas.")
        return

    subject = "Contenido del Archivo Excel"
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = destinatario
    msg['Subject'] = subject
    msg.attach(MIMEText(contenido, 'plain'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        print("Correo enviado correctamente")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")
    finally:
        server.quit()

class MyApp(App):
    def build(self):
        layout = BoxLayout(orientation='vertical')
        self.label = Label(text="Monitoreando...")
        layout.add_widget(self.label)

        button = Button(text='Iniciar Monitoreo')
        button.bind(on_press=self.start_monitoring)
        layout.add_widget(button)

        return layout

    def start_monitoring(self, instance):
        ruta_carpeta = "/storage/emulated/0/scanData"  # Cambia la ruta según necesites
        event_handler = WatcherHandler()
        observer = Observer()
        observer.schedule(event_handler, ruta_carpeta, recursive=False)
        observer.start()
        self.label.text = "Monitoreando: " + ruta_carpeta

if __name__ == "__main__":
    MyApp().run()
