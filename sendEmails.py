import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, ttk
import os
import re
import logging
import base64
import mimetypes

def embed_images_in_html(html_content, html_file_path):
    """Busca las imágenes en el HTML y las incrusta como base64"""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Buscar todas las etiquetas <img> que tienen un atributo src
        for img in soup.find_all('img'):
            img_src = img['src']
            if not img_src.startswith('data:'):  # Verificar que no sea ya base64
                # Obtener la ruta completa de la imagen referenciada
                img_path = os.path.join(os.path.dirname(html_file_path), img_src)
                if os.path.exists(img_path):
                    # Leer la imagen y convertirla a base64
                    with open(img_path, 'rb') as img_file:
                        img_data = img_file.read()
                        img_type, _ = mimetypes.guess_type(img_path)
                        img_base64 = base64.b64encode(img_data).decode('utf-8')
                        
                        # Crear el src de la imagen en base64
                        img['src'] = f"data:{img_type};base64,{img_base64}"
                else:
                    logging.warning(f"Imagen no encontrada: {img_path}")
        
        # Devolver el HTML con las imágenes incrustadas
        return str(soup)
    
    except Exception as e:
        logging.error(f"Error al incrustar imágenes en el HTML: {str(e)}")
        return html_content



# Configuración de logging para debug
logging.basicConfig(level=logging.DEBUG, 
                   format='%(asctime)s - %(levelname)s - %(message)s',
                   filename='email_sender.log')

def validate_email(email):
    """Validar formato de correo electrónico"""
    pattern = r"[^@]+@[^@]+\.[^@]+"
    return re.match(pattern, email)

def load_html_file(html_file):
    """Leer el archivo HTML, incrustar imágenes y devolver su contenido"""
    try:
        with open(html_file, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        # Incrustar las imágenes en el HTML
        html_content = embed_images_in_html(html_content, html_file)
        
        logging.info(f"Archivo HTML leído e imágenes incrustadas correctamente: {html_file}")
        return html_content
    except UnicodeDecodeError:
        # Si utf-8 falla, intentar con ISO-8859-1
        try:
            with open(html_file, 'r', encoding='iso-8859-1') as file:
                html_content = file.read()
            
            # Incrustar las imágenes en el HTML
            html_content = embed_images_in_html(html_content, html_file)
            
            logging.info(f"Archivo HTML leído correctamente con iso-8859-1 e imágenes incrustadas: {html_file}")
            return html_content
        except Exception as e:
            logging.error(f"Error leyendo archivo HTML con iso-8859-1: {str(e)}")
            messagebox.showerror("Error", "Error al leer el archivo HTML con iso-8859-1.")
            return None
    except Exception as e:
        logging.error(f"Error leyendo archivo HTML: {str(e)}")
        messagebox.showerror("Error", "Error al leer el archivo HTML.")
        return None


def send_emails(sender, password, rows, subject, html_message):
    """Enviar los correos electrónicos"""
    try:
        smtp = smtplib.SMTP('smtp-mail.outlook.com', port=587)
        smtp.starttls()
        smtp.login(sender, password)

        progress = ttk.Progressbar(window, orient='horizontal', length=300, mode='determinate')
        progress.grid(row=6, column=1)
        progress['maximum'] = len(rows)

        for idx, row in enumerate(rows):
            try:
                recipient_email = row[0]
                recipient_name = row[1]  # Nombre o correo en caso de estar vacío la columna B
                cc_emails = row[2:]  # Resto de las columnas son correos en copia (CC)

                if not validate_email(recipient_email):
                    logging.warning(f"Email inválido: {recipient_email}")
                    continue

                # Personalizar el mensaje reemplazando @user con el nombre o correo
                personalized_message = html_message.replace("@user", recipient_name)

                email = MIMEMultipart("alternative")
                email["From"] = sender
                email["To"] = recipient_email
                if cc_emails:
                    email["Cc"] = ", ".join(cc_emails)
                email["Subject"] = subject

                # Agregar contenido HTML
                email.attach(MIMEText(personalized_message, "html"))

                # Enviar email
                all_recipients = [recipient_email] + cc_emails
                smtp.sendmail(sender, all_recipients, email.as_string())
                logging.info(f"Email enviado a: {recipient_email}")
                
                # Actualizar barra de progreso
                progress['value'] = idx + 1
                window.update_idletasks()
                
            except Exception as e:
                logging.error(f"Error enviando email a {recipient_email}: {str(e)}")
                continue

        smtp.quit()
        messagebox.showinfo("Éxito", "Correos enviados exitosamente.")
    except Exception as e:
        logging.error(f"Error en send_emails: {str(e)}")
        messagebox.showerror("Error", f"Error al enviar correos: {str(e)}")


def load_rows(excel_file):
    """Cargar destinatarios desde el archivo Excel"""
    try:
        df = pd.read_excel(excel_file, header=None)
        clean_rows = []

        for index, row in df.iterrows():
            recipient_email = str(row[0]).strip()  # Columna A - Correo de destino
            recipient_name = str(row[1]).strip() if pd.notna(row[1]) else recipient_email  # Columna B - Nombre o correo
            cc_emails = [str(email).strip() for email in row[2:] if pd.notna(email)]  # Columna C en adelante - Copias

            # Agregar a la lista el correo, nombre y las copias
            clean_rows.append([recipient_email, recipient_name] + cc_emails)

        logging.info(f"Filas leídas del Excel: {len(clean_rows)}")
        return clean_rows
    except Exception as e:
        logging.error(f"Error en load_rows: {str(e)}")
        raise


def select_excel():
    """Seleccionar archivo Excel"""
    try:
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file:
            excel_entry.delete(0, "end")
            excel_entry.insert(0, file)
    except Exception as e:
        logging.error(f"Error en select_excel: {str(e)}")
        messagebox.showerror("Error", "Error al seleccionar archivo Excel")

def select_html():
    """Seleccionar archivo HTML"""
    try:
        file = filedialog.askopenfilename(filetypes=[("HTML Files", "*.htm;*.html")])
        if file:
            html_entry.delete(0, "end")
            html_entry.insert(0, file)
    except Exception as e:
        logging.error(f"Error en select_html: {str(e)}")
        messagebox.showerror("Error", "Error al seleccionar archivo HTML")

def start_send():
    """Iniciar el envío de correos"""
    try:
        sender = email_entry.get()
        password = password_entry.get()
        excel_file = excel_entry.get()
        html_file = html_entry.get()
        subject = subject_entry.get()

        if not all([sender, password, excel_file, html_file, subject]):
            messagebox.showwarning("Warning", "Por favor completa todos los campos.")
            return

        rows = load_rows(excel_file)
        if not rows:
            messagebox.showwarning("Warning", "No se encontraron destinatarios en el archivo Excel.")
            return

        html_message = load_html_file(html_file)
        if html_message is None:
            return

        send_emails(sender, password, rows, subject, html_message)
        
    except Exception as e:
        logging.error(f"Error en start_send: {str(e)}")
        messagebox.showerror("Error", f"Error al iniciar el envío: {str(e)}")

def confirm_send():
    """Confirmar antes de enviar correos"""
    confirmation = messagebox.askyesno("Confirmar Envío", "¿Estás seguro de que quieres enviar los correos?")
    if confirmation:
        start_send()

# Configuración de la interfaz gráfica
window = Tk()
window.title("Mass Email Sender")
window.geometry("510x470")

# Crear elementos de la interfaz
Label(window, text="Correo:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
email_entry = Entry(window, width=40)
email_entry.grid(row=0, column=1)

Label(window, text="Contraseña:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
password_entry = Entry(window, width=40, show="*")
password_entry.grid(row=1, column=1)

Label(window, text="Archivo Excel (destinatarios):").grid(row=2, column=0, padx=10, pady=10, sticky="e")
excel_entry = Entry(window, width=30)
excel_entry.grid(row=2, column=1, sticky="w")
Button(window, text="Buscar", command=select_excel).grid(row=2, column=2, padx=10)

Label(window, text="Archivo HTML (mensaje):").grid(row=3, column=0, padx=10, pady=10, sticky="e")
html_entry = Entry(window, width=30)
html_entry.grid(row=3, column=1, sticky="w")
Button(window, text="Buscar", command=select_html).grid(row=3, column=2, padx=10)

Label(window, text="Asunto del Correo:").grid(row=4, column=0, padx=10, pady=10, sticky="e")
subject_entry = Entry(window, width=40)
subject_entry.grid(row=4, column=1)

Button(window, text="Enviar Correos", command=confirm_send).grid(row=5, column=1, pady=20)

Label(window, text="Esta no es una aplicación oficial de Pacífico Seguros,\n"
                   "Es un software interno del equipo End User.", 
      fg="red", wraplength=400, justify="center").grid(row=7, column=0, columnspan=3, padx=10, pady=10)

window.mainloop()
