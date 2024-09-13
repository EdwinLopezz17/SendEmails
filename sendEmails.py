import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import mammoth
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox
from docx import Document
import os
from tkinter import messagebox

def confirm_send():
    confirmation = messagebox.askyesno("Confirmar Envío", "¿Estás segura de que quieres enviar correos?")
    if confirmation:
        start_send()

def read_html_message(word_file):
    with open(word_file, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html_message = result.value 
    return html_message

def extract_images_from_docx(word_file):
    doc = Document(word_file)
    images = {}
    image_counter = 0
    image_folder = os.path.dirname(word_file)

    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_counter += 1
            image_filename = f"image_{image_counter}.png"
            image_path = os.path.join(image_folder, image_filename)
            with open(image_path, "wb") as img_file:
                img_file.write(rel.target_part.blob)
            images[f"image_{image_counter}"] = image_path

    return images

def send_emails(sender, password, rows, subject, html_message, images):
    try:
        smtp = smtplib.SMTP('smtp-mail.outlook.com', port=587)
        smtp.starttls()
        smtp.login(sender, password)

        for row in rows:
            recipient = row[0]
            cc = row[1:]  

            personalized_message = html_message.replace("@user", recipient)

            email = MIMEMultipart("related") 
            email["From"] = sender
            email["To"] = recipient
            if cc:
                email["Cc"] = ", ".join(cc)
            email["Subject"] = subject

            email_alternative = MIMEMultipart("alternative")
            email.attach(email_alternative)

            email_alternative.attach(MIMEText(personalized_message, "html"))

            for cid, image_path in images.items():
                with open(image_path, "rb") as img_file:
                    img = MIMEImage(img_file.read())
                    img.add_header("Content-ID", f"<{cid}>")
                    email.attach(img)

            smtp.sendmail(sender, [recipient] + cc, email.as_string())
            print("Email sent to: " + recipient)

        smtp.quit() 
        messagebox.showinfo("Success", "Emails sent successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error sending emails: {str(e)}")

def select_excel():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file:
        excel_entry.delete(0, "end")
        excel_entry.insert(0, file)

def select_word():
    file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if file:
        word_entry.delete(0, "end")
        word_entry.insert(0, file)

def start_send():
    sender = email_entry.get()
    password = password_entry.get()
    excel_file = excel_entry.get()
    word_file = word_entry.get()
    subject = subject_entry.get()

    if not sender or not password or not excel_file or not word_file or not subject:
        messagebox.showwarning("Warning", "Please fill in all fields.")
        return

    rows = load_rows(excel_file)

    html_message = read_html_message(word_file)
    images = extract_images_from_docx(word_file) 

    send_emails(sender, password, rows, subject, html_message, images)

def load_rows(excel_file):
    df = pd.read_excel(excel_file, header=None)
    clean_rows = []

    for index, row in df.iterrows():
        row_list = row.tolist()
        clean_row = [str(item).strip() for item in row_list if pd.notna(item)]
        if clean_row:
            clean_rows.append(clean_row)

    print("Rows read from Excel:", clean_rows)
    return clean_rows

# Graphical Interface
window = Tk()
window.title("Mass Email Sender")
window.geometry("510x470")  

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

Label(window, text="Archivo  Word (mensaje):").grid(row=3, column=0, padx=10, pady=10, sticky="e")
word_entry = Entry(window, width=30)
word_entry.grid(row=3, column=1, sticky="w")
Button(window, text="Buscar", command=select_word).grid(row=3, column=2, padx=10)

Label(window, text="Asunto del Correo:").grid(row=4, column=0, padx=10, pady=10, sticky="e")
subject_entry = Entry(window, width=40)
subject_entry.grid(row=4, column=1)

Button(window, text="Enviar Correos", command=confirm_send).grid(row=5, column=1, pady=20)

Label(window, text="Esta no es una aplicación oficial de Pacífico Seguros,\n"
                    "Es un software interno del equipo.", 
      fg="red", wraplength=400, justify="center").grid(row=7, column=0, columnspan=3, padx=10, pady=10)


window.mainloop()
