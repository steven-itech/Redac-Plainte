import ctypes
import pyttsx3
import smtplib
import re
import tkinter as tk
from fpdf import FPDF
from tkinter import messagebox, Menu
from tkcalendar import Calendar
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

engine = pyttsx3.init()
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)

def tts(message):
    
    engine.say(message)
    engine.runAndWait()

def send_mail():
    
    email = email_entry.get()
    password = password_entry.get()
    recipient = recipient_entry.get()
    message = message_text.get("1.0", tk.END).strip()

    if not all([email, password, recipient, message]):
        
        mail_info = "Veuillez remplir tous les champs pour envoyer le courriel au destinataire !"
        
        tts(mail_info)
        messagebox.showwarning(title="Rédac'Plainte :", message=mail_info)
        
        return

    if not email.endswith("@gmail.com"):
        
        email_info = "Veuillez entrer uniquement une adresse électronique Google !"
        
        tts(email_info)
        messagebox.showwarning(title="Rédac'Plainte :", message=email_info)
        
        return

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email, password)

        msg = MIMEMultipart()
        msg["From"] = email
        msg["Subject"] = "Courriel de la part d'un officier de la Police ou Gendarmerie nationale."
        msg["To"] = recipient
        msg.attach(MIMEText(message, "plain"))

        server.sendmail(email, recipient, msg.as_string())
        server.quit()

        send_mail_info = "Votre courriel vient d'être envoyé !"
        tts(send_mail_info)
        messagebox.showinfo(title="Rédac'Plainte :", message=send_mail_info)

        email_entry.delete(0, tk.END)
        password_entry.delete(0, tk.END)
        recipient_entry.delete(0, tk.END)
        message_text.delete("1.0", tk.END)

    except Exception:
        
        exception_info = "Une erreur est survenue lors de l'envoi du courriel !"
        
        tts(exception_info)
        messagebox.showwarning(title="Rédac'Plainte :", message=exception_info)

def mail_window():
    
    mail_win = tk.Toplevel(window)
    mail_win.title("Envoyer un courriel :")
    mail_win.iconbitmap("mail.ico")
    mail_win.geometry("800x645")
    mail_win.resizable(False, False)

    tk.Label(mail_win, text="Quelle est votre adresse électronique ?", font=("Arial", 12)).pack(pady=10)
    global email_entry
    email_entry = tk.Entry(mail_win, font=("Arial", 12), width=40)
    email_entry.pack(pady=10)

    tk.Label(mail_win, text="Quel est votre mot de passe d'application ?", font=("Arial", 12)).pack(pady=10)
    global password_entry
    password_entry = tk.Entry(mail_win, font=("Arial", 12), show="*", width=40)
    password_entry.pack(pady=5)

    tk.Label(mail_win, text="Quelle est l'adresse électronique du destinataire ?", font=("Arial", 12)).pack(pady=10)
    global recipient_entry
    recipient_entry = tk.Entry(mail_win, font=("Arial", 12), width=40)
    recipient_entry.pack(pady=5)

    tk.Label(mail_win, text="Quel est votre message ?", font=("Arial", 12)).pack(pady=10)
    global message_text
    message_text = tk.Text(mail_win, font=("Arial", 12), height=15, width=60)
    message_text.pack(pady=5)

    send_button = tk.Button(mail_win, text="Envoyer !", font=("Arial", 12), command=send_mail)
    send_button.pack(pady=15)

def submit():
    
    victim_entry = victim.get()
    accused_entry = accused.get()
    events_date_entry = events.get()
    events_entry = description.get("1.0", tk.END).strip()
    date = calendar.get_date()
    officer_entry = officer.get()

    events_entry = re.sub(r'\.\s+', ". \n", events_entry)

    if not all([victim_entry, accused_entry, events_date_entry, events_entry, officer_entry]):
       
        warning = "Veuillez remplir l'ensemble des champs, afin de pouvoir créer la plainte !"
        
        tts(warning)
        messagebox.showwarning(title="Rédac'Plainte :", message=warning)
        
        return

    if not all(any(c.isalpha() for c in entry) and any(c.isspace() for c in entry) for entry in [victim_entry, accused_entry, officer_entry]):
        
        warning = "Veuillez transmettre un prénom et un nom de famille, suivi d'un espace entre ceux-ci !"
        
        tts(warning)
        messagebox.showwarning(title="Rédac'Plainte :", message=warning)
        
        return

    pdf = FPDF()
    pdf.add_page()
    
    pdf.add_font("DejaVuSans", "", "DejaVuSans.ttf", uni=True)
    pdf.add_font("DejaVuSans", "B", "DejaVuSans-Bold.ttf", uni=True)
    
    pdf.set_font("DejaVuSans", size=12)

    pdf.image("ministère.png", x=10, y=10, w=90, h=60)
    pdf.image("nationale.jpg", x=pdf.w - 90 - 10, y=10, w=90, h=60)

    pdf.ln(75)
    
    pdf.set_font("DejaVuSans", "B", 12)
    
    pdf.cell(200, 10, txt="PROCÈS-VERBAL :", ln=True, align='C')

    pdf.ln(3)

    pdf.set_font("DejaVuSans", size=12)

    pdf.cell(0, 10, txt=f"Date de la plainte : {date}", ln=True, align='L')
    pdf.cell(0, 10, txt=f"Victime : {victim_entry}", ln=True, align='L')
    pdf.cell(0, 10, txt=f"Accusé : {accused_entry}", ln=True, align='L')
    pdf.cell(0, 10, txt=f"Date des événements : {events_date_entry}", ln=True, align='L')
    pdf.cell(0, 10, txt=f"Officier de police : {officer_entry}", ln=True, align='L')

    pdf.ln(3)

    pdf.multi_cell(0, 10, txt=f"Description des événements :\n{events_entry}", align='L')

    pdf_name = f"Plainte de {victim_entry}.pdf"
    pdf.output(pdf_name)

    confirmation = f"La plainte a été créée et enregistrée sous le nom de : {pdf_name} !"
    tts(confirmation)
    messagebox.showinfo(title="Rédac'Plainte :", message=confirmation)


if __name__ == "__main__":
    
    ctypes.windll.kernel32.SetConsoleTitleW("Rédac'Plainte :")

    window = tk.Tk()
    window.title("Rédac'Plainte :")
    window.iconbitmap("police.ico")
    window.geometry("880x530")
    window.resizable(False, False)

    tk.Label(window, text="Quel est le prénom et le nom de la victime ?", font=("Arial", 12)).place(x=5, y=10)
    victim = tk.Entry(window, font=("Arial", 12), width=33)
    victim.place(x=10, y=35)

    tk.Label(window, text="Quel est le prénom et nom de l'accusé ?", font=("Arial", 12)).place(x=5, y=65)
    accused = tk.Entry(window, font=("Arial", 12), width=30)
    accused.place(x=10, y=90)

    tk.Label(window, text="Quand les événements ont-ils eu lieu ?", font=("Arial", 12)).place(x=5, y=125)
    events = tk.Entry(window, font=("Arial", 12), width=29)
    events.place(x=10, y=150)

    tk.Label(window, text="Officier de police :", font=("Arial", 12)).place(x=5, y=185)
    officer = tk.Entry(window, font=("Arial", 12), width=29)
    officer.place(x=10, y=210)

    tk.Label(window, text="Description des événements :", font=("Arial", 12)).place(x=5, y=245)
    description = tk.Text(window, font=("Arial", 12), height=10, width=65)
    description.place(x=10, y=270)

    submit_button = tk.Button(window, text="Soumettre !", font=("Arial", 12), command=submit)
    submit_button.place(x=10, y=480, width=150)

    calendar = Calendar(window, selectmode="day", locale='fr_FR', font=("Arial", 12))
    calendar.place(x=500, y=21)

    menu = Menu(window)
    window.config(menu=menu)

    mail_menu = Menu(menu, tearoff=0)
    menu.add_cascade(label="💬", menu=mail_menu)
    mail_menu.add_command(label="Envoyer un courriel.", command=mail_window)

    window.mainloop()