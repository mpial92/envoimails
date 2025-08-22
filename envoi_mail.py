import openpyxl
import smtplib
import os
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Charger le fichier Excel
chemin_fichier_excel = 'liste.xlsx'
wb = openpyxl.load_workbook(chemin_fichier_excel)
sheet = wb.active

# Paramètres du compte et du serveur SMTP
adresse_email = "contact@lecercledesseniorsneuflizeobc.fr"
mot_de_passe = "Cerclenew2021##"
serveur_smtp = "ssl0.ovh.net"
port_smtp = 587

# Connexion au serveur SMTP
server = smtplib.SMTP(serveur_smtp, port_smtp)
server.starttls()
server.login(adresse_email, mot_de_passe)

compteur_emails = 0

# Fonction pour se connecter au serveur SMTP
def connecter_au_serveur_smtp():
    server = smtplib.SMTP(serveur_smtp, port_smtp)
    server.starttls()
    server.login(adresse_email, mot_de_passe)
    return server

# Connexion initiale au serveur SMTP
server = connecter_au_serveur_smtp()

# Parcourir les lignes du fichier Excel
#print(sheet.max_row + 1)
for row in range(2, sheet.max_row + 1):
    destinataire = sheet.cell(row=row, column=12).value
    nom = sheet.cell(row=row, column=2).value

    # Créer le message
    msg = MIMEMultipart()
    msg['From'] = adresse_email
    msg['To'] = destinataire
    msg['Subject'] = "Triste nouvelle"

    # Corps du message
    body = """
<html>
  <body>
    <p>Bonjour à toutes et à tous,</p>

  </body>
</html>
"""
 
    msg.attach(MIMEText(body, 'html'))

    # Envoyer les emails 
    try:
        #server.sendmail(adresse_email, destinataire, msg.as_string())
        print(f"E-mail envoyé à {destinataire}")
        compteur_emails += 1 

        # Pause toutes les 10 e-mails
        if compteur_emails % 10 == 0:
            print("Pause de 1 minutes pour éviter les filtres anti-spam...")
            time.sleep(60)  # Pause de 1 minutes
            server = connecter_au_serveur_smtp()  # Se reconnecter au serveur SMTP
    
    except Exception as e:
        print(f"Erreur lors de l'envoi de l'e-mail à {destinataire} : {e}")

# Déconnexion du serveur SMTP
server.quit()

