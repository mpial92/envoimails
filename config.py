import os

# Configuration email - utilise les variables d'environnement si disponibles, sinon valeurs par défaut
EMAIL_CONFIG = {
    'username': os.getenv('EMAIL_USERNAME', 'contact@lecercledesseniorsneuflizeobc.fr'),
    'password': os.getenv('EMAIL_PASSWORD', 'Cer2021##'),
    'smtp_server': os.getenv('SMTP_SERVER', 'ssl0.ovh.net'),
    'smtp_port': int(os.getenv('SMTP_PORT', '587'))
}

def get_email_config():
    """Retourne la configuration email"""
    return EMAIL_CONFIG

def validate_config():
    """Valide que la configuration est correcte"""
    config = get_email_config()
    
    if not config['password'] or config['password'].strip() == '':
        print("ERREUR: Mot de passe email manquant")
        return False
        
    if not config['username'] or '@' not in config['username']:
        print("ERREUR: Adresse email invalide")
        return False
        
    return True