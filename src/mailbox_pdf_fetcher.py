import os
import re
import sys
import base64
import requests
import mimetypes
from lst_fourni import FOURNISSEURS
from datetime import datetime, timedelta, timezone, time

DOWNLOAD_FOLDER = os.getenv("DOWNLOAD_FOLDER")
if not DOWNLOAD_FOLDER:
    raise ValueError("DOWNLOAD_FOLDER n'est pas définie dans les variables d'environnement.")

TENANT_ID = os.getenv("GRAPH_TENANT_ID")
if not TENANT_ID:
    raise ValueError("TENANT_ID n'est pas définie dans les variables d'environnement.")

CLIENT_ID = os.getenv("GRAPH_CLIENT_ID")
if not CLIENT_ID:
    raise ValueError("CLIENT_ID n'est pas définie dans les variables d'environnement.")

CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")
if not CLIENT_SECRET:
    raise ValueError("CLIENT_SECRET n'est pas définie dans les variables d'environnement.")

FACT_EMAIL = os.getenv("FACT_EMAIL")
if not FACT_EMAIL:
    raise ValueError("FACT_EMAIL n'est pas définie dans les variables d'environnement.")

PROCESS_TRACKER = os.getenv("PROCESS_TRACKER")
if not PROCESS_TRACKER:
    raise ValueError("PROCESS_TRACKER n'est pas définie dans les variables d'environnement.")


# Récupération de la liste des fournisseurs
pattern_liste = "|".join(re.escape(nom) for nom in FOURNISSEURS)
entreprise_regex = re.compile(rf"\b({pattern_liste})\b", re.IGNORECASE)

# Fonction pour éditer un fichier .log horodaté
def write_log(LOG_FILE, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{timestamp} - {message}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# Récupération des factures depuis la boîte mail dédiée

def find_fourni(sender):
    # On initialise cleaned_sender avec la valeur brute du sender par défaut
    cleaned_sender = sender.upper() 

    match_email = re.search(r"([^\s:@]+)@([^\s:@]+)", sender)
    if match_email:
        user = match_email.group(1).lower().strip()
        full_domain = match_email.group(2).lower().strip()
        domain = full_domain.split('.')[0] 

        if "by.sendoc" in full_domain:
            domain = "by.sendoc"

        # Cas des mails génériques : on extrait le USER
        if domain in ("gmail", "yahoo", "orange", "gmx", "outlook", "hotmail", "live", "wanadoo", "icloud", "free", "by.sendoc", "mypbconnect"):
            if "menuiserie2bm" in user: return "2BM MENUISERIE"
            if "camara1933" in user: return "CAMARA MENUISERIE"
            if "gbacot" in user : return None
            
            # Ici, si user est "leuco_facturation", cleaned_sender devient "LEUCO_FACTURATION"
            cleaned_sender = user.upper()

        # Cas des domaines spécifiques : on extrait le DOMAIN
        else:
            if domain == "ebene-tradition": return None
            if domain == "wood-en": return None
            if domain == "cuiraucarre":
                return "HCA INVEST" if "HCA INVEST" in sender.upper() else "CUIR AU CARRE"
            if domain == "comec": return "COMEC INDUSTRIE"
            if "legallais" in full_domain: return "LEGALLAIS"
            if domain == "dailybiz" : return "THOURY AFFUTAGE"
            
            cleaned_sender = domain.upper()

    # On nettoie les tirets/underscores 
    search_zone = cleaned_sender.replace("_", " ").replace("-", " ")

    # Recherche par mots-clés dans la zone nettoyée
    match_ent = entreprise_regex.search(search_zone)
    if match_ent:
        return match_ent.group(0).upper()
    
    # Si la regex ne trouve rien, on renvoie quand même le nom extrait
    return search_zone

# Fonction de récupération du jeton d'accès 
def get_access_token():
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    # Format du corps de la requête pour obtenir un jeton d'accès Microsoft Graph
    token_data = {
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET
    }

    # Indique que les données du POST sont encodées en x-www-form-urlencoded, format requis par l'endpoint token d'Azure AD
    token_headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    # Récupération de la réponse renvoyée par la requête 
    response = requests.post(token_url, data=token_data, headers=token_headers)

    # Si le statut de la réponse est différent de "ok" -> écriture de l'erreur dans le fichier de suivi et sortie du programme 
    if response.status_code != 200:
        write_log(PROCESS_TRACKER, f"Erreur récupération token : {response.json()}")
        sys.exit(1)
    # Récupération du jeton d'accès 
    access_token = response.json().get("access_token")
    # S'il n'y a pas de jeton d'accès -> écriture de l'erreur dans le fichier de suivi et sortie du programme 
    if not access_token:
        write_log(PROCESS_TRACKER, f"Pas de jeton d'accès dans la réponse : {response.json()}")
        sys.exit(1)
    return access_token

# Fonction qui télécharge tous les PDF rencontrés
def download_attachments():

    access_token = get_access_token()
    # Création d'un en-tête HTTP d'authentification au format standard pour envoyer un jeton d'accès dans une requête HTTP
    auth_headers = {"Authorization": f"Bearer {access_token}"}
    write_log(PROCESS_TRACKER, "Jeton obtenu avec succès.")
    # On convertit en UTC (cela va retirer 1h ou 2h selon l'heure d'été/hiver)
    today_utc = (datetime.now()
             .replace(hour=0, minute=0, second=0, microsecond=0)
             .astimezone(timezone.utc))
    today = today_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    # Récuèpre les mails du jour 
    # Attention time UTC a 1h en moins
   # today = (datetime.now(timezone.utc)).replace(hour=0, minute=0, second=0, microsecond=0).strftime("%Y-%m-%dT%H:%M:%SZ")
   # now = datetime.now(timezone.utc)
   # two_days_ago = now - timedelta(days=1)
   # two_days_ago_str = two_days_ago.strftime("%Y-%m-%dT%H:%M:%SZ")

    graph_url = (f"https://graph.microsoft.com/v1.0/users/{FACT_EMAIL}/mailFolders/Inbox/messages"
                f"?$top=50&$filter=receivedDateTime ge {today}")
    
    messages = []

    # Récupération des mails récents avec pagination
        # évite un timeout si on a trop de mails à traiter
        # permet au serveur de répondre rapidement et de réduire la charge réseau
        # contrôle de flux pas pages
    while graph_url:
        try:
            # Envoie une requête GET à l'URL de l'API avec les en-têtes d'authentification
            resp = requests.get(graph_url, headers=auth_headers)
            # Lève une exception si la requête échoue
            resp.raise_for_status()
        except Exception as e:
            write_log(PROCESS_TRACKER, f"Erreur lors de la récupération des mails : {e}")
            sys.exit(1)
        # Convertit la réponse HTTP en format JSON pour pouvoir l'utiliser
        data = resp.json()
        # Ajoute les mails retournés par l'API à la liste "messages", "value" contient la liste des mails
        messages.extend(data.get("value", []))
        # Met à jour "graph_url" avec l'URL de la page suivante (si elle existe). On utilise "@odata.nextLink" pour la pagination
        graph_url = data.get("@odata.nextLink")

    write_log(PROCESS_TRACKER, f"{len(messages)} message(s) récupéré(s) dans la boîte de {FACT_EMAIL}")
    downloaded_files_info = []
    # Téléchargement les PDFs uniquement 
    for msg in messages:
        # Récupération des infos de chaque mail : l'objet et l'expéditeur
        msg_id = msg["id"]
        subject = msg.get("subject", "NoSubject")
        if subject :
            subject = subject.replace("/", "_").replace("\\", "_")
        else :
            subject = "No Subject"
        sender = msg.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
        
        write_log(PROCESS_TRACKER, f"Traitement du mail : {subject} | De : {sender}")

        # Vérifie si le message a des pièces jointes
        if msg.get("hasAttachments", False):
            # Construit l'URL pour récupérer les pièces jointes du message via l'API Microsoft Graph
            attach_url = f"https://graph.microsoft.com/v1.0/users/{FACT_EMAIL}/messages/{msg_id}/attachments"
            try:
                attach_resp = requests.get(attach_url, headers=auth_headers)
                attach_resp.raise_for_status()
            except Exception as e:
                write_log(PROCESS_TRACKER, f"Erreur récupération pièces jointes : {e}")
                continue
            
            attachments = attach_resp.json().get("value", [])

            sender_name = find_fourni(sender)
            if not sender_name:
                write_log(PROCESS_TRACKER, f"Fournisseur non identifié pour : {sender}")
                sender_name = "None"

            for att in attachments:
                # Vérifie si la pièce jointe est de type "fichier" (et non un lien ou un autre type)
                if att.get("@odata.type") != "#microsoft.graph.fileAttachment":
                    write_log(PROCESS_TRACKER ,f"Pièce jointe ignorée (type non supporté) : {att.get('@odata.type')}")
                    continue
                
                # Récupère le nom du fichier
                filename = att.get("name", "fichier_sans_nom")
                # Sécurité : on remplace les caractères interdits dans le nom d'origine
                filename = re.sub(r'[\\/*?:"<>|]', "_", filename)
                # Récupère le contenu binaire de la pièce jointe (encodé en base64)
                content_bytes = att.get("contentBytes")
                # Si le contenu est vide -> écriture de l'erreur dans le fichier de suivi et passe à la pièce jointe suivante
                if not content_bytes:
                    write_log(PROCESS_TRACKER, f"Pièce jointe vide : {filename}")
                    continue
                
                file_data = base64.b64decode(content_bytes)

                # Détermine le type MIME du fichier à partir de son nom
                mime_type, _ = mimetypes.guess_type(filename)
                # Vérifie si le fichier est un PDF via le type MIME
                if mime_type != "application/pdf":
                    # Si le type MIME n'est pas PDF, on vérifie les premiers octets pour confirmer
                    if not file_data.startswith(b"%PDF"):
                        write_log(PROCESS_TRACKER, f"Pièce jointe ignorée (non PDF) : {filename}")
                        continue
                    if re.search(r"CGV", filename, re.IGNORECASE) or re.search(r"Coniditions générales de ventes", filename, re.IGNORECASE):
                        write_log(PROCESS_TRACKER, f"Pièce jointe ignorée (CGV) : {filename}")
                        continue
                    # Si le fichier semble être un PDF mais n'a pas l'extension .pdf, on force son ajout 
                    if not filename.lower().endswith(".pdf"):
                        filename += ".pdf"

                # Gestion des doublons
                base_name = os.path.splitext(filename)[0]
                extension = ".pdf"
                final_filename = filename
                filepath = os.path.join(DOWNLOAD_FOLDER, final_filename)
                counter = 1
                while os.path.exists(filepath):
                    # On crée un nom type : "0326 - FOURNISSEUR - None (1).pdf"
                    final_filename = f"{base_name} ({counter}){extension}"
                    filepath = os.path.join(DOWNLOAD_FOLDER, final_filename)
                    counter += 1

                # Rennomage + décodage du contenu base64 + écriture du fichier sur le disque
                try:
                    with open(filepath, "wb") as f:
                        f.write(file_data)
                    if counter > 1:
                        write_log(PROCESS_TRACKER, f"Doublon détecté : Fichier classé en tant que copie ({counter-1}) -> {final_filename}")
                    else:
                        write_log(PROCESS_TRACKER, f"Fichier enregistré en : {final_filename}")
                    # On ne l'ajoute à la liste que si l'écriture a réussi
                    downloaded_files_info.append({
                        "filepath": filepath,
                        "filename": final_filename,
                        "sender_name": sender_name
                    })
                except Exception as e:
                    write_log(PROCESS_TRACKER, f"Erreur d'écriture pour {filename} : {e}")

    write_log(PROCESS_TRACKER, "Script terminé avec succès.")
    write_log(PROCESS_TRACKER, "========================================\n\n")
    return downloaded_files_info


if __name__ == "__main__":
    download_attachments()