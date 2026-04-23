import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import os
import shutil 
import time
import json
from datetime import datetime, timedelta
import win32com.client as win32
import sys
import webbrowser 
import urllib.request # <--- [NUOVO] Per connettersi a GitHub

# --- GESTIONE PERCORSI E FILE INTERNI ---
def percorso_risorsa(nome_file):
    try:
        percorso_base = sys._MEIPASS
    except Exception:
        percorso_base = os.path.abspath(".")
    return os.path.join(percorso_base, nome_file)

if getattr(sys, 'frozen', False):
    CARTELLA_PROGRAMMA = os.path.dirname(sys.executable)
else:
    CARTELLA_PROGRAMMA = os.path.dirname(os.path.abspath(__file__))

# --- CONFIGURAZIONE E BRANDING ---
NOME_BASE = "SmartSender" 
VERSIONE = "0.9.5"
CREATORE = "Emanuele R. / Intersistemi DC - 2026 ©"

# --- SISTEMA DI LICENZA E AGGIORNAMENTI ---
CHIAVE_SEGRETA = "INVIA-PRO-2026"  
EMAIL_SUPPORTO = "e.rabatti@intersistemidatacapture.it" 
LIMITE_UTILIZZI = 10 
REPO_GITHUB = "M4dGun/SmartSender"

CARTELLA_MODELLI = os.path.join(CARTELLA_PROGRAMMA, "Modelli_Email")
OPZIONE_ESTERNO = "Carica file esterno..."
DIZIONARIO_LINGUE = {"italiano": "ITA", "inglese": "ENG", "italiano+inglese": "ITA+ENG"}

# --- MEMORIA DEL PROGRAMMA ---
FILE_CONFIG = "preferenze.json"
lista_commissari_global = [] 
tag_personalizzati_global = {} 
utilizzi_effettuati = 0 
software_sbloccato = False 

def get_nome_dinamico():
    return f"{NOME_BASE} PRO" if software_sbloccato else f"{NOME_BASE} BASE"

def aggiorna_titoli_gui():
    nome_corrente = get_nome_dinamico()
    root.title(f"{nome_corrente} - v{VERSIONE}")
    lbl_titolo_header.config(text=nome_corrente)

# --- [MOTORE MAGICO] ESTRAZIONE E RICERCA MODELLI ---
def ottieni_modelli_dinamici():
    os.makedirs(CARTELLA_MODELLI, exist_ok=True) 
    modelli_base_nascosti = ["Modello_Standard.oft", "Modello_Standard + Stampa Risposte.oft"] 
    
    for nome_file in modelli_base_nascosti:
        percorso_esterno = os.path.join(CARTELLA_MODELLI, nome_file)
        if not os.path.exists(percorso_esterno):
            percorso_interno = percorso_risorsa(nome_file)
            if os.path.exists(percorso_interno) and percorso_interno != percorso_esterno:
                try: shutil.copy(percorso_interno, percorso_esterno)
                except Exception: pass

    modelli_trovati = {}
    for file in os.listdir(CARTELLA_MODELLI):
        if file.lower().endswith(".oft"):
            nome_pulito = os.path.splitext(file)[0].replace("_", " ").title()
            modelli_trovati[nome_pulito] = os.path.join(CARTELLA_MODELLI, file)
            
    return modelli_trovati

def salva_preferenze():
    config = {
        "prefisso": entry_prefisso.get(),
        "data": entry_data.get(),
        "orario": entry_orario.get(),
        "lingua": combo_lingua.get(),
        "giustificativi": combo_giustificativi.get(),
        "file_csv": entry_csv.get(),
        "file_modello_percorso": entry_modello.get(),
        "modello_selezionato": combo_modelli.get(),
        "saluto": check_saluto_var.get(),
        "bozza": check_bozza_var.get(),
        "programmazione": entry_programmata.get(),
        "commissari": lista_commissari_global,
        "tag_personalizzati": tag_personalizzati_global,
        "utilizzi": utilizzi_effettuati,
        "sbloccato": software_sbloccato,
        "opz_data": usa_data_var.get(),     
        "opz_orario": usa_orario_var.get(), 
        "opz_lingua": usa_lingua_var.get()  
    }
    with open(FILE_CONFIG, 'w') as f:
        json.dump(config, f)

def carica_preferenze():
    global utilizzi_effettuati, software_sbloccato
    if os.path.exists(FILE_CONFIG):
        with open(FILE_CONFIG, 'r') as f:
            dati = json.load(f)
            utilizzi_effettuati = dati.get("utilizzi", 0)
            software_sbloccato = dati.get("sbloccato", False)
            return dati
    return {}

# --- FUNZIONI DEL MENU E OPZIONI ---
def mostra_about():
    messagebox.showinfo("About", f"{get_nome_dinamico()}\nVersione: {VERSIONE}\nCreato da: {CREATORE}\n\nSoftware per l'automazione avanzata di Outlook tramite Mail Merge.")

def gestisci_opzioni_invio():
    finestra = tk.Toplevel(root)
    finestra.title("Opzioni Invio")
    finestra.geometry("320x250")
    finestra.config(padx=15, pady=15)
    finestra.transient(root)
    
    tk.Label(finestra, text="Costruzione Oggetto Email", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 10))
    
    chk_pref = tk.Checkbutton(finestra, text="Prefisso Oggetto (Sempre attivo)", state=tk.DISABLED)
    chk_pref.select() 
    chk_pref.pack(anchor="w", pady=2)
    
    tk.Checkbutton(finestra, text="Includi Data Test", variable=usa_data_var).pack(anchor="w", pady=2)
    tk.Checkbutton(finestra, text="Includi Orario Test", variable=usa_orario_var).pack(anchor="w", pady=2)
    tk.Checkbutton(finestra, text="Includi Lingua Test", variable=usa_lingua_var).pack(anchor="w", pady=2)
    
    tk.Button(finestra, text="Salva e Chiudi", command=lambda: [salva_preferenze(), finestra.destroy()]).pack(pady=20)

# --- [NUOVO] SISTEMA AGGIORNAMENTI INVISIBILE (DOWNLOAD DIRETTO) ---
def controlla_aggiornamenti():
    try:
        url = f"https://api.github.com/repos/{REPO_GITHUB}/releases/latest"
        req = urllib.request.Request(url, headers={'User-Agent': 'SmartSender-App'})
        
        with urllib.request.urlopen(req, timeout=5) as response:
            dati = json.loads(response.read().decode())
            ultima_versione_tag = dati.get('tag_name', '').replace('v', '') 
            
            # Trasforma es. "0.9.4" in (0, 9, 4) per fare un confronto
            v_corrente = tuple(map(int, VERSIONE.split('.')))
            v_online = tuple(map(int, ultima_versione_tag.split('.')))
            
            if v_online > v_corrente:
                risposta = messagebox.askyesno("Aggiornamento Trovato!", f"È disponibile la nuova versione {ultima_versione_tag}!\n\nVuoi scaricarla ora direttamente sul Desktop?")
                
                if risposta:
                    # Cerca il file .exe che hai caricato su GitHub
                    assets = dati.get('assets', [])
                    if assets:
                        link_diretto = assets[0].get('browser_download_url')
                        nome_file_nuovo = assets[0].get('name', 'SmartSender_Aggiornato.exe')
                        
                        # Calcola il percorso del Desktop dell'utente
                        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
                        percorso_salvataggio = os.path.join(desktop, nome_file_nuovo)
                        
                        # Messaggio di attesa
                        messagebox.showinfo("Download in corso", "Il download è iniziato in background.\nAttendi qualche secondo il messaggio di conferma...")
                        root.update() # Aggiorna la grafica per non bloccarla
                        
                        # Scarica il file di nascosto!
                        urllib.request.urlretrieve(link_diretto, percorso_salvataggio)
                        
                        messagebox.showinfo("Successo!", f"Aggiornamento scaricato con successo sul tuo Desktop come:\n\n{nome_file_nuovo}\n\nPuoi chiudere questa vecchia versione e usare quella nuova!")
                    else:
                        messagebox.showerror("Errore", "Nessun file eseguibile trovato nell'aggiornamento online.")
            else:
                messagebox.showinfo("Tutto aggiornato", f"Stai già utilizzando l'ultima versione disponibile ({VERSIONE}).")
                
    except Exception as e:
        messagebox.showerror("Errore di connessione", "Impossibile contattare il server per controllare gli aggiornamenti.\nVerifica la tua connessione internet o il nome del repository.")

def richiedi_chiave_email():
    nome_corrente = get_nome_dinamico()
    oggetto = f"Richiesta Licenza - {nome_corrente} v{VERSIONE}"
    corpo = f"Ciao Emanuele,\n\nVorrei richiedere la chiave di sblocco per continuare ad utilizzare {nome_corrente} in versione illimitata.\n\nGrazie!"
    
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_SUPPORTO
        mail.Subject = oggetto
        mail.Body = corpo
        mail.Display() 
    except Exception as e:
        messagebox.showerror("Errore Outlook", f"Impossibile aprire Outlook:\n{e}")

def gestisci_licenza():
    global software_sbloccato
    finestra = tk.Toplevel(root)
    finestra.title("Gestione Licenza")
    finestra.geometry("350x250")
    finestra.config(padx=20, pady=20)
    finestra.transient(root) 
    
    if software_sbloccato:
        tk.Label(finestra, text="✅ VERSIONE PRO SBLOCCATA", font=("Arial", 12, "bold"), fg="#4CAF50").pack(pady=20)
        tk.Label(finestra, text="Hai accesso illimitato a tutte le funzioni.", font=("Arial", 10)).pack()
        tk.Button(finestra, text="Chiudi", command=finestra.destroy).pack(pady=20)
    else:
        tk.Label(finestra, text="Stato: VERSIONE BASE", font=("Arial", 12, "bold"), fg="#ff9800").pack(pady=(0, 5))
        rimanenti = max(0, LIMITE_UTILIZZI - utilizzi_effettuati)
        tk.Label(finestra, text=f"Utilizzi rimanenti: {rimanenti} su {LIMITE_UTILIZZI}", font=("Arial", 10)).pack(pady=(0, 15))
        
        tk.Label(finestra, text="Inserisci la Chiave di Sblocco:").pack(anchor="w")
        entry_chiave = tk.Entry(finestra, width=30, justify="center")
        entry_chiave.pack(pady=5)
        
        def verifica():
            global software_sbloccato
            if entry_chiave.get().strip() == CHIAVE_SEGRETA:
                software_sbloccato = True
                salva_preferenze()
                aggiorna_titoli_gui() 
                messagebox.showinfo("Sblocco", "Chiave corretta! Hai sbloccato la versione illimitata.", parent=finestra)
                finestra.destroy()
            else:
                messagebox.showerror("Errore", "La chiave inserita non è valida.", parent=finestra)
                
        tk.Button(finestra, text="🔐 Sblocca Ora", command=verifica, bg="#2196F3", fg="white", font=("Arial", 10, "bold")).pack(pady=10)

def mostra_tag_base():
    testo_tag = (
        "TAG PREDEFINITI DISPONIBILI:\n\n"
        "#NOME# : Nome del destinatario\n"
        "#COGNOME# : Cognome del destinatario\n"
        "#AULA# : Aula assegnata\n"
        "#DATA# : Data del test\n"
        "#ORARIO# : Orario del test\n"
        "#LINGUA# : Lingua estesa (es. italiano)\n"
        "#ORARIO_ANTICIPATO# : Orario anticipato di 15 min\n"
        "#ORARIO_APPELLO# : Range +5/+10 min dall'orario\n"
        "#COMMISSIONE# : Elenco commissari formattato\n"
        "#GIUSTIFICATIVO_1# : Formattazione dinamica riga 1\n"
        "#GIUSTIFICATIVO_2# : Formattazione dinamica riga 2"
    )
    messagebox.showinfo("Lista Tag Base", testo_tag)

def gestisci_tag_personalizzati():
    finestra = tk.Toplevel(root)
    finestra.title("Configura Tag Personalizzati")
    finestra.geometry("400x350")
    finestra.config(padx=10, pady=10)
    finestra.transient(root) 
    
    tk.Label(finestra, text="Aggiungi nuovi Tag mappandoli all'Excel", font=("Arial", 10, "bold")).pack(pady=(0,10))
    
    frame_add = tk.Frame(finestra)
    frame_add.pack(fill="x")
    tk.Label(frame_add, text="Tag (es. #RUOLO#):").grid(row=0, column=0, sticky="w")
    entry_tag = tk.Entry(frame_add, width=15)
    entry_tag.grid(row=1, column=0, padx=(0,5))
    
    tk.Label(frame_add, text="Colonna Excel esatta:").grid(row=0, column=1, sticky="w")
    entry_col = tk.Entry(frame_add, width=20)
    entry_col.grid(row=1, column=1, padx=(0,5))
    
    def aggiungi_tag():
        tag = entry_tag.get().strip().upper()
        col = entry_col.get().strip()
        if not tag.startswith("#") or not tag.endswith("#"):
            messagebox.showerror("Errore", "Il Tag deve iniziare e finire con '#'", parent=finestra)
            return
        if not col: return
        tag_personalizzati_global[tag] = col
        aggiorna_lista()
        entry_tag.delete(0, tk.END); entry_col.delete(0, tk.END)

    tk.Button(frame_add, text="Aggiungi", command=aggiungi_tag).grid(row=1, column=2)
    
    lista_box = tk.Listbox(finestra, height=10)
    lista_box.pack(fill="both", expand=True, pady=10)
    
    def aggiorna_lista():
        lista_box.delete(0, tk.END)
        for t, c in tag_personalizzati_global.items():
            lista_box.insert(tk.END, f"{t} --> CSV: [{c}]")
            
    def elimina_tag():
        selezionato = lista_box.curselection()
        if selezionato:
            testo = lista_box.get(selezionato)
            chiave = testo.split(" ")[0]
            del tag_personalizzati_global[chiave]
            aggiorna_lista()

    tk.Button(finestra, text="Elimina Selezionato", command=elimina_tag, fg="red").pack()
    aggiorna_lista()

def apri_menu_aiuto():
    x = btn_menu.winfo_rootx()
    y = btn_menu.winfo_rooty() + btn_menu.winfo_height()
    menu_info = tk.Menu(root, tearoff=0)
    menu_info.add_command(label="ℹ️ About (Info Programma)", command=mostra_about)
    menu_info.add_separator()
    menu_info.add_command(label="⚙️ Opzioni invio", command=gestisci_opzioni_invio) 
    menu_info.add_command(label="🔖 Lista Tag Predefiniti", command=mostra_tag_base) # Icona corretta!
    menu_info.add_command(label="➕ Crea Tag Personalizzati...", command=gestisci_tag_personalizzati)
    menu_info.add_separator()
    if software_sbloccato:
        menu_info.add_command(label="✅ Licenza PRO Attiva", command=gestisci_licenza)
        # Il tasto Aggiornamenti appare solo se PRO
        menu_info.add_separator()
        menu_info.add_command(label="🔄 Controlla Aggiornamenti", command=controlla_aggiornamenti)
    else:
        rimanenti = max(0, LIMITE_UTILIZZI - utilizzi_effettuati)
        menu_info.add_command(label=f"🔑 Sblocca Programma ({rimanenti} invii rimasti)", command=gestisci_licenza)
        menu_info.add_command(label="✉️ Richiedi Chiave di Sblocco", command=richiedi_chiave_email)
    menu_info.post(x, y)

# --- FUNZIONI GUI DI SERVIZIO ---
def gestisci_scelta_modello(event=None):
    if combo_modelli.get() == OPZIONE_ESTERNO:
        entry_modello.config(state=tk.NORMAL)
        btn_sfoglia_modello.config(state=tk.NORMAL)
    else:
        entry_modello.config(state=tk.DISABLED)
        btn_sfoglia_modello.config(state=tk.DISABLED)

def imposta_commissione():
    global lista_commissari_global
    finestra = tk.Toplevel(root)
    finestra.title("Commissari")
    finestra.geometry("300x400")
    tk.Label(finestra, text="Inserisci un commissario per riga:").pack(pady=5)
    txt_area = tk.Text(finestra, width=30, height=15)
    txt_area.pack(padx=10, pady=5)
    txt_area.insert("1.0", "\n".join(lista_commissari_global))
    
    def salva_e_chiudi():
        global lista_commissari_global
        contenuto = txt_area.get("1.0", tk.END).strip()
        lista_commissari_global = [line.strip() for line in contenuto.split('\n') if line.strip()]
        finestra.destroy()
        lbl_comm_status.config(text=f"Commissari impostati: {len(lista_commissari_global)}")
        btn_commissione.config(bg="#e8f5e9" if lista_commissari_global else "SystemButtonFace")

    tk.Button(finestra, text="Conferma", command=salva_e_chiudi, bg="#4CAF50", fg="white").pack(pady=10)

def crea_struttura_base():
    cartella = filedialog.askdirectory()
    if not cartella: return
    percorso_csv = os.path.join(cartella, "elenco.csv")
    cartella_comuni = os.path.join(cartella, "File_comuni")
    cartella_personali = os.path.join(cartella, "File_personali")
    
    if os.path.exists(percorso_csv) or os.path.exists(cartella_comuni) or os.path.exists(cartella_personali):
        risposta = messagebox.askyesno("Attenzione", "Esistono già dei file di progetto qui.\nIl file elenco.csv verrà sovrascritto.\n\nVuoi procedere comunque?")
        if not risposta: return
        
    try:
        os.makedirs(cartella_comuni, exist_ok=True)
        os.makedirs(cartella_personali, exist_ok=True)
        with open(percorso_csv, mode='w', encoding='latin-1') as f:
            f.write("Cognome;Nome;Email;Aula;Nome_File_Personale\n")
        entry_csv.delete(0, tk.END); entry_csv.insert(0, percorso_csv)
        messagebox.showinfo("OK", "Struttura creata correttamente!")
    except Exception as e: 
        messagebox.showerror("Errore", f"Impossibile creare la struttura:\n{e}")

def seleziona_csv():
    p = filedialog.askopenfilename(filetypes=[("File CSV", "*.csv")])
    if p: entry_csv.delete(0, tk.END); entry_csv.insert(0, p)

def seleziona_modello_esterno():
    p = filedialog.askopenfilename(filetypes=[("Modello Outlook", "*.oft")])
    if p:
        entry_modello.delete(0, tk.END)
        entry_modello.insert(0, p)

# --- MOTORE DI INVIO EMAIL ---
def esegui_invio():
    global utilizzi_effettuati, software_sbloccato
    salva_preferenze()

    if not software_sbloccato and utilizzi_effettuati >= LIMITE_UTILIZZI:
        messagebox.showwarning("Periodo di Prova Terminato", "Hai raggiunto il limite massimo di invii gratuiti.\n\nInserisci la chiave di licenza per continuare a utilizzare il software senza limiti.")
        gestisci_licenza()
        return

    DATA_TEST = entry_data.get()
    ORARIO_TEST = entry_orario.get()
    PREFISSO_OGGETTO = entry_prefisso.get()
    LINGUA_BODY = combo_lingua.get() 
    LINGUA_OBJ = DIZIONARIO_LINGUE.get(LINGUA_BODY, "ITA") 
    GIUSTIFICATIVI = combo_giustificativi.get() 
    
    MODALITA_BOZZA = check_bozza_var.get() 
    USA_SALUTO = check_saluto_var.get()
    FILE_CSV = entry_csv.get()
    INVIO_PROGRAMMATO = entry_programmata.get().strip()
    
    scelta = combo_modelli.get()
    if scelta == OPZIONE_ESTERNO:
        FILE_MODELLO = entry_modello.get()
    else:
        FILE_MODELLO = modelli_disponibili.get(scelta, "")

    if not os.path.exists(FILE_CSV) or not os.path.exists(FILE_MODELLO):
        messagebox.showerror("Errore File", "CSV o Modello mancante.\nVerifica di aver selezionato i file corretti.")
        return

    percorso_base = os.path.dirname(FILE_CSV) 
    CARTELLA_COMUNI = os.path.join(percorso_base, "File_comuni")
    CARTELLA_PERSONALI = os.path.join(percorso_base, "File_personali")

    data_posticipata = None
    if INVIO_PROGRAMMATO:
        try: data_posticipata = datetime.strptime(INVIO_PROGRAMMATO, "%d/%m/%Y %H:%M")
        except: messagebox.showerror("Errore Data", "Formato: GG/MM/AAAA HH:MM"); return

    orario_anticipato = "N/D"; orario_appello = ""
    try:
        ora_obj = datetime.strptime(ORARIO_TEST, "%H:%M")
        orario_anticipato = (ora_obj - timedelta(minutes=15)).strftime("%H:%M")
        orario_appello = f"{(ora_obj + timedelta(minutes=5)).strftime('%H:%M')}-{(ora_obj + timedelta(minutes=10)).strftime('%H:%M')}"
    except: pass

    testo_commissione = "<br>".join(lista_commissari_global) if lista_commissari_global else "I nomi verranno comunicati sul gruppo"
    
    if GIUSTIFICATIVI == "SI":
        html_g1 = "<u><b><font color='red'>Riguardo i giustificativi/permessi per il lavoro</font></b> segna sul registro – nell’apposita colonna - chi ne fa esplicita richiesta, <b><font color='red'>vanno quindi RICHIESTI!</font></b></u>"
        html_g2 = "<ul><li><b>Campo PERMESSO per il LAVORO:</b> durante la fase 1 chiedi a tutti chi ne ha necessità, segna poi <b>SI</b> oppure <b>NO</b> - a tutti mi raccomando;</li></ul>"
    else:
        html_g1 = "<u>I giustificativi/permessi per il lavoro <b><font color='red'>NON</font></b> andranno <b><font color='red'>RICHIESTI!</font></b></u>"
        html_g2 = ""

    pulsante_avvia.config(text="Elaborazione in corso...", state=tk.DISABLED)
    lbl_progress.config(text="Lettura dati...")
    root.update()

    try:
        with open(FILE_CSV, mode='r', encoding='latin-1') as file:
            lettore = csv.DictReader(file, delimiter=';')
            dati_puliti = list(lettore) 

        errori = [f"{p['Cognome']} {p['Nome']}" for p in dati_puliti if not os.path.exists(os.path.join(CARTELLA_PERSONALI, p.get('Nome_File_Personale', '').strip()))]
        if errori and not messagebox.askyesno("File Mancanti", "Mancano file personali per alcune persone. Saltarle e procedere?"):
            ripristina_pulsante(); return 

        outlook = win32.Dispatch('outlook.application')
        file_comuni = [os.path.join(CARTELLA_COMUNI, f) for f in os.listdir(CARTELLA_COMUNI) if os.path.isfile(os.path.join(CARTELLA_COMUNI, f))] if os.path.exists(CARTELLA_COMUNI) else []

        totale_email = len([p for p in dati_puliti if p.get('Nome_File_Personale', '').strip() not in errori])
        progress['maximum'] = totale_email; progress['value'] = 0; email_processate = 0

        for persona in dati_puliti:
            nome_file = persona.get('Nome_File_Personale', '').strip()
            if not nome_file or not os.path.exists(os.path.join(CARTELLA_PERSONALI, nome_file)): continue 
                
            messaggio = outlook.CreateItemFromTemplate(FILE_MODELLO)
            messaggio.To = persona.get('Email', '')
            
            parti_oggetto = [PREFISSO_OGGETTO.strip()]
            if usa_data_var.get() and DATA_TEST: 
                parti_oggetto.append(DATA_TEST)
            if usa_orario_var.get() and ORARIO_TEST: 
                parti_oggetto.append(f"- ore {ORARIO_TEST}")
            
            aula = persona.get('Aula','')
            if aula: 
                parti_oggetto.append(f"- Aula {aula}")
                
            if usa_lingua_var.get() and LINGUA_OBJ: 
                parti_oggetto.append(f"({LINGUA_OBJ})")
            
            messaggio.Subject = " ".join(parti_oggetto)
            
            html = messaggio.HTMLBody
            html = html.replace("#NOME#", persona.get('Nome','')).replace("#COGNOME#", persona.get('Cognome',''))
            html = html.replace("#AULA#", persona.get('Aula','')).replace("#DATA#", DATA_TEST)
            html = html.replace("#ORARIO#", ORARIO_TEST).replace("#ORARIO_ANTICIPATO#", orario_anticipato)
            html = html.replace("#LINGUA#", LINGUA_BODY) 
            html = html.replace("#ORARIO_APPELLO#", orario_appello)
            html = html.replace("#COMMISSIONE#", testo_commissione) 
            html = html.replace("#GIUSTIFICATIVO_1#", html_g1).replace("#GIUSTIFICATIVO_2#", html_g2)
            
            for tag_pers, colonna_csv in tag_personalizzati_global.items():
                valore_colonna = persona.get(colonna_csv, "")
                html = html.replace(tag_pers, valore_colonna)
            
            messaggio.HTMLBody = html
            if USA_SALUTO: messaggio.HTMLBody = f"<p style='font-family: Calibri; font-size: 11pt;'>Ciao {persona.get('Nome','')},</p>" + messaggio.HTMLBody

            for f in file_comuni: messaggio.Attachments.Add(f)
            messaggio.Attachments.Add(os.path.join(CARTELLA_PERSONALI, nome_file))
            if data_posticipata: messaggio.DeferredDeliveryTime = data_posticipata

            if MODALITA_BOZZA: messaggio.Save()
            else: messaggio.Send(); time.sleep(1) 

            email_processate += 1
            progress['value'] = email_processate
            lbl_progress.config(text=f"Elaborazione: {email_processate} di {totale_email}")
            root.update()

        messagebox.showinfo("Completato!", f"{email_processate} email sono state elaborate con successo!")
        
        if not software_sbloccato:
            utilizzi_effettuati += 1
            salva_preferenze()
            aggiorna_titoli_gui() 
        
    except Exception as e: messagebox.showerror("Errore Critico", str(e))
    ripristina_pulsante()

def ripristina_pulsante():
    pulsante_avvia.config(text="🚀 INVIA EMAIL", state=tk.NORMAL)
    lbl_progress.config(text=""); progress['value'] = 0

# --- SETUP INTERFACCIA PRINCIPALE ---
root = tk.Tk()
root.geometry("540x820") 
root.config(padx=15, pady=15)

try: root.iconbitmap(percorso_risorsa("logo.ico"))
except: pass

pref = carica_preferenze() 
lista_commissari_global = pref.get("commissari", [])
tag_personalizzati_global = pref.get("tag_personalizzati", {})

usa_data_var = tk.BooleanVar(value=pref.get("opz_data", True))
usa_orario_var = tk.BooleanVar(value=pref.get("opz_orario", True))
usa_lingua_var = tk.BooleanVar(value=pref.get("opz_lingua", True))

modelli_disponibili = ottieni_modelli_dinamici()

# HEADER
frame_header = tk.Frame(root)
frame_header.pack(fill="x", pady=(0,10))
lbl_titolo_header = tk.Label(frame_header, text="", font=("Arial", 16, "bold"), fg="#333")
lbl_titolo_header.pack(side="left")

# [NUOVO] Pulsante Ingranaggio
btn_menu = tk.Button(frame_header, text=" ⚙️ ", font=("Arial", 12), cursor="hand2", command=apri_menu_aiuto)
btn_menu.pack(side="right")

aggiorna_titoli_gui() 

# SEZIONE 1
tk.Button(root, text="🛠️ Genera Struttura Progetto (Cartelle e CSV)", command=crea_struttura_base, bg="#2196F3", fg="white", font=("Arial", 10, "bold")).pack(fill="x", pady=(0, 15))

frame_file = tk.LabelFrame(root, text=" Sorgente Dati ", padx=10, pady=10)
frame_file.pack(fill="x", pady=5)
frame_file.columnconfigure(0, weight=1)

tk.Label(frame_file, text="Elenco Dati (file CSV):").grid(row=0, column=0, sticky="w")
entry_csv = tk.Entry(frame_file)
entry_csv.insert(0, pref.get("file_csv", "elenco.csv"))
entry_csv.grid(row=1, column=0, pady=2, padx=(0,5), sticky="ew") 
tk.Button(frame_file, text="Sfoglia...", command=seleziona_csv).grid(row=1, column=1)

tk.Label(frame_file, text="Modello Email (file OFT):").grid(row=2, column=0, sticky="w", pady=(10,0))

lista_opzioni = list(modelli_disponibili.keys()) + [OPZIONE_ESTERNO]
combo_modelli = ttk.Combobox(frame_file, values=lista_opzioni, state="readonly")
modello_salvato = pref.get("modello_selezionato", OPZIONE_ESTERNO)
if modello_salvato not in lista_opzioni: modello_salvato = OPZIONE_ESTERNO
combo_modelli.set(modello_salvato)

combo_modelli.grid(row=3, column=0, columnspan=2, pady=2, sticky="ew")
combo_modelli.bind("<<ComboboxSelected>>", gestisci_scelta_modello)

frame_mod_esterno = tk.Frame(frame_file)
frame_mod_esterno.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(2,0))
frame_mod_esterno.columnconfigure(0, weight=1)
entry_modello = tk.Entry(frame_mod_esterno)
entry_modello.insert(0, pref.get("file_modello_percorso", ""))
entry_modello.grid(row=0, column=0, sticky="ew", padx=(0,5))
btn_sfoglia_modello = tk.Button(frame_mod_esterno, text="Sfoglia...", command=seleziona_modello_esterno)
btn_sfoglia_modello.grid(row=0, column=1)
gestisci_scelta_modello() 

# SEZIONE 2
frame_dati = tk.LabelFrame(root, text=" Impostazioni Email: Oggetto e Contenuto ", padx=10, pady=10)
frame_dati.pack(fill="x", pady=10)

frame_riga1 = tk.Frame(frame_dati)
frame_riga1.pack(fill="x", pady=(0, 10))
frame_riga1.columnconfigure(0, weight=1)
frame_riga1.columnconfigure(1, weight=1)

tk.Label(frame_riga1, text="Prefisso Oggetto:").grid(row=0, column=0, sticky="w")
entry_prefisso = tk.Entry(frame_riga1)
entry_prefisso.insert(0, pref.get("prefisso", "Materiali Test"))
entry_prefisso.grid(row=1, column=0, sticky="ew", padx=(0, 10))

tk.Label(frame_riga1, text="Data Test (es. 15 Aprile):").grid(row=0, column=1, sticky="w")
entry_data = tk.Entry(frame_riga1)
entry_data.insert(0, pref.get("data", "15 Aprile"))
entry_data.grid(row=1, column=1, sticky="ew")

frame_riga2 = tk.Frame(frame_dati)
frame_riga2.pack(fill="x", pady=(0, 10))
frame_riga2.columnconfigure(0, weight=1)
frame_riga2.columnconfigure(1, weight=1)
frame_riga2.columnconfigure(2, weight=1)

tk.Label(frame_riga2, text="Orario (HH:MM):").grid(row=0, column=0, sticky="w")
entry_orario = tk.Entry(frame_riga2)
entry_orario.insert(0, pref.get("orario", "10:00"))
entry_orario.grid(row=1, column=0, sticky="ew", padx=(0, 10))

tk.Label(frame_riga2, text="Lingua Test:").grid(row=0, column=1, sticky="w")
combo_lingua = ttk.Combobox(frame_riga2, values=list(DIZIONARIO_LINGUE.keys()), state="readonly")
combo_lingua.set(pref.get("lingua", "italiano"))
combo_lingua.grid(row=1, column=1, sticky="ew", padx=(0, 10))

tk.Label(frame_riga2, text="Giustificativi Lavoro:").grid(row=0, column=2, sticky="w")
combo_giustificativi = ttk.Combobox(frame_riga2, values=["SI", "NO"], state="readonly")
combo_giustificativi.set(pref.get("giustificativi", "SI"))
combo_giustificativi.grid(row=1, column=2, sticky="ew")

frame_comm = tk.LabelFrame(frame_dati, text=" Commissione ", padx=10, pady=10)
frame_comm.pack(fill="x", pady=(5, 0))

colore_comm = "#e8f5e9" if lista_commissari_global else "SystemButtonFace"
btn_commissione = tk.Button(frame_comm, text="👥 Imposta Commissari", command=imposta_commissione, bg=colore_comm)
btn_commissione.pack(side="left")
lbl_comm_status = tk.Label(frame_comm, text=f"Commissari impostati: {len(lista_commissari_global)}")
lbl_comm_status.pack(side="left", padx=10)

# SEZIONE 3
frame_opz = tk.Frame(root)
frame_opz.pack(fill="x", pady=(5, 0))

tk.Label(frame_opz, text="Programma Invio Posticipato (opzionale) (Es. GG/MM/AAAA HH:MM):", fg="#555").pack(anchor="w")
entry_programmata = tk.Entry(frame_opz, width=40)
entry_programmata.insert(0, pref.get("programmazione", ""))
entry_programmata.pack(anchor="w", pady=(0, 15))

tk.Frame(frame_opz, height=1, bg="#ddd").pack(fill="x", pady=(0, 10))

check_saluto_var = tk.BooleanVar(value=pref.get("saluto", False)) 
tk.Checkbutton(frame_opz, text="Aggiungi 'Ciao [Nome],' personalizzato in cima", variable=check_saluto_var).pack(anchor="w")

check_bozza_var = tk.BooleanVar(value=pref.get("bozza", True)) 
tk.Checkbutton(frame_opz, text="Modalità Sicura (Salva in Bozze, non invia)", variable=check_bozza_var, fg="#d32f2f", font=("Arial", 9, "bold")).pack(anchor="w")

# --- AVVIO E PROGRESS BAR ---
pulsante_avvia = tk.Button(root, text="🚀 INVIA EMAIL", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", height=2, command=esegui_invio)
pulsante_avvia.pack(fill="x", pady=(10, 2))

lbl_progress = tk.Label(root, text="", font=("Arial", 9, "italic"), fg="#666")
lbl_progress.pack()
progress = ttk.Progressbar(root, orient="horizontal", length=480, mode="determinate")
progress.pack(pady=(0, 10))

root.mainloop()
