"""
Script principale per l'analisi mensile dei flussi di cassa.
Esegue in sequenza: estrazione dati -> generazione grafici -> report personalizzato

USO:
    python analisi_mensile.py

Questo script:
1. Estrae i dati dal file Excel "Flusso di cassa.xlsx"
2. Genera i file CSV e JSON con i dati strutturati
3. Crea i grafici di analisi per ogni categoria/sottocategoria
4. Genera grafici aggregati per le categorie in Categorie_per_grafici.csv
5. Produce un report Markdown con i grafici selezionati

Eseguire ogni mese dopo aver aggiornato il file Excel.
"""

import subprocess
import sys
import os
import glob
from datetime import datetime

# Directory dello script (per path relativi - permette di spostare la cartella)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORK_DIR = SCRIPT_DIR
PYTHON_EXE = os.path.join(SCRIPT_DIR, ".venv", "Scripts", "python.exe")


def esegui_script(script_name, descrizione):
    """Esegue uno script Python e gestisce gli errori."""
    
    script_path = os.path.join(WORK_DIR, script_name)
    
    print(f"\n{'─' * 60}")
    print(f"▶ {descrizione}")
    print(f"  Script: {script_name}")
    print(f"{'─' * 60}\n")
    
    try:
        result = subprocess.run(
            [PYTHON_EXE, script_path],
            cwd=WORK_DIR,
            capture_output=False,
            text=True
        )
        
        if result.returncode != 0:
            print(f"\n❌ ERRORE: {script_name} terminato con codice {result.returncode}")
            return False
            
        return True
        
    except Exception as e:
        print(f"\n❌ ERRORE nell'esecuzione di {script_name}: {e}")
        return False


def pulisci_file_temporanei():
    """Elimina i file CSV e JSON temporanei generati durante l'analisi."""
    print("\n🗑️  Pulizia file temporanei...")
    
    files_temp = [
        os.path.join(WORK_DIR, "flussi_cassa_riepilogo.csv"),
        os.path.join(WORK_DIR, "flussi_cassa_dettaglio.csv"),
        os.path.join(WORK_DIR, "flussi_cassa.json"),
    ]
    
    eliminati = 0
    for f in files_temp:
        if os.path.exists(f):
            try:
                os.remove(f)
                eliminati += 1
            except Exception as e:
                print(f"   ⚠️ Errore eliminando {os.path.basename(f)}: {e}")
    
    print(f"   ✅ Eliminati {eliminati} file temporanei")


def main():
    """Funzione principale."""
    
    print("\n" + "=" * 70)
    print("  ANALISI MENSILE FLUSSI DI CASSA")
    print(f"  Data esecuzione: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 70)
    
    # Verifica che il file Excel esista
    excel_file = os.path.join(WORK_DIR, "Flusso di cassa.xlsx")
    if not os.path.exists(excel_file):
        print(f"\n❌ ERRORE: File non trovato: {excel_file}")
        print("   Assicurati che il file Excel sia nella cartella corretta.")
        sys.exit(1)
    
    print(f"\n✅ File Excel trovato: {excel_file}")
    print(f"   Ultima modifica: {datetime.fromtimestamp(os.path.getmtime(excel_file)).strftime('%d/%m/%Y %H:%M')}")
    
    # Step 1: Estrazione dati
    if not esegui_script("estrai_flussi_cassa.py", "STEP 1: Estrazione dati da Excel"):
        print("\n⛔ Analisi interrotta per errore nell'estrazione dati.")
        sys.exit(1)
    
    # Step 2: Generazione grafici
    if not esegui_script("genera_grafici.py", "STEP 2: Generazione grafici"):
        print("\n⚠️ Attenzione: errore nella generazione grafici, ma i dati sono stati estratti.")
    
    # Step 3 & 4: Generazione report personalizzato (se esiste il file config)
    config_file = os.path.join(WORK_DIR, "Categorie_per_grafici.csv")
    if os.path.exists(config_file):
        if not esegui_script("genera_report.py", "STEP 3-4: Grafici aggregati e Report Markdown"):
            print("\n⚠️ Attenzione: errore nella generazione report personalizzato.")
    else:
        print(f"\n⚠️ File {config_file} non trovato, skip step 3-4.")
    
    # Pulizia file temporanei
    pulisci_file_temporanei()
    
    # Riepilogo finale
    print("\n" + "=" * 70)
    print("  ✅ ANALISI COMPLETATA CON SUCCESSO")
    print("=" * 70)
    
    print("\n📁 FILE GENERATI:")
    print(f"   • {os.path.join(WORK_DIR, 'grafici', '*.png')}")
    print(f"   • {os.path.join(WORK_DIR, 'grafici', 'statistiche_report.txt')}")
    
    if os.path.exists(config_file):
        print(f"   • {os.path.join(WORK_DIR, 'Report_Flussi_Cassa.md')}")
    
    print("\n💡 PROSSIMI PASSI:")
    print("   1. Apri la cartella 'grafici' per visualizzare i grafici")
    print("   2. Consulta 'statistiche_report.txt' per le statistiche chiave")
    print("   3. Apri 'Report_Flussi_Cassa.md' per il report personalizzato")


if __name__ == "__main__":
    main()
