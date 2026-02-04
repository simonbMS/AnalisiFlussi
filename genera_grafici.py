"""
Script per generare grafici dell'andamento dei flussi di cassa.
Eseguire dopo aver lanciato estrai_flussi_cassa.py

Uso: python genera_grafici.py
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
import os
import glob

# Directory dello script (per path relativi)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Configurazione (path relativi)
CSV_RIEPILOGO = os.path.join(SCRIPT_DIR, "flussi_cassa_riepilogo.csv")
CSV_DETTAGLIO = os.path.join(SCRIPT_DIR, "flussi_cassa_dettaglio.csv")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "grafici")

# Stile grafici
plt.style.use('seaborn-v0_8-whitegrid')
COLORS = {
    'entrate': '#2ecc71',
    'uscite': '#e74c3c', 
    'saldo': '#3498db',
    'saldo_positivo': '#27ae60',
    'saldo_negativo': '#c0392b'
}

# Colori distinti per le categorie di spesa
CATEGORY_COLORS = [
    '#e74c3c',  # Rosso
    '#3498db',  # Blu
    '#2ecc71',  # Verde
    '#f39c12',  # Arancione
    '#9b59b6',  # Viola
    '#1abc9c',  # Turchese
    '#e67e22',  # Arancione scuro
    '#34495e',  # Grigio scuro
    '#16a085',  # Verde acqua
    '#c0392b',  # Rosso scuro
    '#8e44ad',  # Viola scuro
    '#27ae60',  # Verde scuro
    '#d35400',  # Arancione bruciato
    '#2980b9',  # Blu scuro
    '#f1c40f',  # Giallo
]


def elimina_grafici_vecchi():
    """Elimina tutti i grafici esistenti prima di generare i nuovi."""
    print("\n🗑️  Pulizia grafici esistenti...")
    
    if os.path.exists(OUTPUT_DIR):
        # Elimina tutti i file .png nella cartella grafici
        files_png = glob.glob(os.path.join(OUTPUT_DIR, "*.png"))
        for f in files_png:
            try:
                os.remove(f)
            except Exception as e:
                print(f"   ⚠️ Errore eliminando {f}: {e}")
        
        # Elimina anche statistiche_report.txt se esiste
        stats_file = os.path.join(OUTPUT_DIR, "statistiche_report.txt")
        if os.path.exists(stats_file):
            try:
                os.remove(stats_file)
            except Exception as e:
                print(f"   ⚠️ Errore eliminando {stats_file}: {e}")
        
        print(f"   ✅ Eliminati {len(files_png)} grafici")
    else:
        print("   ℹ️  Cartella grafici non esistente, verrà creata")


def carica_dati():
    """Carica i dati dai file CSV."""
    print("📂 Caricamento dati...")
    
    df_riepilogo = pd.read_csv(CSV_RIEPILOGO)
    df_dettaglio = pd.read_csv(CSV_DETTAGLIO)
    
    # Converti la colonna data in datetime per ordinamento
    df_riepilogo['data_dt'] = pd.to_datetime(df_riepilogo['data'] + '-01')
    df_riepilogo = df_riepilogo.sort_values('data_dt')
    
    df_dettaglio['data_dt'] = pd.to_datetime(df_dettaglio['data'] + '-01')
    df_dettaglio = df_dettaglio.sort_values('data_dt')
    
    print(f"   ✅ Riepilogo: {len(df_riepilogo)} mesi")
    print(f"   ✅ Dettaglio: {len(df_dettaglio)} record")
    
    return df_riepilogo, df_dettaglio


def grafico_andamento_mensile(df):
    """Grafico 1: Andamento entrate, uscite e saldo nel tempo."""
    
    fig, ax = plt.subplots(figsize=(14, 7))
    
    x = range(len(df))
    
    # Barre per entrate e uscite
    width = 0.35
    bars_entrate = ax.bar([i - width/2 for i in x], df['totale_entrate'], 
                          width, label='Entrate', color=COLORS['entrate'], alpha=0.8)
    bars_uscite = ax.bar([i + width/2 for i in x], df['totale_uscite'].abs(), 
                         width, label='Uscite', color=COLORS['uscite'], alpha=0.8)
    
    # Linea per il saldo
    ax2 = ax.twinx()
    line_saldo = ax2.plot(x, df['saldo'], 'o-', color=COLORS['saldo'], 
                          linewidth=2.5, markersize=8, label='Saldo')
    
    # Evidenzia saldo negativo
    for i, saldo in enumerate(df['saldo']):
        if saldo < 0:
            ax2.plot(i, saldo, 'o', color=COLORS['saldo_negativo'], markersize=12, zorder=5)
    
    # Linea zero per saldo
    ax2.axhline(y=0, color='gray', linestyle='--', alpha=0.5)
    
    # Etichette
    ax.set_xlabel('Mese', fontsize=12)
    ax.set_ylabel('Importo (€) - Entrate/Uscite', fontsize=12)
    ax2.set_ylabel('Saldo (€)', fontsize=12, color=COLORS['saldo'])
    
    ax.set_xticks(x)
    ax.set_xticklabels(df['data_label'], rotation=45, ha='right')
    
    # Legenda combinata
    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
    
    ax.set_title('Andamento Flussi di Cassa Mensili', fontsize=14, fontweight='bold')
    
    # Aggiungi valori sopra le barre del saldo
    for i, (_, row) in enumerate(df.iterrows()):
        color = COLORS['saldo_positivo'] if row['saldo'] >= 0 else COLORS['saldo_negativo']
        ax2.annotate(f"€{row['saldo']:,.0f}", 
                     xy=(i, row['saldo']), 
                     xytext=(0, 10), textcoords='offset points',
                     ha='center', fontsize=8, color=color, fontweight='bold')
    
    plt.tight_layout()
    return fig


def grafico_categorie_spesa(df_dettaglio):
    """Grafico 3: Composizione delle spese per categoria con colori distinti."""
    
    # Filtra solo uscite e raggruppa per categoria
    df_uscite = df_dettaglio[df_dettaglio['tipo'] == 'uscita'].copy()
    df_uscite['importo'] = df_uscite['importo'].abs()
    
    categorie = df_uscite.groupby('categoria')['importo'].sum().sort_values(ascending=False)
    
    # Raggruppa categorie piccole in "Altro"
    threshold = categorie.sum() * 0.03  # 3% del totale
    categorie_principali = categorie[categorie >= threshold]
    altre = categorie[categorie < threshold].sum()
    
    if altre > 0:
        categorie_principali['Altro'] = altre
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    
    # Usa colori distinti invece di sfumature
    colors = CATEGORY_COLORS[:len(categorie_principali)]
    
    # Pie chart
    wedges, texts, autotexts = ax1.pie(categorie_principali, labels=None, 
                                        autopct='%1.1f%%', colors=colors,
                                        pctdistance=0.75, startangle=90)
    
    ax1.legend(wedges, [f"{cat}: €{val:,.0f}" for cat, val in categorie_principali.items()],
               title="Categorie", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    ax1.set_title('Distribuzione Spese per Categoria', fontsize=14, fontweight='bold')
    
    # Bar chart orizzontale
    y_pos = range(len(categorie_principali))
    ax2.barh(y_pos, categorie_principali.values, color=colors)
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(categorie_principali.index)
    ax2.invert_yaxis()
    ax2.set_xlabel('Importo (€)')
    ax2.set_title('Spese per Categoria (Totale Periodo)', fontsize=14, fontweight='bold')
    
    # Aggiungi valori
    for i, v in enumerate(categorie_principali.values):
        ax2.text(v + 100, i, f'€{v:,.0f}', va='center', fontsize=9)
    
    plt.tight_layout()
    return fig


def grafico_trend_categorie(df_dettaglio):
    """Grafico 4: Trend delle principali categorie di spesa nel tempo."""
    
    df_uscite = df_dettaglio[df_dettaglio['tipo'] == 'uscita'].copy()
    df_uscite['importo'] = df_uscite['importo'].abs()
    
    # Trova le top 5 categorie per totale speso
    top_categorie = df_uscite.groupby('categoria')['importo'].sum().nlargest(5).index
    
    # Pivot per avere mesi come colonne e categorie come righe
    df_pivot = df_uscite[df_uscite['categoria'].isin(top_categorie)].pivot_table(
        index='data_label', 
        columns='categoria', 
        values='importo', 
        aggfunc='sum',
        fill_value=0
    )
    
    # Riordina per data
    df_pivot = df_pivot.reindex(df_dettaglio['data_label'].unique())
    
    fig, ax = plt.subplots(figsize=(14, 7))
    
    for i, cat in enumerate(top_categorie):
        if cat in df_pivot.columns:
            ax.plot(df_pivot.index, df_pivot[cat], 'o-', label=cat, 
                    linewidth=2, markersize=6, color=CATEGORY_COLORS[i])
    
    ax.set_xlabel('Mese', fontsize=12)
    ax.set_ylabel('Importo (€)', fontsize=12)
    ax.set_title('Trend delle Principali Categorie di Spesa', fontsize=14, fontweight='bold')
    ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1))
    
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    return fig


def grafico_singola_categoria_spesa(df_dettaglio, categoria, sottocategoria=None, color_idx=0):
    """Genera un grafico per una singola categoria/sottocategoria di spesa nel tempo."""
    
    # Filtra per categoria
    df_cat = df_dettaglio[(df_dettaglio['tipo'] == 'uscita') & 
                          (df_dettaglio['categoria'] == categoria)].copy()
    
    # Se specificata una sottocategoria, filtra anche per quella
    if sottocategoria is not None:
        df_cat = df_cat[df_cat['sottocategoria'] == sottocategoria]
    
    if df_cat.empty:
        return None
    
    df_cat['importo'] = df_cat['importo'].abs()
    
    # Raggruppa per mese
    df_mensile = df_cat.groupby(['data_label', 'data_dt'])['importo'].sum().reset_index()
    df_mensile = df_mensile.sort_values('data_dt')
    
    fig, ax = plt.subplots(figsize=(12, 5))
    
    x = range(len(df_mensile))
    color = CATEGORY_COLORS[color_idx % len(CATEGORY_COLORS)]
    
    # Barre
    bars = ax.bar(x, df_mensile['importo'], color=color, alpha=0.8, edgecolor='white')
    
    # Linea di tendenza
    ax.plot(x, df_mensile['importo'], 'o-', color='#2c3e50', linewidth=2, markersize=6)
    
    # Media
    media = df_mensile['importo'].mean()
    ax.axhline(y=media, color='#e74c3c', linestyle='--', linewidth=2, 
               label=f'Media: €{media:,.0f}')
    
    # Etichette valori
    for i, v in enumerate(df_mensile['importo']):
        ax.annotate(f'€{v:,.0f}', xy=(i, v), xytext=(0, 5), 
                    textcoords='offset points', ha='center', fontsize=8)
    
    ax.set_xlabel('Mese', fontsize=12)
    ax.set_ylabel('Importo (€)', fontsize=12)
    ax.set_xticks(x)
    ax.set_xticklabels(df_mensile['data_label'], rotation=45, ha='right')
    
    # Titolo con categoria e sottocategoria
    if sottocategoria:
        titolo = f'Andamento Spese: {categoria} > {sottocategoria}'
    else:
        titolo = f'Andamento Spese: {categoria}'
    ax.set_title(titolo, fontsize=14, fontweight='bold')
    ax.legend(loc='upper right')
    
    # Statistiche nel grafico
    totale = df_mensile['importo'].sum()
    max_val = df_mensile['importo'].max()
    min_val = df_mensile['importo'].min()
    
    stats_text = f'Totale: €{totale:,.0f}\nMax: €{max_val:,.0f}\nMin: €{min_val:,.0f}'
    ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, fontsize=10,
            verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    
    plt.tight_layout()
    return fig


def grafico_stipendi(df_dettaglio):
    """Grafico: Andamento delle entrate da stipendio (Lavoro) nel tempo."""
    
    # Filtra solo la categoria "Lavoro"
    df_lavoro = df_dettaglio[(df_dettaglio['tipo'] == 'entrata') & 
                             (df_dettaglio['categoria'] == 'Lavoro')].copy()
    
    if df_lavoro.empty:
        print("   ⚠️  Nessun dato trovato per la categoria 'Lavoro'")
        return None
    
    # Raggruppa per mese
    df_mensile = df_lavoro.groupby(['data_label', 'data_dt'])['importo'].sum().reset_index()
    df_mensile = df_mensile.sort_values('data_dt')
    
    fig, ax = plt.subplots(figsize=(14, 6))
    
    x = range(len(df_mensile))
    
    # Barre verdi per gli stipendi
    bars = ax.bar(x, df_mensile['importo'], color=COLORS['entrate'], 
                  alpha=0.8, edgecolor='white', linewidth=1.5)
    
    # Linea di tendenza
    ax.plot(x, df_mensile['importo'], 'o-', color='#27ae60', linewidth=2.5, markersize=8)
    
    # Media
    media = df_mensile['importo'].mean()
    ax.axhline(y=media, color='#3498db', linestyle='--', linewidth=2, 
               label=f'Media mensile: €{media:,.0f}')
    
    # Etichette valori sopra le barre
    for i, v in enumerate(df_mensile['importo']):
        ax.annotate(f'€{v:,.0f}', xy=(i, v), xytext=(0, 8), 
                    textcoords='offset points', ha='center', fontsize=9, fontweight='bold')
    
    ax.set_xlabel('Mese', fontsize=12)
    ax.set_ylabel('Stipendio (€)', fontsize=12)
    ax.set_xticks(x)
    ax.set_xticklabels(df_mensile['data_label'], rotation=45, ha='right')
    ax.set_title('Andamento Stipendi nel Tempo', fontsize=14, fontweight='bold')
    ax.legend(loc='upper right')
    
    # Box con statistiche
    totale = df_mensile['importo'].sum()
    max_val = df_mensile['importo'].max()
    min_val = df_mensile['importo'].min()
    mese_max = df_mensile.loc[df_mensile['importo'].idxmax(), 'data_label']
    mese_min = df_mensile.loc[df_mensile['importo'].idxmin(), 'data_label']
    
    stats_text = (f'Totale periodo: €{totale:,.0f}\n'
                  f'Massimo: €{max_val:,.0f} ({mese_max})\n'
                  f'Minimo: €{min_val:,.0f} ({mese_min})')
    
    ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, fontsize=10,
            verticalalignment='top', bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
    
    plt.tight_layout()
    return fig


def genera_grafici_per_categoria(df_dettaglio):
    """Genera un grafico per ogni combinazione categoria/sottocategoria di spesa."""
    
    df_uscite = df_dettaglio[df_dettaglio['tipo'] == 'uscita'].copy()
    df_uscite['importo'] = df_uscite['importo'].abs()
    
    # Verifica se esiste la colonna sottocategoria
    has_sottocategoria = 'sottocategoria' in df_uscite.columns
    
    grafici_generati = []
    
    if has_sottocategoria:
        # Raggruppa per categoria + sottocategoria
        grouped = df_uscite.groupby(['categoria', 'sottocategoria'])['importo'].sum()
        grouped = grouped.sort_values(ascending=False)
        
        for i, ((categoria, sottocategoria), totale) in enumerate(grouped.items()):
            # Genera nome file sicuro
            cat_sicuro = categoria.replace(' ', '_').replace(',', '').replace('/', '_')
            
            if pd.notna(sottocategoria):
                sotto_sicuro = sottocategoria.replace(' ', '_').replace(',', '').replace('/', '_')
                filename = f"cat_{i+1:02d}_{cat_sicuro}_{sotto_sicuro}.png"
                label = f"{categoria} > {sottocategoria}"
            else:
                filename = f"cat_{i+1:02d}_{cat_sicuro}.png"
                label = categoria
                sottocategoria = None
            
            fig = grafico_singola_categoria_spesa(df_dettaglio, categoria, sottocategoria, i)
            
            if fig is not None:
                grafici_generati.append((filename, fig, label))
    else:
        # Fallback: solo per categoria (compatibilità con dati vecchi)
        categorie = df_uscite.groupby('categoria')['importo'].sum().sort_values(ascending=False)
        
        for i, (categoria, totale) in enumerate(categorie.items()):
            nome_sicuro = categoria.replace(' ', '_').replace(',', '').replace('/', '_')
            filename = f"cat_{i+1:02d}_{nome_sicuro}.png"
            
            fig = grafico_singola_categoria_spesa(df_dettaglio, categoria, None, i)
            
            if fig is not None:
                grafici_generati.append((filename, fig, categoria))
    
    return grafici_generati


def genera_report_statistiche(df_riepilogo, df_dettaglio):
    """Genera un file di testo con statistiche chiave."""
    
    report_path = os.path.join(OUTPUT_DIR, "statistiche_report.txt")
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("=" * 70 + "\n")
        f.write("REPORT STATISTICHE FLUSSI DI CASSA\n")
        f.write(f"Generato il: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n")
        f.write("=" * 70 + "\n\n")
        
        f.write(f"📅 Periodo: {df_riepilogo['data_label'].iloc[0]} - {df_riepilogo['data_label'].iloc[-1]}\n")
        f.write(f"📊 Mesi analizzati: {len(df_riepilogo)}\n\n")
        
        f.write("─" * 50 + "\n")
        f.write("RIEPILOGO GENERALE\n")
        f.write("─" * 50 + "\n")
        f.write(f"Entrate totali:     €{df_riepilogo['totale_entrate'].sum():>15,.2f}\n")
        f.write(f"Uscite totali:      €{df_riepilogo['totale_uscite'].sum():>15,.2f}\n")
        f.write(f"Saldo totale:       €{df_riepilogo['saldo'].sum():>15,.2f}\n\n")
        
        f.write("─" * 50 + "\n")
        f.write("MEDIE MENSILI\n")
        f.write("─" * 50 + "\n")
        f.write(f"Entrate medie:      €{df_riepilogo['totale_entrate'].mean():>15,.2f}\n")
        f.write(f"Uscite medie:       €{df_riepilogo['totale_uscite'].mean():>15,.2f}\n")
        f.write(f"Saldo medio:        €{df_riepilogo['saldo'].mean():>15,.2f}\n\n")
        
        mese_migliore = df_riepilogo.loc[df_riepilogo['saldo'].idxmax()]
        mese_peggiore = df_riepilogo.loc[df_riepilogo['saldo'].idxmin()]
        
        f.write("─" * 50 + "\n")
        f.write("ESTREMI\n")
        f.write("─" * 50 + "\n")
        f.write(f"Mese migliore: {mese_migliore['data_label']} (€{mese_migliore['saldo']:,.2f})\n")
        f.write(f"Mese peggiore: {mese_peggiore['data_label']} (€{mese_peggiore['saldo']:,.2f})\n\n")
        
        # Top categorie
        df_uscite = df_dettaglio[df_dettaglio['tipo'] == 'uscita']
        top_uscite = df_uscite.groupby('categoria')['importo'].sum().sort_values().head(10)
        
        f.write("─" * 50 + "\n")
        f.write("TOP 10 CATEGORIE DI SPESA\n")
        f.write("─" * 50 + "\n")
        for cat, val in top_uscite.items():
            f.write(f"{cat:35} €{val:>12,.2f}\n")
    
    print(f"   💾 Report statistiche: {report_path}")


def main():
    """Funzione principale."""
    
    print("=" * 70)
    print("GENERAZIONE GRAFICI FLUSSI DI CASSA")
    print("=" * 70)
    
    # Crea cartella output
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"\n📁 Cartella output: {OUTPUT_DIR}")
    
    # Elimina grafici esistenti prima di creare i nuovi
    elimina_grafici_vecchi()
    
    # Carica dati
    df_riepilogo, df_dettaglio = carica_dati()
    
    # Grafici principali
    grafici_principali = [
        ("01_andamento_mensile.png", grafico_andamento_mensile, [df_riepilogo]),
        ("02_categorie_spesa.png", grafico_categorie_spesa, [df_dettaglio]),
        ("03_trend_categorie.png", grafico_trend_categorie, [df_dettaglio]),
        ("04_stipendi.png", grafico_stipendi, [df_dettaglio]),
    ]
    
    print("\n📊 Generazione grafici principali...")
    
    for filename, func, args in grafici_principali:
        print(f"\n   🔄 {filename}...")
        try:
            fig = func(*args)
            if fig is not None:
                filepath = os.path.join(OUTPUT_DIR, filename)
                fig.savefig(filepath, dpi=150, bbox_inches='tight', facecolor='white')
                plt.close(fig)
                print(f"   ✅ Salvato: {filepath}")
        except Exception as e:
            print(f"   ❌ Errore: {e}")
    
    # Grafici per singola categoria
    print("\n📊 Generazione grafici per categoria di spesa...")
    
    grafici_categorie = genera_grafici_per_categoria(df_dettaglio)
    
    for filename, fig, categoria in grafici_categorie:
        print(f"   🔄 {categoria}...")
        try:
            filepath = os.path.join(OUTPUT_DIR, filename)
            fig.savefig(filepath, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            print(f"   ✅ Salvato: {filepath}")
        except Exception as e:
            print(f"   ❌ Errore: {e}")
    
    # Genera report statistiche
    print("\n📝 Generazione report...")
    genera_report_statistiche(df_riepilogo, df_dettaglio)
    
    print("\n" + "=" * 70)
    print("✅ GENERAZIONE COMPLETATA")
    print("=" * 70)
    print(f"\n📁 Grafici salvati in: {OUTPUT_DIR}")
    print(f"   • {len(grafici_principali)} grafici principali")
    print(f"   • {len(grafici_categorie)} grafici per categoria")


if __name__ == "__main__":
    main()
