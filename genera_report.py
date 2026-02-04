"""
Script per generare grafici aggregati per categoria e report Markdown.
Legge le categorie da Categorie_per_grafici.csv

Uso: python genera_report.py
"""

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import html

# Directory dello script (per path relativi)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Configurazione (path relativi)
CSV_DETTAGLIO = os.path.join(SCRIPT_DIR, "flussi_cassa_dettaglio.csv")
CSV_CATEGORIE = os.path.join(SCRIPT_DIR, "Categorie_per_grafici.csv")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "grafici")
REPORT_FILE = os.path.join(SCRIPT_DIR, "Report_Flussi_Cassa.md")
REPORT_HTML_FILE = os.path.splitext(REPORT_FILE)[0] + ".html"

# Stile grafici
plt.style.use('seaborn-v0_8-whitegrid')

# Colori distinti per le sottocategorie
CATEGORY_COLORS = [
    '#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6',
    '#1abc9c', '#e67e22', '#34495e', '#16a085', '#c0392b',
    '#8e44ad', '#27ae60', '#d35400', '#2980b9', '#f1c40f',
]


def carica_dati():
    """Carica i dati dal CSV dettaglio."""
    print("📂 Caricamento dati...")
    
    df = pd.read_csv(CSV_DETTAGLIO)
    df['data_dt'] = pd.to_datetime(df['data'] + '-01')
    df = df.sort_values('data_dt')
    
    print(f"   ✅ Dettaglio: {len(df)} record")
    return df


def carica_categorie_config():
    """Carica la configurazione delle categorie dal CSV."""
    print("📂 Caricamento configurazione categorie...")
    
    df_config = pd.read_csv(CSV_CATEGORIE)
    
    # Normalizza i nomi (sostituisci _ con spazi)
    df_config['Categoria'] = df_config['Categoria'].str.replace('_', ' ')
    df_config['Sottocategoria'] = df_config['Sottocategoria'].str.replace('_', ' ')
    
    print(f"   ✅ Categorie configurate: {len(df_config)}")
    return df_config


def grafico_categoria_aggregata(df_dettaglio, categoria, sottocategorie_filtro):
    """
    Genera un grafico a barre che mostra l'andamento mensile 
    di una categoria con le sottocategorie impilate.
    
    Args:
        df_dettaglio: DataFrame con i dati
        categoria: Nome della categoria
        sottocategorie_filtro: Lista di sottocategorie o ['*'] per tutte
    """
    
    # Filtra per categoria e tipo uscita
    df_cat = df_dettaglio[(df_dettaglio['categoria'] == categoria) & 
                          (df_dettaglio['tipo'] == 'uscita')].copy()
    
    if df_cat.empty:
        return None, None
    
    # Se '*', prendi tutte le sottocategorie
    if '*' in sottocategorie_filtro:
        sottocategorie = df_cat['sottocategoria'].dropna().unique().tolist()
    else:
        sottocategorie = sottocategorie_filtro
        df_cat = df_cat[df_cat['sottocategoria'].isin(sottocategorie)]
    
    if df_cat.empty:
        return None, None
    
    df_cat['importo'] = df_cat['importo'].abs()
    
    # Pivot per avere mesi come righe e sottocategorie come colonne
    df_pivot = df_cat.pivot_table(
        index='data_label',
        columns='sottocategoria',
        values='importo',
        aggfunc='sum',
        fill_value=0
    )
    
    # Ordina per data
    mesi_ordinati = df_dettaglio.sort_values('data_dt')['data_label'].unique()
    df_pivot = df_pivot.reindex([m for m in mesi_ordinati if m in df_pivot.index])
    
    # Ordina le colonne per totale (sottocategorie più costose prima)
    col_order = df_pivot.sum().sort_values(ascending=False).index
    df_pivot = df_pivot[col_order]
    
    # Crea il grafico a barre impilate
    fig, ax = plt.subplots(figsize=(14, 7))
    
    x = range(len(df_pivot))
    width = 0.7
    bottom = [0] * len(df_pivot)
    
    bars_list = []
    for i, sotto in enumerate(df_pivot.columns):
        color = CATEGORY_COLORS[i % len(CATEGORY_COLORS)]
        bars = ax.bar(x, df_pivot[sotto], width, bottom=bottom, 
                      label=sotto, color=color, alpha=0.85)
        bars_list.append(bars)
        bottom = [b + v for b, v in zip(bottom, df_pivot[sotto])]
    
    # Aggiungi totale sopra ogni barra
    for i, total in enumerate(bottom):
        ax.annotate(f'€{total:,.0f}', xy=(i, total), xytext=(0, 5),
                    textcoords='offset points', ha='center', fontsize=9, fontweight='bold')
    
    # Etichette
    ax.set_xlabel('Mese', fontsize=12)
    ax.set_ylabel('Importo (€)', fontsize=12)
    ax.set_xticks(x)
    ax.set_xticklabels(df_pivot.index, rotation=45, ha='right')
    
    # Titolo
    if '*' in sottocategorie_filtro:
        titolo = f'Andamento Spese: {categoria} (tutte le sottocategorie)'
    else:
        titolo = f'Andamento Spese: {categoria}'
    ax.set_title(titolo, fontsize=14, fontweight='bold')
    
    # Legenda
    ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1), title='Sottocategorie')
    
    # Box statistiche
    totale_periodo = df_pivot.sum().sum()
    media_mensile = df_pivot.sum(axis=1).mean()
    
    stats_text = f'Totale periodo: €{totale_periodo:,.0f}\nMedia mensile: €{media_mensile:,.0f}'
    ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, fontsize=10,
            verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    
    plt.tight_layout()
    
    # Informazioni per il report
    info = {
        'categoria': categoria,
        'sottocategorie': list(df_pivot.columns),
        'totale': totale_periodo,
        'media_mensile': media_mensile,
        'mese_max': df_pivot.sum(axis=1).idxmax(),
        'max_val': df_pivot.sum(axis=1).max(),
    }
    
    return fig, info


def genera_grafici_aggregati(df_dettaglio, df_config):
    """Genera i grafici aggregati secondo la configurazione."""
    
    print("\n📊 Generazione grafici aggregati...")
    
    grafici_info = []
    
    # Raggruppa per categoria
    categorie_config = df_config.groupby('Categoria')['Sottocategoria'].apply(list).to_dict()
    
    for i, (categoria, sottocategorie) in enumerate(categorie_config.items()):
        print(f"\n   🔄 {categoria}...")
        
        fig, info = grafico_categoria_aggregata(df_dettaglio, categoria, sottocategorie)
        
        if fig is not None:
            # Nome file
            nome_sicuro = categoria.replace(' ', '_').replace(',', '').replace('/', '_')
            filename = f"agg_{i+1:02d}_{nome_sicuro}.png"
            filepath = os.path.join(OUTPUT_DIR, filename)
            
            fig.savefig(filepath, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            
            info['filename'] = filename
            grafici_info.append(info)
            
            print(f"   ✅ Salvato: {filepath}")
        else:
            print(f"   ⚠️  Nessun dato per {categoria}")
    
    return grafici_info


def _get_html_css() -> str:
    """Restituisce il CSS per il report HTML."""
    return """
:root { --text:#24292f; --muted:#57606a; --border:#d0d7de; --bg:#ffffff; --bg-subtle:#f6f8fa; }
body { background: var(--bg); margin: 0; }
.report-body { box-sizing: border-box; min-width: 200px; max-width: 900px; margin: 24px auto; padding: 0 20px; color: var(--text); font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; }
.report-body h1 { font-size: 2em; border-bottom: 1px solid var(--border); padding-bottom: .3em; }
.report-body h2 { font-size: 1.5em; border-bottom: 1px solid var(--border); padding-bottom: .3em; }
.report-body h3 { font-size: 1.25em; }
.report-body h1, .report-body h2, .report-body h3 { margin-top: 1.4em; }
.report-body a { color: #0969da; text-decoration: none; }
.report-body a:hover { text-decoration: underline; }
.report-body img { max-width: 100%; height: auto; border: 1px solid #eee; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
.report-body hr { border: 0; border-top: 1px solid var(--border); margin: 24px 0; }
.report-body ul, .report-body ol { padding-left: 1.5em; }
.report-body code { background: var(--bg-subtle); padding: 0.1em 0.4em; border-radius: 4px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
.report-body table { border-collapse: collapse; width: 100%; margin: 16px 0; }
.report-body th, .report-body td { border: 1px solid var(--border); padding: 8px; text-align: left; }
.report-body th { background: #f7f7f7; }
.report-body .toc { font-size: .95em; color: var(--muted); }
.report-body .date { color: var(--muted); font-style: italic; }
"""


def genera_report_html_nativo(grafici_info, df_config) -> str:
    """Genera il report HTML con tag nativi (senza conversione da Markdown)."""

    html_parts = []

    # Intestazione
    html_parts.append(f'<h1>📊 Report Flussi di Cassa</h1>')
    html_parts.append(f'<p class="date">Generato il: {datetime.now().strftime("%d/%m/%Y %H:%M")}</p>')
    html_parts.append('<hr>')

    # Indice
    html_parts.append('<h2>📑 Indice</h2>')
    html_parts.append('<ol class="toc">')
    html_parts.append('<li><a href="#panoramica-generale">Panoramica generale</a></li>')
    for info in grafici_info:
        anchor = info['categoria'].lower().replace(' ', '-').replace(',', '')
        html_parts.append(f'<li><a href="#{anchor}">{html.escape(info["categoria"])}</a></li>')
    html_parts.append('</ol>')
    html_parts.append('<hr>')

    # Sezione panoramica con grafici generali
    html_parts.append('<h2 id="panoramica-generale">Panoramica generale</h2>')
    html_parts.append('<h3>📈 Andamento mensile</h3>')
    html_parts.append('<p><img src="grafici/01_andamento_mensile.png" alt="Andamento mensile"></p>')
    html_parts.append('<h3>💰 Categorie di spesa</h3>')
    html_parts.append('<p><img src="grafici/02_categorie_spesa.png" alt="Categorie spesa"></p>')
    html_parts.append('<hr>')

    # Sezioni per ogni categoria
    for info in grafici_info:
        categoria = info['categoria']
        anchor = categoria.lower().replace(' ', '-').replace(',', '')
        
        html_parts.append(f'<h2 id="{anchor}">{html.escape(categoria)}</h2>')

        # Statistiche
        html_parts.append('<h3>📈 Statistiche</h3>')
        html_parts.append('<table>')
        html_parts.append('<thead><tr><th>Metrica</th><th>Valore</th></tr></thead>')
        html_parts.append('<tbody>')
        html_parts.append(f'<tr><td><strong>Totale periodo</strong></td><td>€{info["totale"]:,.2f}</td></tr>')
        html_parts.append(f'<tr><td><strong>Media mensile</strong></td><td>€{info["media_mensile"]:,.2f}</td></tr>')
        html_parts.append(f'<tr><td><strong>Mese con spesa max</strong></td><td>{html.escape(info["mese_max"])} (€{info["max_val"]:,.2f})</td></tr>')
        html_parts.append(f'<tr><td><strong>Sottocategorie</strong></td><td>{len(info["sottocategorie"])}</td></tr>')
        html_parts.append('</tbody>')
        html_parts.append('</table>')

        # Lista sottocategorie
        html_parts.append('<h3>📋 Sottocategorie incluse</h3>')
        html_parts.append('<ul>')
        for sotto in info['sottocategorie']:
            html_parts.append(f'<li>{html.escape(sotto)}</li>')
        html_parts.append('</ul>')

        # Grafico
        html_parts.append('<h3>📊 Grafico</h3>')
        html_parts.append(f'<p><img src="grafici/{info["filename"]}" alt="{html.escape(categoria)}"></p>')
        html_parts.append('<hr>')

    # Footer
    html_parts.append('<h2>📁 File di riferimento</h2>')
    html_parts.append('<ul>')
    html_parts.append('<li><strong>Dati dettaglio</strong>: <code>flussi_cassa_dettaglio.csv</code></li>')
    html_parts.append('<li><strong>Riepilogo mensile</strong>: <code>flussi_cassa_riepilogo.csv</code></li>')
    html_parts.append('<li><strong>Configurazione grafici</strong>: <code>Categorie_per_grafici.csv</code></li>')
    html_parts.append('<li><strong>Grafici</strong>: cartella <code>grafici/</code></li>')
    html_parts.append('</ul>')

    body = '\n        '.join(html_parts)
    css = _get_html_css()

    return f"""<!doctype html>
<html lang="it">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Report Flussi di Cassa</title>
    <style>{css}</style>
</head>
<body>
    <article class="report-body">
        {body}
    </article>
</body>
</html>"""


def _scrivi_report_files(md_content: str, grafici_info: list, df_config) -> None:
    """Scrive sia il report Markdown che la versione HTML nativa."""

    with open(REPORT_FILE, 'w', encoding='utf-8') as f_md:
        f_md.write(md_content)

    html_content = genera_report_html_nativo(grafici_info, df_config)
    with open(REPORT_HTML_FILE, 'w', encoding='utf-8') as f_html:
        f_html.write(html_content)


def genera_report_markdown(grafici_info, df_config):
    """Genera il report in Markdown e crea anche l'HTML equivalente."""

    print("\n📝 Generazione report Markdown + HTML...")

    lines = []

    # Intestazione
    lines.append("# 📊 Report Flussi di Cassa\n")
    lines.append(f"\n*Generato il: {datetime.now().strftime('%d/%m/%Y %H:%M')}*\n\n")
    lines.append("---\n\n")

    # Indice
    lines.append("## 📑 Indice\n\n")
    lines.append("1. [Panoramica generale](#panoramica-generale)\n")
    for i, info in enumerate(grafici_info, 2):
        anchor = info['categoria'].lower().replace(' ', '-').replace(',', '')
        lines.append(f"{i}. [{info['categoria']}](#{anchor})\n")
    lines.append("\n---\n\n")

    # Sezione panoramica con grafici generali
    lines.append("## Panoramica generale\n\n")
    lines.append("### 📈 Andamento mensile\n\n")
    lines.append("![Andamento mensile](grafici/01_andamento_mensile.png)\n\n")
    lines.append("### 💰 Categorie di spesa\n\n")
    lines.append("![Categorie spesa](grafici/02_categorie_spesa.png)\n\n")
    lines.append("---\n\n")

    # Sezioni per ogni categoria
    for info in grafici_info:
        categoria = info['categoria']
        lines.append(f"## {categoria}\n\n")

        # Statistiche
        lines.append("### 📈 Statistiche\n\n")
        lines.append("| Metrica | Valore |\n")
        lines.append("|---------|--------|\n")
        lines.append(f"| **Totale periodo** | €{info['totale']:,.2f} |\n")
        lines.append(f"| **Media mensile** | €{info['media_mensile']:,.2f} |\n")
        lines.append(f"| **Mese con spesa max** | {info['mese_max']} (€{info['max_val']:,.2f}) |\n")
        lines.append(f"| **Sottocategorie** | {len(info['sottocategorie'])} |\n\n")

        # Lista sottocategorie
        lines.append("### 📋 Sottocategorie incluse\n\n")
        for sotto in info['sottocategorie']:
            lines.append(f"- {sotto}\n")
        lines.append("\n")

        # Grafico
        lines.append("### 📊 Grafico\n\n")
        lines.append(f"![{categoria}](grafici/{info['filename']})\n\n")
        lines.append("---\n\n")

    # Footer
    lines.append("## 📁 File di riferimento\n\n")
    lines.append("- **Dati dettaglio**: `flussi_cassa_dettaglio.csv`\n")
    lines.append("- **Riepilogo mensile**: `flussi_cassa_riepilogo.csv`\n")
    lines.append("- **Configurazione grafici**: `Categorie_per_grafici.csv`\n")
    lines.append("- **Grafici**: cartella `grafici/`\n")

    md_content = "".join(lines)

    # Salva entrambi i formati
    _scrivi_report_files(md_content, grafici_info, df_config)

    print(f"   💾 Report MD: {REPORT_FILE}")
    print(f"   💾 Report HTML: {REPORT_HTML_FILE}")


def main():
    """Funzione principale."""
    
    print("=" * 70)
    print("GENERAZIONE REPORT PERSONALIZZATO")
    print("=" * 70)
    
    # Verifica esistenza file configurazione
    if not os.path.exists(CSV_CATEGORIE):
        print(f"\n❌ File non trovato: {CSV_CATEGORIE}")
        print("   Crea il file con le colonne: Categoria,Sottocategoria")
        return
    
    # Carica dati
    df_dettaglio = carica_dati()
    df_config = carica_categorie_config()

    # Crea la cartella grafici se manca (serve anche per le immagini referenziate nell'HTML)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Genera grafici aggregati
    grafici_info = genera_grafici_aggregati(df_dettaglio, df_config)
    
    # Genera report Markdown
    genera_report_markdown(grafici_info, df_config)
    
    print("\n" + "=" * 70)
    print("✅ REPORT COMPLETATO")
    print("=" * 70)
    print(f"\n📁 Grafici generati: {len(grafici_info)}")
    print(f"📄 Report: {REPORT_FILE}")
    print(f"🌐 Report HTML: {REPORT_HTML_FILE}")


if __name__ == "__main__":
    main()
