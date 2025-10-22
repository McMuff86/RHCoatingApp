# Usage.md - Schuler Türen Datenbanksystem

## Schnellstart

### 1. Installation

```bash
pip install pandas openpyxl
```

### 2. Datenbank erstellen

```python
# create_complete_database.py ausführen
python create_complete_database.py
```

Dies erstellt:
- `Schuler_Tueren_Complete_Database.csv`
- `Schuler_Tueren_Complete_Database.xlsx`

### 3. Daten laden und analysieren

```python
import pandas as pd
from load_data import create_normalized_dataframe

# Daten laden
df = pd.read_csv('Schuler_Tueren_Complete_Database.csv')

# Normalisiertes Format für bessere Analyse
df_normalized = create_normalized_dataframe()
```

## Datenbank-Filter

### Nach Anforderung filtern

```python
# Alle Schalldämmtüren
schall_tueren = df[df['Anforderung'] == 'Schall']

# Alle Brandschutztüren
brandschutz = df[df['Anforderung'] == 'Brandschutz EI30']

# Alle Klima-Türen
klima_tueren = df[df['Anforderung'] == 'Klima']
```

### Nach Türtyp filtern

```python
# Alle VS-Türen
vs_tueren = df[df['Tuertyp'].str.contains('VS')]

# Alle LS-Türen
ls_tueren = df[df['Tuertyp'].str.contains('LS')]

# Spezifischer Typ
ei30_vs = df[df['Tuertyp'] == 'EI30 VS']
```

### Nach Dicke filtern

```python
# Dicke Türen (≥ 50mm)
dicke_tueren = df[df['Dicke_mm'] >= 50]

# Bestimmte Dicke
tueren_57mm = df[df['Dicke_mm'] == 57]
```

### Kombinierte Filter

```python
# Schalldämmende Brandschutztüren
schall_brandschutz = df[
    (df['Anforderung'] == 'Brandschutz EI30') &
    (df['Tuertyp'].str.contains('Schall'))
]

# Preiswerte Klima-Türen
preiswerte_klima = df[
    (df['Anforderung'] == 'Klima') &
    (df['Preis_Standarddeck_CHF_m2'] < 400)
].sort_values('Preis_Standarddeck_CHF_m2')
```

## Datenvisualisierung in Rhino

### Option 1: Eto.Forms DataGridView (Empfohlen)

```python
from rhino_table_display import show_dataframe_in_eto_form

# Komplette Datenbank anzeigen
show_dataframe_in_eto_form(df, "Schuler Türen Datenbank")

# Gefilterte Daten anzeigen
show_dataframe_in_eto_form(schall_tueren, "Schalldämmtüren")
```

### Option 2: HTML-Export

```python
from rhino_table_display import show_dataframe_as_html

# In Browser öffnen
show_dataframe_as_html(df, "Schuler Türen Übersicht")
```

### Option 3: Excel-Export

```python
from rhino_table_display import export_to_excel

# Excel-Datei erstellen und öffnen
export_to_excel(df, "meine_tueren_auswahl.xlsx")
```

## Datenanalyse

### Preisvergleich

```python
# Durchschnittspreise nach Anforderung
durchschnitt_preise = df.groupby('Anforderung')['Preis_Standarddeck_CHF_m2'].mean()

# Günstigste Option pro Kategorie
guenstigste = df.loc[df.groupby('Anforderung')['Preis_Standarddeck_CHF_m2'].idxmin()]
```

### Gewichtsanalyse

```python
# Schwerste Türen
schwerste = df.nlargest(10, 'Gewicht_Standarddeck_kg_m2')

# Durchschnittsgewicht nach Dicke
gewicht_dicke = df.groupby('Dicke_mm')['Gewicht_Standarddeck_kg_m2'].mean()
```

### U-Wert-Analyse

```python
# Beste Wärmedämmung
beste_daemmung = df.nsmallest(10, 'U_Wert_Mittelzone_W_m2K')

# U-Wert nach Anforderung
u_wert_anforderung = df.groupby('Anforderung')['U_Wert_Mittelzone_W_m2K'].mean()
```

## Datenexport

### CSV-Export

```python
# Gefilterte Daten speichern
schall_tueren.to_csv('schalldämmtüren.csv', index=False)
```

### Excel-Export mit Formatierung

```python
# Mehrere Sheets in einer Excel-Datei
with pd.ExcelWriter('tueren_analyse.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Alle Daten', index=False)
    schall_tueren.to_excel(writer, sheet_name='Schall', index=False)
    klima_tueren.to_excel(writer, sheet_name='Klima', index=False)
```

## Troubleshooting

### Häufige Probleme

**Problem**: Rhino stürzt beim Eto.Forms-Fenster ab
**Lösung**: Verwende HTML-Export als Alternative:
```python
show_dataframe_as_html(df, "Türen Daten")
```

**Problem**: Memory-Fehler bei großen Datenmengen
**Lösung**: Filtere Daten vor der Visualisierung:
```python
# Nur relevante Spalten anzeigen
anzeige_spalten = ['Anforderung', 'Tuertyp', 'Dicke_mm', 'Preis_Standarddeck_CHF_m2']
show_dataframe_in_eto_form(df[anzeige_spalten].head(50))
```

**Problem**: Pandas Import-Fehler
**Lösung**: Installiere fehlende Bibliotheken:
```bash
pip install pandas openpyxl
```

## Best Practices

### Datenpflege
- Regelmäßige Aktualisierung der Datenbank bei Preisänderungen
- Backup der CSV-Datei vor Änderungen
- Versionsverfolgung in Git

### Performance
- Filtere Daten vor der Visualisierung
- Verwende `head()` für Vorschauen
- Speichere häufige Filter als Variablen

### Erweiterung
- Neue Türtypen nach dem etablierten Schema hinzufügen
- Zusätzliche Spalten für spezifische Anforderungen
- Integration neuer Datenquellen

## Tastenkürzel und Shortcuts

### Python/Rhino Integration
```python
# Schnelle Datenbank-Initialisierung
import pandas as pd
df = pd.read_csv('Schuler_Tueren_Complete_Database.csv')

# Schnellfilter für gängige Anfragen
def schnelle_suche(anforderung=None, tuertyp=None, max_preis=None):
    result = df.copy()
    if anforderung:
        result = result[result['Anforderung'] == anforderung]
    if tuertyp:
        result = result[result['Tuertyp'].str.contains(tuertyp)]
    if max_preis:
        result = result[result['Preis_Standarddeck_CHF_m2'] <= max_preis]
    return result
```

## Support

Bei Fragen oder Problemen:
1. Prüfe die Dokumentation
2. Teste mit kleinen Datenmengen
3. Verwende die HTML-Alternative bei Rhino-Problemen
4. Überprüfe Python/Pandas-Installationen

