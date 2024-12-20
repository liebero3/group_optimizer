## Gruppeneinteilung mit ILP und Lokaler Suche

Dieses Python-Projekt ermöglicht es, Personen anhand ihrer Gruppenpräferenzen (Wünsche und No-Go) passenden Gruppen zuzuordnen, wobei die Kapazitäten der einzelnen Gruppen nicht überschritten werden dürfen. Unter der Haube werden die **OR-Tools** von Google verwendet, um das Problem als Integer Lineares Programm (ILP) zu lösen. Im Anschluss kommt eine lokale Suche zum Einsatz, um die ILP-Lösung weiter zu verbessern.

---

### Inhalt
- [Gruppeneinteilung mit ILP und Lokaler Suche](#gruppeneinteilung-mit-ilp-und-lokaler-suche)
  - [Inhalt](#inhalt)
- [Funktionsumfang](#funktionsumfang)
- [Voraussetzungen](#voraussetzungen)
- [Installation](#installation)
- [Eingabedateien](#eingabedateien)
  - [1) capacities.xlsx](#1-capacitiesxlsx)
  - [2) preferences.xlsx](#2-preferencesxlsx)
- [Benutzung](#benutzung)
  - [Parameter](#parameter)
- [Beispielaufruf](#beispielaufruf)
- [Erweiterungsmöglichkeiten](#erweiterungsmöglichkeiten)

---

## Funktionsumfang
- **Einlesen von Gruppenkapazitäten** aus einer Excel-Datei (`capacities.xlsx`)
- **Einlesen von Präferenzen** aus einer zweiten Excel-Datei (`preferences.xlsx`), die Wünsche (W1 bis W10) und No-Go-Gruppen (N1, N2, N3) für jede Person enthalten kann
- **ILP-Lösung** mit Zeitlimit (default 60 Sekunden) zur Berechnung einer möglichst guten Zuweisung
- **Lokale Suche** (randomisierte Tauschoperationen) für eine zusätzliche Optimierung innerhalb eines selbst festlegbaren Zeitfensters (default 30 Sekunden)
- **Ausgabe des Scores** und Auflistung der finalen Zuweisung pro Gruppe, sowohl als CSV (`ergebnis.csv`) als auch als Excel-Datei (`ergebnis.xlsx`)
- **Zusätzliche Analyse** zu erfüllten Wünschen pro Person (W1..W10) und Anzahl an Personen, die keinen ihrer Wünsche bekommen haben

---

## Voraussetzungen
- **Python 3.x**
- Folgende Python-Pakete werden benötigt:
  - `pandas`
  - `numpy`
  - `openpyxl` (für das Auslesen von Excel-Dateien, in manchen pandas-Versionen bereits enthalten)
  - `ortools` (Python-Wrapper für Google OR-Tools)

---

## Installation
1. **Python installieren:** Stelle sicher, dass mindestens Python 3.7 oder höher installiert ist.
2. **Abhängigkeiten installieren:**
   ```bash
   pip install pandas numpy openpyxl ortools
   ```
3. **Code herunterladen/klonen** (entweder als ZIP oder via `git clone` des GitHub-Repositories).

---

## Eingabedateien

### 1) capacities.xlsx
Erwartetes Format (mindestens die ersten zwei Zeilen sind relevant):
- **Erste Zeile:** Namen der Gruppen (z.B. G1, G2, G3 oder „Gruppe A“, „Gruppe B“, …)
- **Zweite Zeile:** Kapazitätsgrenzen jeder entsprechenden Gruppe

Beispielsweise könnte die Datei so aussehen:

|       |   0      |   1      |   2       |
|-------|---------:|---------:|----------:|
| 0     | Gruppe A | Gruppe B | Gruppe C  |
| 1     | 10       | 12       | 8         |

### 2) preferences.xlsx
Erwartetes Format:
- **Spalte "Name":** Name jeder Person
- **Spalten "W1", "W2", ..., "W10":** Wunsch-Gruppen in absteigender Priorität (W1 > W2 > … > W10). Werden nur so viele Spalten genutzt, wie man wirklich benötigt.
- **Spalten "N1", "N2", "N3":** No-Go-Gruppen (max. drei No-Go-Gruppen pro Person).

Beispiel (mit nur W1 bis W3 und N1 bis N3):

| Name  | W1       | W2       | W3       | N1       | N2       | N3       |
|-------|----------|----------|----------|----------|----------|----------|
| Max   | Gruppe A | Gruppe B | nan      | Gruppe C | nan      | nan      |
| Julia | Gruppe B | nan      | nan      | Gruppe A | Gruppe C | nan      |
| Anne  | Gruppe C | Gruppe A | Gruppe B | nan      | nan      | nan      |

Wobei „nan“ bedeutet, dass kein Wert angegeben ist (z.B. wenn es nur einen oder zwei Wünsche/No-Go gibt).

**Gewichtungen (Beispiel, im Code festgelegt):**
- W1: +90 Punkte
- W2: +30 Punkte
- W3: +10 Punkte
- W4: +5 Punkte
- W5: +4 Punkte
- W6: +3 Punkte
- W7: +2 Punkte
- W8: +1 Punkt
- W9: +1 Punkt
- W10: +1 Punkt
- No-Go (N1, N2, N3): -10 Punkte

---

## Benutzung

1. **Terminal öffnen** (in dem Ordner, in dem sich die Python-Skripte befinden).
2. **Skript ausführen** mit:

   ```bash
   python gruppeneinteilung.py --prefs <Pfad_zur_preferences.xlsx> --caps <Pfad_zur_capacities.xlsx> \
       --mip_time <Zeit_in_Sekunden> --local_time <Zeit_in_Sekunden>
   ```

### Parameter
- `--prefs`: Pfad zur `preferences.xlsx`
- `--caps`: Pfad zur `capacities.xlsx`
- `--mip_time` (optional, default 60): Zeitlimit in Sekunden für die ILP-Berechnung
- `--local_time` (optional, default 30): Zeitlimit in Sekunden für die lokale Suche

---

## Beispielaufruf

```bash
python gruppeneinteilung.py \
    --prefs preferences.xlsx \
    --caps capacities.xlsx \
    --mip_time 120 \
    --local_time 60
```

Ablauf:
1. **Einlesen der Gruppenkapazitäten** aus `capacities.xlsx`
2. **Einlesen der Präferenzen** (W1..W10 und N1..N3) aus `preferences.xlsx`
3. **Durchführung des ILP-Solvers** (max. 120 Sekunden)
4. **Lokale Suche** (nochmal 60 Sekunden, zufällige Tauschoperationen)
5. **Ausgabe** von:
   - Insgesamt erzielter Score
   - Finaler Gruppenaufteilung (pro Gruppe)
   - Anzahl erfüllter Wünsche pro Kategorie (W1..W10)
   - Anzahl Personen, die keinen Wunsch bekommen haben
   - Speicher der Ergebnisse in `ergebnis.csv` und `ergebnis.xlsx`

---

## Erweiterungsmöglichkeiten
- **Weitere Wunsch-Ränge**: Das Programm ist bereits vorbereitet bis W10. Falls weitere benötigt werden, kann das Mapping in `wish_weights` erweitert bzw. angepasst werden. 
- **Variation der Gewichte**: Die Gewichtungen für Wunschrang und No-Go sind im Code änderbar.
- **Mehrfache Zufallsstarts**: Die lokale Suche könnte mehrfach mit unterschiedlichen Startlösungen laufen, um noch bessere Ergebnisse zu erzielen.
- **Weitere lokale Suchroutinen**: Z.B. Annealing, Tabu Search oder genetische Algorithmen, um eine umfassendere Optimierung durchzuführen.
- **GUI**: Eine grafische Oberfläche könnte das Einlesen und Ausgeben vereinfachen.

---

Viel Erfolg bei der Gruppeneinteilung!
Bei Fragen oder Anmerkungen gerne ein Issue im GitHub-Repository erstellen oder einen Pull-Request mit Verbesserungen einreichen.