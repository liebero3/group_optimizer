import argparse
import pandas as pd
import numpy as np
import time
from ortools.linear_solver import pywraplp


def read_capacities(path):
    """
    Liest die Kapazitäten-Excel-Datei und gibt zwei Listen zurück.

    Die Excel-Datei sollte folgendes Format haben:
      - In der ersten Zeile: Namen der Gruppen (z.B. G1, G2, ...)
      - In der zweiten Zeile: Die zugehörigen Kapazitäten (z.B. 10, 12, ...)

    Parameter
    ----------
    path : str
        Pfad zur Excel-Datei, die die Kapazitäten enthält.

    Returns
    -------
    group_names : list of str
        Liste mit den Namen der Gruppen (z.B. ["Gruppe A", "Gruppe B", ...]).
    group_capacities : list of int
        Liste mit den zugehörigen Kapazitäten für jede Gruppe.
    """
    df = pd.read_excel(path, header=None, nrows=2)

    # Erste Zeile enthält die Namen der Gruppen
    group_names = df.iloc[0].tolist()

    # Zweite Zeile enthält die Kapazitäten der Gruppen
    group_capacities = df.iloc[1].tolist()

    return group_names, group_capacities


def read_preferences(path, group_names):
    """
    Liest die Präferenzen aus einer Excel-Datei und gibt die Personen und ihre Präferenzen zurück.

    Erwartetes Format in der Excel-Datei (Spalten):
      - Name (Name der Person)
      - W1, W2, W3 (Wunsch-Gruppen in absteigender Priorität)
      - N1, N2, N3 (No-Go-Gruppen)

    Beispiel-Struktur:
      | Name     | W1       | W2       | W3       | N1       | N2       | N3       |
      |----------|----------|----------|----------|----------|----------|----------|
      | Max      | Gruppe A | Gruppe B | nan      | Gruppe C | nan      | nan      |
      | Julia    | Gruppe B | nan      | nan      | Gruppe A | Gruppe C | nan      |
      | ...      | ...      | ...      | ...      | ...      | ...      | ...      |

    Parameter
    ----------
    path : str
        Pfad zur Excel-Datei, die die Präferenzen enthält.
    group_names : list of str
        Liste der verfügbaren Gruppen (muss zu capacities.xlsx passen).

    Returns
    -------
    persons : pd.Series
        Liste (Pandas Series) mit den Namen aller Personen.
    preferences : np.array
        2D-Array der Größe [Anzahl_Personen, Anzahl_Gruppen],
        das die Präferenzwerte jeder Person für jede Gruppe enthält.

    Hinweise
    --------
    - W1 wird mit +30, W2 mit +10 und W3 mit +3 gewichtet.
    - N1, N2, N3 werden mit -10 gewichtet (No-Go).
    - Gruppen, die nicht in der Liste der gültigen Gruppen (group_names) auftauchen,
      werden als ungültig erkannt und nur mit einer Warnmeldung berücksichtigt.
    """
    data = pd.read_excel(path)
    num_persons = data.shape[0]
    num_groups = len(group_names)

    # Präferenzmatrix initialisieren
    preferences = np.zeros((num_persons, num_groups))

    # Gewichtungen für Wünsche und No-Go
    wish_weights = {"W1": 30, "W2": 10, "W3": 3}
    no_go_categories = ["N1", "N2", "N3"]
    no_go_weight = -10

    # Fülle die Präferenzmatrix
    for p in range(num_persons):
        # Positive Wünsche
        for wish_col, weight in wish_weights.items():
            group_name = data[wish_col].iloc[p]
            if pd.notna(group_name):
                try:
                    g_idx = group_names.index(group_name)
                    preferences[p, g_idx] += weight
                except ValueError:
                    # Falls eine Gruppe im Excel steht, die nicht in capacities.xlsx existiert
                    print(
                        f"[WARN] Person {data['Name'].iloc[p]} hat ungültige Wunsch-Gruppe: {group_name}"
                    )

        # Negative Wünsche (No-Gos)
        for no_col in no_go_categories:
            group_name = data[no_col].iloc[p]
            if pd.notna(group_name):
                try:
                    g_idx = group_names.index(group_name)
                    preferences[p, g_idx] += no_go_weight
                except ValueError:
                    print(
                        f"[WARN] Person {data['Name'].iloc[p]} hat ungültige No-Go-Gruppe: {group_name}"
                    )

    return data["Name"], preferences


def solve_ilp_with_ortools(
    persons, group_names, group_capacities, preferences, time_limit=60
):
    """
    Löst das Gruppeneinteilungsproblem als Integer Lineares Programm (ILP) mithilfe von OR-Tools.

    Das Problem:
    1) Jede Person soll genau einer der verfügbaren Gruppen zugewiesen werden.
    2) Die Kapazitätsgrenzen der Gruppen dürfen nicht überschritten werden.
    3) Die Summe der Präferenzen (basierend auf Wünschen und No-Go) soll maximiert werden.

    Parameter
    ----------
    persons : pd.Series
        Liste der zuzuweisenden Personen.
    group_names : list of str
        Liste der verfügbaren Gruppen.
    group_capacities : list of int
        Kapazitätsgrenzen pro Gruppe.
    preferences : np.array
        2D-Array, das für jede Person und jede Gruppe die Präferenzwerte enthält.
    time_limit : int, optional
        Zeitlimit (in Sekunden) für den ILP-Solver. Standardwert ist 60 Sekunden.

    Returns
    -------
    best_score : float
        Der maximale Gesamt-Score (Summe aller Präferenzen), falls eine Lösung gefunden wurde.
    assignment : list of int
        Liste der Länge 'Anzahl_Personen', wobei assignment[p] den Index der Gruppe für Person p angibt.
        Gibt None, None zurück, wenn keine zulässige Lösung gefunden wurde.

    Hinweise
    --------
    - Der Solver wird im SCIP-Modus von OR-Tools gestartet.
    - Das Zeitlimit wird in Millisekunden an den Solver übergeben.
    - Der Rückgabestatus kann sein:
        * OPTIMAL: Eine optimale Lösung wurde gefunden.
        * FEASIBLE: Es wurde eine Lösung gefunden, die aber nicht garantiert optimal ist.
        * INFEASIBLE: Das Modell ist unlösbar.
        * UNBOUNDED: Das Modell ist unbeschränkt.
        * Andere Werte: Keine Lösung gefunden oder Suche abgebrochen.
    """
    num_persons = len(persons)
    num_groups = len(group_names)

    # 1) Solver anlegen
    solver = pywraplp.Solver.CreateSolver("SCIP")
    if not solver:
        raise RuntimeError("Kein gültiger OR-Tools-Solver verfügbar.")

    # Zeitlimit in Millisekunden setzen
    solver.SetTimeLimit(int(time_limit * 1000))

    # 2) Binäre Variablen x[p,g] in {0,1} für die Zuordnung von Person p zu Gruppe g
    x = {}
    for p in range(num_persons):
        for g in range(num_groups):
            x[(p, g)] = solver.BoolVar(name=f"x_{p}_{g}")

    # 3) Constraints definieren
    # (A) Jede Person muss genau in einer Gruppe sein
    for p in range(num_persons):
        solver.Add(sum(x[(p, g)] for g in range(num_groups)) == 1)

    # (B) Kapazitätsgrenze: Anzahl Personen in Gruppe g darf deren Kapazität nicht überschreiten
    for g in range(num_groups):
        solver.Add(sum(x[(p, g)] for p in range(num_persons)) <= group_capacities[g])

    # 4) Zielfunktion: Summe der Präferenzen maximieren
    objective = solver.Objective()
    for p in range(num_persons):
        for g in range(num_groups):
            objective.SetCoefficient(x[(p, g)], preferences[p, g])
    objective.SetMaximization()

    # 5) Lösen
    status = solver.Solve()

    # --- Debug-Infos ausgeben ---
    used_time_s = solver.wall_time() / 1000.0
    print(f"[DEBUG] Solver-Status: {status}")
    print(f"[DEBUG] Benötigte Zeit: {used_time_s:.2f} Sekunden")
    print(f"[DEBUG] Branch-and-Bound-Knoten: {solver.nodes()}")

    # Status interpretieren
    if status == pywraplp.Solver.OPTIMAL:
        print("[INFO] Der Solver hat eine optimale Lösung gefunden.")
    elif status == pywraplp.Solver.FEASIBLE:
        print(
            "[WARN] Der Solver hat eine zulässige Lösung gefunden, "
            "jedoch ist nicht bewiesen, dass sie optimal ist."
        )
    elif status == pywraplp.Solver.INFEASIBLE:
        print("[ERROR] Das Modell ist unlösbar (infeasible).")
        return None, None
    elif status == pywraplp.Solver.UNBOUNDED:
        print("[ERROR] Das Modell ist unbeschränkt (unbounded).")
        return None, None
    else:
        print("[ERROR] Keine Lösung (abgebrochen). Status:", status)
        return None, None

    # 6) Ergebnisse auslesen
    assignment = [-1] * num_persons
    for p in range(num_persons):
        for g in range(num_groups):
            if x[(p, g)].solution_value() > 0.5:
                assignment[p] = g
                break

    # Score berechnen
    best_score = sum(preferences[p, assignment[p]] for p in range(num_persons))

    return best_score, assignment


def calculate_score(assignment, preferences):
    """
    Berechnet den Gesamt-Score einer (Person->Gruppe)-Zuordnung.

    Parameter
    ----------
    assignment : list of int
        Liste, in der assignment[p] = Index der Gruppe ist, in die Person p zugewiesen wurde.
    preferences : np.array
        2D-Array der Präferenzwerte [Anzahl_Personen, Anzahl_Gruppen].

    Returns
    -------
    score : float
        Summe aller Präferenzen (Score) für das gegebene Assignment.
    """
    score = 0
    for p, g_idx in enumerate(assignment):
        score += preferences[p, g_idx]
    return score


def local_search_improvement(assignment, group_capacities, preferences, runtime=30):
    """
    Einfache lokale Suche, die versucht, die ILP-Lösung durch zufällige Tauschoperationen zu verbessern.

    Vorgehen:
      1) Es wird eine bestimmte Laufzeit (runtime) definiert.
      2) Innerhalb dieser Zeit werden wiederholt zwei zufällige Personen ausgewählt.
      3) Wenn sie unterschiedlichen Gruppen zugewiesen sind, wird versucht, sie zu tauschen.
      4) Ist danach die Kapazitätsbeschränkung noch eingehalten und der Score höher, wird der Tausch übernommen.

    Parameter
    ----------
    assignment : list of int
        Anfangszuordnung der Personen zu Gruppen (z.B. vom ILP-Löser).
    group_capacities : list of int
        Kapazitätsgrenzen der Gruppen.
    preferences : np.array
        2D-Array der Präferenzwerte [Anzahl_Personen, Anzahl_Gruppen].
    runtime : int, optional
        Maximale Suchzeit (in Sekunden) für diese lokale Verbesserungsstrategie.

    Returns
    -------
    best_score : float
        Verbesserter oder unveränderter Score nach Abschluss der lokalen Suche.
    best_assignment : list of int
        Endgültige (lokal verbesserte) Zuordnung der Personen zu Gruppen.
    """
    start_time = time.time()
    best_assignment = assignment[:]
    best_score = calculate_score(best_assignment, preferences)
    num_persons = len(assignment)
    num_groups = len(group_capacities)

    def check_capacities(ass):
        """
        Prüft, ob eine Zuordnung 'ass' die Gruppenkapazitäten nicht überschreitet.

        Parameter
        ----------
        ass : list of int
            Zuordnungsliste (Person -> Gruppenindex).

        Returns
        -------
        bool
            True, wenn alle Kapazitäten eingehalten werden, sonst False.
        """
        counts = [0] * num_groups
        for g_idx in ass:
            counts[g_idx] += 1
        for g_idx, c in enumerate(counts):
            if c > group_capacities[g_idx]:
                return False
        return True

    # Solange Zeit übrig ist, versuche zufällige Tausch-Verbesserungen
    while time.time() - start_time < runtime:
        # Wähle zufällig 2 Personen
        p1, p2 = np.random.choice(num_persons, 2, replace=False)
        if best_assignment[p1] == best_assignment[p2]:
            # Tausch macht keinen Sinn, da beide bereits in derselben Gruppe sind
            continue

        # Führe den Tausch durch (temporär)
        new_assignment = best_assignment[:]
        new_assignment[p1], new_assignment[p2] = new_assignment[p2], new_assignment[p1]

        # Prüfe Kapazitäten
        if not check_capacities(new_assignment):
            continue

        # Berechne neuen Score
        new_score = calculate_score(new_assignment, preferences)
        if new_score > best_score:
            best_assignment = new_assignment
            best_score = new_score

    return best_score, best_assignment


def combined_approach(
    persons, group_names, group_capacities, preferences, ilp_time=60, local_time=30
):
    """
    Kombinierter Ansatz aus ILP und lokaler Suche:
      1) Zuerst wird das Problem als ILP über OR-Tools gelöst (mit Zeitlimit ilp_time).
      2) Die gefundene Lösung wird anschließend mittels einer lokalen Suche weiter verbessert.

    Parameter
    ----------
    persons : pd.Series
        Liste der zuzuweisenden Personen.
    group_names : list of str
        Liste der verfügbaren Gruppen.
    group_capacities : list of int
        Kapazitätsgrenzen pro Gruppe.
    preferences : np.array
        2D-Array, das für jede Person und jede Gruppe die Präferenzwerte enthält.
    ilp_time : int, optional
        Zeitlimit (in Sekunden) für den ILP-Solver (Standard 60).
    local_time : int, optional
        Zeitlimit (in Sekunden) für die lokale Suche (Standard 30).

    Returns
    -------
    best_score : float
        Letztendlich gefundener (ggf. verbesserter) Score.
    best_assignment : list of int
        Zugehörige Zuordnung (Person -> Gruppenindex).
    """
    print(f"[INFO] Starte ILP für max. {ilp_time} Sekunden ...")
    ilp_score, ilp_assignment = solve_ilp_with_ortools(
        persons, group_names, group_capacities, preferences, time_limit=ilp_time
    )
    if ilp_assignment is None:
        print("[ERROR] ILP hat keine Lösung geliefert.")
        return

    print(f"[INFO] ILP-Lösung gefunden mit Score = {ilp_score:.2f}")

    print(f"\n[INFO] Starte lokale Suche für weitere {local_time} Sekunden ...")
    best_score, best_assignment = local_search_improvement(
        ilp_assignment, group_capacities, preferences, runtime=local_time
    )
    print(f"[INFO] Lokale Suche beendet. Bester Score = {best_score:.2f}")

    return best_score, best_assignment


def main():
    """
    Hauptfunktion zum Ausführen des Programms über die Kommandozeile.

    Beispiel-Aufruf:
        python gruppeneinteilung.py --prefs preferences.xlsx --caps capacities.xlsx
             --mip_time 60 --local_time 30

    Parameter (über argparse):
    --------------------------
    --prefs : str
        Pfad zur Excel-Datei mit den Präferenzen (preferences.xlsx).
    --caps : str
        Pfad zur Excel-Datei mit den Gruppenkapazitäten (capacities.xlsx).
    --mip_time : int, optional
        Zeitlimit für den ILP-Teil (Standard: 60 Sekunden).
    --local_time : int, optional
        Zeitlimit für die lokale Suche (Standard: 30 Sekunden).

    Ablauf:
    -------
    1) Einlesen der Gruppennamen und Kapazitäten.
    2) Einlesen der Präferenzen für jede Person.
    3) Ausführung des kombinierten ILP- und Local-Search-Ansatzes.
    4) Ausgabe der Lösung und des Scores.
    """
    parser = argparse.ArgumentParser(
        description="Kombinierte ILP- und Local-Search-Lösung für Gruppeneinteilungen."
    )
    parser.add_argument("--prefs", required=True, help="Pfad zu preferences.xlsx")
    parser.add_argument("--caps", required=True, help="Pfad zu capacities.xlsx")
    parser.add_argument(
        "--mip_time",
        type=int,
        default=60,
        help="Zeitlimit für ILP in Sekunden (Default: 60)",
    )
    parser.add_argument(
        "--local_time",
        type=int,
        default=30,
        help="Zeitlimit für lokale Suche in Sekunden (Default: 30)",
    )
    args = parser.parse_args()

    # 1) Daten einlesen
    print("[INFO] Lese Gruppenkapazitäten...")
    group_names, group_capacities = read_capacities(args.caps)
    print(f"  Gruppen = {group_names}")
    print(f"  Kapazitäten = {group_capacities}")

    print("[INFO] Lese Präferenzen...")
    persons, preferences = read_preferences(args.prefs, group_names)
    print(f"  Anzahl Personen: {len(persons)}")

    # 2) Kombinierte Lösung aus ILP und lokaler Suche
    best_score, best_assignment = combined_approach(
        persons,
        group_names,
        group_capacities,
        preferences,
        ilp_time=args.mip_time,
        local_time=args.local_time,
    )

    print("\n====================== Ergebnisse ======================")
    print(f"Endgültiger Score: {best_score:.2f}")

    # Dictionary: {Gruppenindex: [(Personenindex, Personenname), ...]}
    assignment_by_group = {}
    for i, g_idx in enumerate(best_assignment):
        person_name = persons.iloc[i]
        assignment_by_group.setdefault(g_idx, []).append((i, person_name))

    # Für jede Gruppe eine Spalte in einer DataFrame erzeugen
    result_dict = {}
    for g_idx in sorted(assignment_by_group.keys()):
        group_count = len(assignment_by_group[g_idx])
        group_capacity = group_capacities[g_idx]

        # Spaltenüberschrift z.B. "Gruppe 1 (14 von 20)"
        column_header = f"Gruppe {g_idx+1} ({group_count} von {group_capacity})"

        # Nur die Namen in eine Liste übernehmen
        participants = [person_name for (_, person_name) in assignment_by_group[g_idx]]

        # In einem Dictionary ablegen
        result_dict[column_header] = participants

    # Damit Pandas die DataFrame korrekt anlegt,
    # müssen alle Listen die gleiche Länge haben
    max_length = max(len(lst) for lst in result_dict.values())
    for header, participant_list in result_dict.items():
        # Mit leeren Strings auffüllen
        if len(participant_list) < max_length:
            participant_list.extend([""] * (max_length - len(participant_list)))

    df_result = pd.DataFrame(result_dict)

    # Speichere als CSV
    df_result.to_csv("ergebnis.csv", index=False, encoding="utf-8-sig")
    print("[INFO] Die Ergebnisse wurden in 'ergebnis.csv' gespeichert.")

    # Speichere auch als XLSX
    df_result.to_excel("ergebnis.xlsx", index=False)
    print("[INFO] Die Ergebnisse wurden in 'ergebnis.xlsx' gespeichert.")

if __name__ == "__main__":
    main()
