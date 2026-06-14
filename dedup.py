import os
import sys
import json
from difflib import SequenceMatcher

import pandas as pd


EXCEL_EXT = (".xlsx", ".xlsm", ".xls")
APP_CONFIG_FILE = "dedup_config.json"

# Etichette di stato (usate anche come chiavi nei controlli successivi)
STATO_NON_DOPPIA = "non doppia"
STATO_SICURA = "sicuramente doppia"
STATO_PROBABILE = "probabilmente doppia"

# Chiavi univoche: se due righe le condividono sono la stessa persona,
# a prescindere da come è scritto il nome (es. Codice Fiscale).
STRONG_KEYS = ("cf",)
# Chiavi deboli: possono essere condivise (es. telefono di famiglia, email
# di coppia), quindi richiedono anche l'accordo su nome/cognome.
WEAK_KEYS = ("email", "cell", "tel")

DEFAULT_CONFIG = {
    "dedup_url_template": (
        "https://dc.directchannel.it/nonprofit/nomi_merge.asp"
        "?id_cedente={id_cedente}&id_vincente={id_vincente}"
    ),
    "dedup_link_label": "Deduplica",
    "saved_mapping_file": ".dedup_last_mapping.json",
    "name_similarity_threshold": 0.8,
}


def find_excel_files(path="."):
    files = []
    for name in os.listdir(path):
        if name.startswith("~$"):
            # file temporanei di Excel
            continue
        if name.lower().endswith(EXCEL_EXT):
            files.append(os.path.join(path, name))
    return files


def input_int(prompt, min_val=None, max_val=None, allow_zero=False):
    while True:
        raw = input(prompt).strip()
        try:
            value = int(raw)
        except ValueError:
            print("Per favore inserisci un numero intero.")
            continue

        if value == 0 and allow_zero:
            return 0

        if min_val is not None and value < min_val:
            print(f"Per favore inserisci un numero >= {min_val}.")
            continue
        if max_val is not None and value > max_val:
            print(f"Per favore inserisci un numero <= {max_val}.")
            continue
        return value


def input_yes_no(prompt, default=True):
    suffix = "[S/n]" if default else "[s/N]"
    while True:
        raw = input(f"{prompt} {suffix}: ").strip().lower()
        if not raw:
            return default
        if raw in {"s", "si", "sì", "y", "yes"}:
            return True
        if raw in {"n", "no"}:
            return False
        print("Per favore rispondi s oppure n.")


def choose_from_list(options, message):
    """
    Stampa un elenco numerato e fa scegliere un elemento.
    Restituisce l'indice (0-based).
    """
    if not options:
        raise ValueError("La lista è vuota, non puoi scegliere nulla.")

    print(message)
    for i, opt in enumerate(options, start=1):
        print(f"{i}. {opt}")

    choice = input_int(
        f"Seleziona un numero tra 1 e {len(options)}: ",
        min_val=1,
        max_val=len(options),
    )
    return choice - 1


def load_app_config():
    config = DEFAULT_CONFIG.copy()
    path = os.path.join(os.getcwd(), APP_CONFIG_FILE)

    if not os.path.isfile(path):
        return config

    try:
        with open(path, "r", encoding="utf-8") as fh:
            user_config = json.load(fh)
    except (OSError, json.JSONDecodeError):
        print(f"Attenzione: non riesco a leggere {APP_CONFIG_FILE}, userò i valori di default.")
        return config

    if not isinstance(user_config, dict):
        print(f"Attenzione: {APP_CONFIG_FILE} non contiene un oggetto JSON valido, userò i default.")
        return config

    for key in DEFAULT_CONFIG:
        if key in user_config:
            config[key] = user_config[key]

    # La soglia deve essere un numero in [0, 1]: la validiamo subito per non
    # far fallire l'elaborazione a metà (dopo la mappatura interattiva).
    try:
        threshold = float(config["name_similarity_threshold"])
    except (TypeError, ValueError):
        threshold = None
    if threshold is None or not 0.0 <= threshold <= 1.0:
        print(
            f"Attenzione: 'name_similarity_threshold' non valido in {APP_CONFIG_FILE}, "
            f"uso il default {DEFAULT_CONFIG['name_similarity_threshold']}."
        )
        config["name_similarity_threshold"] = DEFAULT_CONFIG["name_similarity_threshold"]
    else:
        config["name_similarity_threshold"] = threshold

    return config


def get_mapping_config_path(config):
    return os.path.join(os.getcwd(), str(config["saved_mapping_file"]))


def load_saved_mapping(config):
    path = get_mapping_config_path(config)
    if not os.path.isfile(path):
        return None

    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, json.JSONDecodeError):
        return None

    if not isinstance(data, dict):
        return None
    if not isinstance(data.get("mapping"), dict):
        return None
    return data


def save_mapping(mapping, source_file, sheet_name, df_columns, config):
    payload = {
        "source_file": os.path.basename(source_file),
        "sheet_name": sheet_name,
        "columns": list(df_columns),
        "mapping": mapping,
    }
    path = get_mapping_config_path(config)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def is_saved_mapping_compatible(saved_data, df):
    saved_columns = saved_data.get("columns")
    mapping = saved_data.get("mapping", {})

    if not isinstance(saved_columns, list):
        return False
    if list(df.columns) != saved_columns:
        return False

    for value in mapping.values():
        cols = value if isinstance(value, list) else [value]
        for col_name in cols:
            if col_name is not None and col_name not in df.columns:
                return False
    return True


def print_mapping(mapping, title):
    print(title)
    for k, v in mapping.items():
        shown = ", ".join(v) if isinstance(v, list) else v
        print(f"  {k}: {shown}")
    print()


def normalize_string(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def normalize_phone(s):
    if pd.isna(s):
        return ""
    digits = "".join(ch for ch in str(s) if ch.isdigit())
    # Rimuove il prefisso internazionale italiano per far combaciare
    # "+39 333..." / "0039333..." con il numero in formato nazionale.
    if digits.startswith("0039"):
        digits = digits[4:]
    elif digits.startswith("39") and len(digits) > 10:
        digits = digits[2:]
    return digits


def similarity(a, b):
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def normalize_id_for_compare(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def id_sort_key(value):
    value = normalize_id_for_compare(value)
    if not value:
        return (2, "")
    if value.isdigit():
        return (0, int(value))
    return (1, value.lower())


def escape_excel_string(value):
    # In una formula Excel le virgolette dentro una stringa si raddoppiano.
    return str(value).replace('"', '""')


def build_dedup_link(current_id, duplicate_id, config):
    current_id = normalize_id_for_compare(current_id)
    duplicate_id = normalize_id_for_compare(duplicate_id)

    if not current_id or not duplicate_id:
        return ""

    winner_id, loser_id = sorted([current_id, duplicate_id], key=id_sort_key)
    url = str(config["dedup_url_template"]).format(
        id_cedente=loser_id,
        id_vincente=winner_id,
    )
    label = str(config["dedup_link_label"])
    return f'=HYPERLINK("{escape_excel_string(url)}","{escape_excel_string(label)}")'


def build_dedup_links_columns(row_ids, duplicate_ids, states, config):
    all_links = []

    for current_id, duplicates_raw, state in zip(row_ids, duplicate_ids, states):
        if state not in {STATO_SICURA, STATO_PROBABILE} or not duplicates_raw:
            all_links.append([])
            continue

        duplicates = [
            normalize_id_for_compare(item)
            for item in str(duplicates_raw).split(",")
        ]
        duplicates = [item for item in duplicates if item]

        links = [
            build_dedup_link(current_id, duplicate_id, config)
            for duplicate_id in duplicates
        ]
        links = [link for link in links if link]
        all_links.append(links)

    max_links = max((len(links) for links in all_links), default=0)
    columns = {}

    for idx in range(max_links):
        columns[f"Link_dedup_{idx + 1}"] = [
            links[idx] if idx < len(links) else ""
            for links in all_links
        ]

    return columns


def map_columns(df, config, source_file=None, sheet_name=None):
    print("\nColonne trovate nel file:")
    for idx, col in enumerate(df.columns, start=1):
        print(f"{idx}. {col}")

    saved_data = load_saved_mapping(config)
    if saved_data and is_saved_mapping_compatible(saved_data, df):
        saved_mapping = saved_data["mapping"]
        source_label = saved_data.get("source_file") or "file precedente"
        sheet_label = saved_data.get("sheet_name") or "foglio precedente"
        print_mapping(
            saved_mapping,
            (
                f"\nHo trovato una mappatura salvata compatibile "
                f"({source_label} / {sheet_label}):"
            ),
        )
        if input_yes_no("Vuoi riutilizzare questa mappatura?", default=True):
            return saved_mapping

    print("\nMappatura colonne (inserisci il numero della colonna; 0 se il campo NON esiste):")

    id_idx = input_int("Colonna ID univoco (obbligatorio): ", min_val=1, max_val=len(df.columns))
    id_col = df.columns[id_idx - 1]

    email_idx = input_int("Colonna Email (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    email2_idx = input_int("Colonna Email secondaria (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    cf_idx = input_int("Colonna Codice Fiscale (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    cell_idx = input_int("Colonna Cellulare (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    cell2_idx = input_int("Colonna Cellulare secondario (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    tel_idx = input_int("Colonna Telefono fisso (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    nome_idx = input_int("Colonna Nome (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    cognome_idx = input_int("Colonna Cognome (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)

    def col_or_none(idx):
        if idx == 0:
            return None
        return df.columns[idx - 1]

    def cols_list(*idxs):
        # Raccoglie le colonne indicate (saltando 0 e i duplicati) in una lista.
        cols = []
        for idx in idxs:
            col = col_or_none(idx)
            if col and col not in cols:
                cols.append(col)
        return cols

    mapping = {
        "id": id_col,
        "email": cols_list(email_idx, email2_idx),
        "cf": col_or_none(cf_idx),
        "cell": cols_list(cell_idx, cell2_idx),
        "tel": col_or_none(tel_idx),
        "nome": col_or_none(nome_idx),
        "cognome": col_or_none(cognome_idx),
    }

    print_mapping(mapping, "\nMappatura confermata:")

    if source_file and sheet_name:
        save_mapping(
            mapping,
            source_file=source_file,
            sheet_name=sheet_name,
            df_columns=df.columns,
            config=config,
        )
        print(f"Mappatura salvata in {get_mapping_config_path(config)}\n")

    return mapping


# Funzione di normalizzazione per i campi a valore singolo.
SINGLE_NORMALIZERS = {
    "cf": lambda v: normalize_string(v).upper(),
    "tel": normalize_phone,
    "nome": normalize_string,
    "cognome": normalize_string,
}

# Campi a valore multiplo: piu colonne confluiscono nello stesso "spazio"
# (es. Email + Email2). Una riga porta l'insieme dei suoi valori e due righe
# combaciano se ne condividono uno qualunque, anche incrociato.
MULTI_NORMALIZERS = {
    "email": normalize_string,
    "cell": normalize_phone,
}


def build_normalized_columns(df, mapping):
    """
    Crea colonne normalizzate in un dict separato, per comodità.
    Restituisce un dict con chiavi: id, email, cf, cell, tel, nome, cognome.
    I campi a valore singolo sono liste di stringhe (una per riga); i campi a
    valore multiplo (email, cell) sono liste di liste (i valori di ogni riga).
    """
    n = len(df)
    norm = {}

    # ID: lo manteniamo così com'è (stringa)
    norm["id"] = ["" if pd.isna(v) else str(v) for v in df[mapping["id"]].values]

    for key, normalizer in SINGLE_NORMALIZERS.items():
        col = mapping.get(key)
        if col:
            norm[key] = [normalizer(v) for v in df[col].values]
        else:
            norm[key] = [""] * n

    for key, normalizer in MULTI_NORMALIZERS.items():
        cols = mapping.get(key) or []
        if isinstance(cols, str):  # compatibilità con mappature salvate vecchie
            cols = [cols]
        per_col = [[normalizer(v) for v in df[col].values] for col in cols]
        # Per ogni riga: lista dei valori normalizzati non vuoti dalle sue colonne.
        norm[key] = [
            [values[i] for values in per_col if values[i]]
            for i in range(n)
        ]

    return norm


def collect_groups_by_key(values):
    """
    values: una voce per riga. Ogni voce è una stringa (campo singolo) oppure
    una lista di stringhe (campo multi-valore, es. Email + Email2).
    Ritorna dict: valore -> lista di indici (righe) in cui compare, solo se
    valore non vuoto. Una riga compare una sola volta per ciascun valore.
    """
    groups = {}
    for idx, val in enumerate(values):
        row_values = val if isinstance(val, (list, tuple, set)) else (val,)
        for v in set(row_values):
            if not v:
                continue
            groups.setdefault(v, []).append(idx)
    # teniamo solo gruppi con almeno 2 righe
    groups = {k: v for k, v in groups.items() if len(v) > 1}
    return groups


def names_match_similar(nome1, cog1, nome2, cog2, threshold):
    # Similarità approssimata (solo se abbiamo entrambi nome e cognome).
    sim_nome = similarity(nome1, nome2) if nome1 and nome2 else 0.0
    sim_cog = similarity(cog1, cog2) if cog1 and cog2 else 0.0
    return sim_nome >= threshold and sim_cog >= threshold


def deduplicate(norm, config):
    """
    norm: dict con chiavi id,email,cf,cell,tel,nome,cognome
    Ritorna:
      stato: lista con valori "non doppia" / "sicuramente doppia" / "probabilmente doppia"
      id_doppi: lista di stringhe con gli ID delle righe duplicate
    """
    n = len(norm["id"])
    sure_dups = [set() for _ in range(n)]
    prob_dups = [set() for _ in range(n)]

    # soglia di similarità per "probabilmente doppia"
    name_threshold = float(config["name_similarity_threshold"])

    # Chiavi univoche: condividere il valore basta per "sicuramente doppia".
    for key in STRONG_KEYS:
        groups = collect_groups_by_key(norm[key])
        for indices in groups.values():
            for i in range(len(indices)):
                for j in range(i + 1, len(indices)):
                    r1, r2 = indices[i], indices[j]
                    sure_dups[r1].add(norm["id"][r2])
                    sure_dups[r2].add(norm["id"][r1])

    # Chiavi deboli: servono anche nome e cognome coincidenti (o simili).
    # Per non confrontare ogni coppia di righe (O(k^2) di SequenceMatcher sui
    # gruppi grandi, es. un telefono fisso condiviso da molti), raggruppiamo
    # prima per (nome, cognome) normalizzati: il match esatto diventa O(k) e
    # la similarità si calcola solo tra i bucket distinti (di norma pochi).
    for key in WEAK_KEYS:
        groups = collect_groups_by_key(norm[key])
        for indices in groups.values():
            buckets = {}
            for idx in indices:
                name_key = (norm["nome"][idx], norm["cognome"][idx])
                buckets.setdefault(name_key, []).append(idx)

            # Match esatto: stesso nome e cognome (entrambi presenti).
            for (nome, cog), rows in buckets.items():
                if not (nome and cog) or len(rows) < 2:
                    continue
                for i in range(len(rows)):
                    for j in range(i + 1, len(rows)):
                        r1, r2 = rows[i], rows[j]
                        sure_dups[r1].add(norm["id"][r2])
                        sure_dups[r2].add(norm["id"][r1])

            # Match simile: confronto solo tra bucket distinti.
            bucket_items = list(buckets.items())
            for a in range(len(bucket_items)):
                (nome1, cog1), rows1 = bucket_items[a]
                for b in range(a + 1, len(bucket_items)):
                    (nome2, cog2), rows2 = bucket_items[b]
                    if not names_match_similar(nome1, cog1, nome2, cog2, name_threshold):
                        continue
                    for r1 in rows1:
                        for r2 in rows2:
                            prob_dups[r1].add(norm["id"][r2])
                            prob_dups[r2].add(norm["id"][r1])

    stato = []
    id_doppi = []

    for i in range(n):
        # se esistono duplicati sicuri, i "probabili" su quelle stesse righe li ignoriamo
        if sure_dups[i]:
            label = STATO_SICURA
            # un id può finire sia tra i sicuri che tra i probabili: evitiamo doppioni
            ids = sorted(sure_dups[i], key=id_sort_key)
        elif prob_dups[i]:
            label = STATO_PROBABILE
            ids = sorted(prob_dups[i], key=id_sort_key)
        else:
            label = STATO_NON_DOPPIA
            ids = []

        stato.append(label)
        id_doppi.append(", ".join(ids))

    return stato, id_doppi


def main():
    print("=== Deduplica Excel (anagrafiche) ===\n")
    config = load_app_config()

    # 1. Cerca file Excel nella cartella corrente
    current_path = os.getcwd()
    files = find_excel_files(current_path)

    if not files:
        print("Nessun file Excel trovato nella cartella corrente.")
        new_path = input("Inserisci il percorso di una cartella da cercare: ").strip()
        if not new_path:
            print("Nessun percorso inserito. Esco.")
            sys.exit(1)
        if not os.path.isdir(new_path):
            print("Il percorso indicato non è una cartella valida. Esco.")
            sys.exit(1)
        files = find_excel_files(new_path)
        if not files:
            print("Nessun file Excel trovato nella cartella indicata. Esco.")
            sys.exit(1)

    # 2. Scegli il file
    file_idx = choose_from_list(
        [os.path.basename(f) for f in files],
        "File Excel disponibili:"
    )
    excel_path = files[file_idx]
    print(f"\nHai scelto: {excel_path}\n")

    # 3. Scegli il foglio
    xls = pd.ExcelFile(excel_path)
    sheet_names = xls.sheet_names

    if len(sheet_names) == 1:
        sheet_name = sheet_names[0]
        print(f"C'è un solo foglio: userò '{sheet_name}'.\n")
    else:
        sheet_idx = choose_from_list(
            sheet_names,
            "Fogli disponibili nel file:"
        )
        sheet_name = sheet_names[sheet_idx]
        print(f"\nHai scelto il foglio: {sheet_name}\n")

    # 4. Carica il foglio in DataFrame
    df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)

    # 5. Mappatura colonne
    mapping = map_columns(df, config, source_file=excel_path, sheet_name=sheet_name)

    # 6. Normalizzazione valori
    norm = build_normalized_columns(df, mapping)

    # 7. Deduplica
    print("Eseguo la deduplica, potrebbe volerci qualche istante...\n")
    stato, id_doppi = deduplicate(norm, config)

    # 8. Aggiungo le colonne al DataFrame
    df["Stato_dedup"] = stato
    df["ID_doppi"] = id_doppi
    dedup_link_columns = build_dedup_links_columns(norm["id"], id_doppi, stato, config)
    for col_name, values in dedup_link_columns.items():
        df[col_name] = values

    # 9. Salvo il risultato
    base, ext = os.path.splitext(excel_path)
    out_path = f"{base}_dedup.xlsx"

    try:
        df.to_excel(out_path, sheet_name=sheet_name, index=False)
    except PermissionError:
        print(
            f"Impossibile scrivere {out_path}: il file è aperto in Excel o non hai i permessi.\n"
            "Chiudilo e riprova."
        )
        sys.exit(1)
    print(f"Deduplica completata.\nFile salvato come:\n{out_path}")


if __name__ == "__main__":
    main()
