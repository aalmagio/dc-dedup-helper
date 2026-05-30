import os
import sys
import json
from difflib import SequenceMatcher

import pandas as pd


EXCEL_EXT = (".xlsx", ".xlsm", ".xls")
MAPPING_CONFIG_FILE = ".dedup_last_mapping.json"


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


def get_mapping_config_path():
    return os.path.join(os.getcwd(), MAPPING_CONFIG_FILE)


def load_saved_mapping():
    path = get_mapping_config_path()
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


def save_mapping(mapping, source_file, sheet_name, df_columns):
    payload = {
        "source_file": os.path.basename(source_file),
        "sheet_name": sheet_name,
        "columns": list(df_columns),
        "mapping": mapping,
    }
    path = get_mapping_config_path()
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def is_saved_mapping_compatible(saved_data, df):
    saved_columns = saved_data.get("columns")
    mapping = saved_data.get("mapping", {})

    if not isinstance(saved_columns, list):
        return False
    if list(df.columns) != saved_columns:
        return False

    for col_name in mapping.values():
        if col_name is not None and col_name not in df.columns:
            return False
    return True


def print_mapping(mapping, title):
    print(title)
    for k, v in mapping.items():
        print(f"  {k}: {v}")
    print()


def normalize_string(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def normalize_phone(s):
    if pd.isna(s):
        return ""
    s = str(s)
    digits = "".join(ch for ch in s if ch.isdigit())
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


def build_dedup_link(current_id, duplicate_id):
    current_id = normalize_id_for_compare(current_id)
    duplicate_id = normalize_id_for_compare(duplicate_id)

    if not current_id or not duplicate_id:
        return ""

    winner_id, loser_id = sorted([current_id, duplicate_id], key=id_sort_key)
    url = (
        "https://dc.directchannel.it/nonprofit/nomi_merge.asp"
        f"?id_cedente={loser_id}&id_vincente={winner_id}"
    )
    return f'=HYPERLINK("{url}","Deduplica")'


def build_dedup_links_columns(row_ids, duplicate_ids, states):
    all_links = []

    for current_id, duplicates_raw, state in zip(row_ids, duplicate_ids, states):
        if state not in {"sicuramente doppia", "probabilmente doppia"} or not duplicates_raw:
            all_links.append([])
            continue

        duplicates = [
            normalize_id_for_compare(item)
            for item in str(duplicates_raw).split(",")
        ]
        duplicates = [item for item in duplicates if item]

        links = [build_dedup_link(current_id, duplicate_id) for duplicate_id in duplicates]
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


def map_columns(df, source_file=None, sheet_name=None):
    print("\nColonne trovate nel file:")
    for idx, col in enumerate(df.columns, start=1):
        print(f"{idx}. {col}")

    saved_data = load_saved_mapping()
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
    cf_idx = input_int("Colonna Codice Fiscale (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    cell_idx = input_int("Colonna Cellulare (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    tel_idx = input_int("Colonna Telefono fisso (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    nome_idx = input_int("Colonna Nome (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)
    cognome_idx = input_int("Colonna Cognome (0 se assente): ", min_val=0, max_val=len(df.columns), allow_zero=True)

    def col_or_none(idx):
        if idx == 0:
            return None
        return df.columns[idx - 1]

    mapping = {
        "id": id_col,
        "email": col_or_none(email_idx),
        "cf": col_or_none(cf_idx),
        "cell": col_or_none(cell_idx),
        "tel": col_or_none(tel_idx),
        "nome": col_or_none(nome_idx),
        "cognome": col_or_none(cognome_idx),
    }

    print_mapping(mapping, "\nMappatura confermata:")

    if source_file and sheet_name:
        save_mapping(mapping, source_file=source_file, sheet_name=sheet_name, df_columns=df.columns)
        print(f"Mappatura salvata in {get_mapping_config_path()}\n")

    return mapping


def build_normalized_columns(df, mapping):
    """
    Crea colonne normalizzate in un dict separato, per comodità.
    Restituisce un dict con chiavi:
    id, email, cf, cell, tel, nome, cognome
    ciascuna è una lista di lunghezza len(df).
    """
    n = len(df)
    norm = {k: [""] * n for k in ["id", "email", "cf", "cell", "tel", "nome", "cognome"]}

    # ID: lo manteniamo così com'è (stringa)
    id_col = mapping["id"]
    norm["id"] = ["" if pd.isna(v) else str(v) for v in df[id_col].values]

    # Email
    if mapping["email"]:
        norm["email"] = [normalize_string(v) for v in df[mapping["email"]].values]

    # Codice Fiscale
    if mapping["cf"]:
        norm["cf"] = [normalize_string(v).upper() for v in df[mapping["cf"]].values]

    # Cellulare
    if mapping["cell"]:
        norm["cell"] = [normalize_phone(v) for v in df[mapping["cell"]].values]

    # Telefono fisso
    if mapping["tel"]:
        norm["tel"] = [normalize_phone(v) for v in df[mapping["tel"]].values]

    # Nome
    if mapping["nome"]:
        norm["nome"] = [normalize_string(v) for v in df[mapping["nome"]].values]

    # Cognome
    if mapping["cognome"]:
        norm["cognome"] = [normalize_string(v) for v in df[mapping["cognome"]].values]

    return norm


def collect_groups_by_key(values):
    """
    values: lista di stringhe (una per riga) per una singola chiave (es. tutte le email normalizzate)
    Ritorna dict: valore -> lista di indici (righe) in cui compare, solo se valore non vuoto
    """
    groups = {}
    for idx, val in enumerate(values):
        if not val:
            continue
        groups.setdefault(val, []).append(idx)
    # teniamo solo gruppi con almeno 2 righe
    groups = {k: v for k, v in groups.items() if len(v) > 1}
    return groups


def deduplicate(norm):
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
    name_threshold = 0.8

    # Pre-gruppi per chiavi principali
    key_names = ["email", "cf", "cell", "tel"]

    for key in key_names:
        values = norm[key]
        groups = collect_groups_by_key(values)

        for val, indices in groups.items():
            # Considera tutte le coppie di righe che condividono questa chiave
            for i in range(len(indices)):
                for j in range(i + 1, len(indices)):
                    r1 = indices[i]
                    r2 = indices[j]

                    nome1, nome2 = norm["nome"][r1], norm["nome"][r2]
                    cog1, cog2 = norm["cognome"][r1], norm["cognome"][r2]

                    names_equal = (nome1 and nome2 and nome1 == nome2 and cog1 == cog2)

                    # Similarità approssimata (se abbiamo entrambi nome e cognome)
                    sim_nome = similarity(nome1, nome2) if nome1 and nome2 else 0.0
                    sim_cog = similarity(cog1, cog2) if cog1 and cog2 else 0.0
                    names_similar = (sim_nome >= name_threshold and sim_cog >= name_threshold)

                    id1 = norm["id"][r1]
                    id2 = norm["id"][r2]

                    if names_equal:
                        # sicuramente doppia
                        sure_dups[r1].add(id2)
                        sure_dups[r2].add(id1)
                    elif names_similar:
                        # probabilmente doppia
                        prob_dups[r1].add(id2)
                        prob_dups[r2].add(id1)

    stato = []
    id_doppi = []

    for i in range(n):
        # se esistono duplicati sicuri, i "probabili" su quelle stesse righe li ignoriamo
        if sure_dups[i]:
            label = "sicuramente doppia"
            ids = sorted(sure_dups[i], key=id_sort_key)
        elif prob_dups[i]:
            label = "probabilmente doppia"
            ids = sorted(prob_dups[i], key=id_sort_key)
        else:
            label = "non doppia"
            ids = []

        stato.append(label)
        id_doppi.append(", ".join(ids))

    return stato, id_doppi


def main():
    print("=== Deduplica Excel (anagrafiche) ===\n")

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
    mapping = map_columns(df, source_file=excel_path, sheet_name=sheet_name)

    # 6. Normalizzazione valori
    norm = build_normalized_columns(df, mapping)

    # 7. Deduplica
    print("Eseguo la deduplica, potrebbe volerci qualche istante...\n")
    stato, id_doppi = deduplicate(norm)

    # 8. Aggiungo le colonne al DataFrame
    df["Stato_dedup"] = stato
    df["ID_doppi"] = id_doppi
    dedup_link_columns = build_dedup_links_columns(norm["id"], id_doppi, stato)
    for col_name, values in dedup_link_columns.items():
        df[col_name] = values

    # 9. Salvo il risultato
    base, ext = os.path.splitext(excel_path)
    out_path = f"{base}_dedup.xlsx"

    df.to_excel(out_path, sheet_name=sheet_name, index=False)
    print(f"Deduplica completata.\nFile salvato come:\n{out_path}")


if __name__ == "__main__":
    main()
