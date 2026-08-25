from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd


def normalize_id(value: object) -> str:
    if pd.isna(value):
        return ""
    cleaned = re.sub(r"[^A-Za-z0-9]", "", str(value).strip().upper())
    return cleaned


def normalize_customer_name(value: object) -> str:
    if pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def standardize_sales_frame(frame: pd.DataFrame, filiale: str | None = None) -> pd.DataFrame:
    cleaned = frame.copy()
    cleaned.columns = [str(col).strip().lower().replace(" ", "_") for col in cleaned.columns]

    if "id_vendita" not in cleaned.columns:
        raise ValueError("Il dataframe vendite non contiene la colonna ID_Vendita.")

    cleaned["ID_Vendita"] = cleaned["id_vendita"].map(normalize_id)
    cleaned["Data"] = pd.to_datetime(cleaned["data"], errors="coerce", dayfirst=True, format="mixed")
    cleaned["Cliente"] = cleaned["cliente"].map(normalize_customer_name)
    cleaned["Importo"] = pd.to_numeric(cleaned["importo"], errors="coerce")
    cleaned["Metodo_Pagamento"] = cleaned["metodo_pagamento"].fillna("N.D.").astype(str).str.strip().str.title()
    cleaned["Filiale"] = filiale or "N.D."
    return cleaned[["ID_Vendita", "Data", "Cliente", "Importo", "Metodo_Pagamento", "Filiale"]]


def standardize_payments_frame(frame: pd.DataFrame) -> pd.DataFrame:
    cleaned = frame.copy()
    cleaned.columns = [str(col).strip().lower().replace(" ", "_") for col in cleaned.columns]

    if "id_vendita" not in cleaned.columns:
        raise ValueError("Il dataframe incassi non contiene la colonna ID_Vendita.")

    cleaned["ID_Vendita"] = cleaned["id_vendita"].map(normalize_id)
    cleaned["Data_Incasso"] = pd.to_datetime(cleaned["data_incasso"], errors="coerce", dayfirst=True, format="mixed")
    cleaned["Importo_Incassato"] = pd.to_numeric(cleaned["importo_incassato"], errors="coerce")
    return cleaned[["ID_Vendita", "Data_Incasso", "Importo_Incassato"]]


def read_sales_files(directory: str | Path) -> pd.DataFrame:
    directory = Path(directory)
    frames = []
    for sales_file in sorted(directory.glob("vendite_*.xlsx")):
        frame = pd.read_excel(sales_file)
        filiale = sales_file.stem.replace("vendite_", "")
        frames.append(standardize_sales_frame(frame, filiale))

    if not frames:
        raise FileNotFoundError(f"Nessun file vendite_*.xlsx trovato in {directory}")

    return pd.concat(frames, ignore_index=True)


def read_payments_file(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File incassi non trovato: {path}")
    frame = pd.read_excel(path)
    return standardize_payments_frame(frame)


def reconcile_sales_and_payments(sales_df: pd.DataFrame, payments_df: pd.DataFrame) -> dict[str, int | list[str]]:
    if "ID_Vendita" not in sales_df.columns or "ID_Vendita" not in payments_df.columns:
        raise ValueError("I dataframe devono contenere la colonna ID_Vendita.")

    sales = sales_df.copy()
    payments = payments_df.copy()

    duplicate_sales = sales[sales["ID_Vendita"].duplicated(keep=False)]
    duplicate_payments = payments[payments["ID_Vendita"].duplicated(keep=False)]

    sales_unique = sales[["ID_Vendita", "Importo"]].drop_duplicates(subset="ID_Vendita", keep="first")
    payments_unique = payments[["ID_Vendita", "Importo_Incassato"]].drop_duplicates(subset="ID_Vendita", keep="last")
    sales_with_payment = sales_unique.merge(payments_unique, on="ID_Vendita", how="left")

    vendite_non_incassate = sales_with_payment[sales_with_payment["Importo_Incassato"].isna()].copy()
    incassi_non_coerenti = sales_with_payment[
        ~sales_with_payment["Importo_Incassato"].isna()
        & (sales_with_payment["Importo"] != 0)
        & (
            (abs(sales_with_payment["Importo_Incassato"] - sales_with_payment["Importo"]) / sales_with_payment["Importo"]) > 0.05
        )
    ].copy()

    report = {
        "vendite_senza_incasso": int(vendite_non_incassate["ID_Vendita"].nunique()),
        "incassi_non_coerenti": int(incassi_non_coerenti["ID_Vendita"].nunique()),
        "id_vendita_duplicati": int(duplicate_sales["ID_Vendita"].nunique()),
        "id_incasso_duplicati": int(duplicate_payments["ID_Vendita"].nunique()),
        "id_vendita_non_incassate": sorted(vendite_non_incassate["ID_Vendita"].drop_duplicates().tolist()),
        "id_vendita_con_differenza_importo": sorted(incassi_non_coerenti["ID_Vendita"].drop_duplicates().tolist()),
        "vendite_duplicati": sorted(duplicate_sales["ID_Vendita"].drop_duplicates().tolist()),
        "incassi_duplicati": sorted(duplicate_payments["ID_Vendita"].drop_duplicates().tolist()),
    }
    return report


def _friendly_sheet_name(name: str) -> str:
    name = name.replace("_", " ")
    if name.startswith("Dettaglio - "):
        base = name.replace("Dettaglio - ", "")
        mapping = {
            "Vendite senza incasso": "Dettaglio - Vendite senza",
            "Incassi non coerenti": "Dettaglio - Incassi non",
            "Vendite duplicate": "Dettaglio - Vendite dupl",
            "Incassi duplicati": "Dettaglio - Incassi dup",
        }
        return mapping.get(base, base[:31])
    return name[:31]


def _build_detail_frames(sales_df: pd.DataFrame, payments_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    sales = sales_df.copy()
    payments = payments_df.copy()

    if "Filiale" not in sales.columns:
        sales["Filiale"] = "N.D."

    sales_unique = sales[["ID_Vendita", "Filiale", "Cliente", "Data", "Importo", "Metodo_Pagamento"]].drop_duplicates(subset="ID_Vendita", keep="first")
    payments_unique = payments[["ID_Vendita", "Data_Incasso", "Importo_Incassato"]].drop_duplicates(subset="ID_Vendita", keep="last")
    merged = sales_unique.merge(payments_unique, on="ID_Vendita", how="left")

    missing = merged[merged["Importo_Incassato"].isna()].copy()
    missing["Data_Incasso"] = pd.NaT
    missing["Importo_Incassato"] = pd.NA
    missing["Differenza_Importo"] = pd.NA
    missing["Note"] = "Vendita senza incasso corrispondente"
    missing = missing[[
        "ID_Vendita",
        "Filiale",
        "Cliente",
        "Data",
        "Importo",
        "Metodo_Pagamento",
        "Data_Incasso",
        "Importo_Incassato",
        "Differenza_Importo",
        "Note",
    ]].rename(columns={"Data": "Data_Vendita"})

    mismatched = merged[
        ~merged["Importo_Incassato"].isna()
        & (merged["Importo"] != 0)
        & ((abs(merged["Importo_Incassato"] - merged["Importo"]) / merged["Importo"]) > 0.05)
    ].copy()
    mismatched["Differenza_Importo"] = mismatched["Importo_Incassato"] - mismatched["Importo"]
    mismatched["Note"] = "Importo incassato diverso dalla vendita"
    mismatched = mismatched[[
        "ID_Vendita",
        "Filiale",
        "Cliente",
        "Data",
        "Importo",
        "Metodo_Pagamento",
        "Data_Incasso",
        "Importo_Incassato",
        "Differenza_Importo",
        "Note",
    ]].rename(columns={"Data": "Data_Vendita"})

    duplicate_sales = sales[sales["ID_Vendita"].duplicated(keep=False)].copy()
    duplicate_sales["Note"] = "ID duplicato in vendite"
    duplicate_sales = duplicate_sales[[
        "ID_Vendita",
        "Filiale",
        "Cliente",
        "Data",
        "Importo",
        "Metodo_Pagamento",
        "Note",
    ]].rename(columns={"Data": "Data_Vendita"})
    duplicate_sales["Data_Incasso"] = pd.NaT
    duplicate_sales["Importo_Incassato"] = pd.NA
    duplicate_sales["Differenza_Importo"] = pd.NA
    duplicate_sales = duplicate_sales[[
        "ID_Vendita",
        "Filiale",
        "Cliente",
        "Data_Vendita",
        "Importo",
        "Metodo_Pagamento",
        "Data_Incasso",
        "Importo_Incassato",
        "Differenza_Importo",
        "Note",
    ]]

    duplicate_payments = payments[payments["ID_Vendita"].duplicated(keep=False)].copy()
    duplicate_payments["Note"] = "ID duplicato in incassi"
    duplicate_payments = duplicate_payments[[
        "ID_Vendita",
        "Data_Incasso",
        "Importo_Incassato",
        "Note",
    ]].copy()
    duplicate_payments["Filiale"] = pd.NA
    duplicate_payments["Cliente"] = pd.NA
    duplicate_payments["Data_Vendita"] = pd.NaT
    duplicate_payments["Importo"] = pd.NA
    duplicate_payments["Metodo_Pagamento"] = pd.NA
    duplicate_payments["Differenza_Importo"] = pd.NA
    duplicate_payments = duplicate_payments[[
        "ID_Vendita",
        "Filiale",
        "Cliente",
        "Data_Vendita",
        "Importo",
        "Metodo_Pagamento",
        "Data_Incasso",
        "Importo_Incassato",
        "Differenza_Importo",
        "Note",
    ]]

    return {
        "Vendite senza incasso": missing,
        "Incassi non coerenti": mismatched,
        "Vendite duplicate": duplicate_sales,
        "Incassi duplicati": duplicate_payments,
    }


def export_reconciliation_report(
    report: dict[str, int | list[str]],
    output_path: str | Path,
    sales_df: pd.DataFrame | None = None,
    payments_df: pd.DataFrame | None = None,
    csv_path: str | Path | None = None,
) -> tuple[Path, Path | None]:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    temp_output = output.with_suffix(f"{output.suffix}.tmp")
    if temp_output.exists():
        temp_output.unlink(missing_ok=True)

    detail_csv_path = Path(csv_path) if csv_path is not None else output.with_name(f"{output.stem}_dettagli.csv")

    try:
        with pd.ExcelWriter(temp_output) as writer:
            summary_rows = [
                ("Vendite senza incasso", report.get("vendite_senza_incasso", 0)),
                ("Incassi non coerenti", report.get("incassi_non_coerenti", 0)),
                ("ID vendite duplicate", report.get("id_vendita_duplicati", 0)),
                ("ID incassi duplicati", report.get("id_incasso_duplicati", 0)),
            ]
            pd.DataFrame(summary_rows, columns=["Categoria", "Valore"]).to_excel(writer, sheet_name="Riepilogo", index=False)

            sheet_map = {
                "Vendite senza incasso": report.get("id_vendita_non_incassate", []),
                "Incassi non coerenti": report.get("id_vendita_con_differenza_importo", []),
                "Vendite duplicate": report.get("vendite_duplicati", []),
                "Incassi duplicati": report.get("incassi_duplicati", []),
            }

            for sheet_name, ids in sheet_map.items():
                pd.DataFrame({"ID_Vendita": ids}).to_excel(writer, sheet_name=_friendly_sheet_name(sheet_name), index=False)

            if sales_df is not None and payments_df is not None:
                detail_frames = _build_detail_frames(sales_df, payments_df)
                combined_details = []
                for sheet_name, frame in detail_frames.items():
                    frame = frame.copy()
                    frame["Categoria"] = sheet_name
                    combined_details.append(frame)
                    frame.to_excel(writer, sheet_name=_friendly_sheet_name(f"Dettaglio - {sheet_name}"), index=False)

                if combined_details:
                    all_details = pd.concat(combined_details, ignore_index=True)
                    all_details.to_csv(detail_csv_path, index=False)

        try:
            if output.exists():
                output.unlink()
        except PermissionError:
            pass

        try:
            temp_output.replace(output)
        except PermissionError:
            pass
    except PermissionError:
        if sales_df is not None and payments_df is not None:
            detail_frames = _build_detail_frames(sales_df, payments_df)
            combined_details = pd.concat([frame.assign(Categoria=name) for name, frame in detail_frames.items()], ignore_index=True)
            combined_details.to_csv(detail_csv_path, index=False)
        return output, detail_csv_path

    return output, detail_csv_path
