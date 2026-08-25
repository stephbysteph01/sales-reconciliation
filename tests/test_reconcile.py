from pathlib import Path

import pandas as pd

from sales_reconciliation.reconcile import export_reconciliation_report, normalize_customer_name, normalize_id, reconcile_sales_and_payments


def test_normalize_id_and_customer_name():
    assert normalize_id(" rom-123 ") == "ROM123"
    assert normalize_customer_name("  Rossi   Mario  ") == "rossi mario"


def test_reconcile_flags_known_anomalies():
    sales = pd.DataFrame(
        [
            {"ID_Vendita": "A1", "Data": "2024-01-05", "Cliente": "Rossi Mario", "Importo": 100.0, "Metodo_Pagamento": "Carta"},
            {"ID_Vendita": "A1", "Data": "2024-01-05", "Cliente": "rossi mario", "Importo": 100.0, "Metodo_Pagamento": "Carta"},
            {"ID_Vendita": "B2", "Data": "2024-01-06", "Cliente": "Bianchi Anna", "Importo": 220.0, "Metodo_Pagamento": "Bonifico"},
            {"ID_Vendita": "C3", "Data": "2024-01-07", "Cliente": "Verdi Luca", "Importo": 300.0, "Metodo_Pagamento": "PayPal"},
        ]
    )
    payments = pd.DataFrame(
        [
            {"ID_Vendita": "A1", "Data_Incasso": "2024-01-05", "Importo_Incassato": 100.0},
            {"ID_Vendita": "A1", "Data_Incasso": "2024-01-05", "Importo_Incassato": 99.0},
            {"ID_Vendita": "C3", "Data_Incasso": "2024-01-07", "Importo_Incassato": 280.0},
        ]
    )

    report = reconcile_sales_and_payments(sales, payments)
    assert report["vendite_senza_incasso"] == 1
    assert report["incassi_non_coerenti"] == 1
    assert report["id_vendita_duplicati"] == 1
    assert report["vendite_duplicati"] == ["A1"]

    output = Path("tmp_reconciliation_report.xlsx")
    csv_output = Path("tmp_reconciliation_details.csv")
    export_reconciliation_report(report, output, sales, payments, csv_path=csv_output)
    with pd.ExcelFile(output) as xls:
        expected = {
            "Riepilogo",
            "Vendite senza incasso",
            "Incassi non coerenti",
            "Vendite duplicate",
            "Incassi duplicati",
            "Dettaglio - Vendite senza",
            "Dettaglio - Incassi non",
            "Dettaglio - Vendite dupl",
            "Dettaglio - Incassi dup",
        }
        assert expected.issubset(set(xls.sheet_names))
    assert csv_output.exists()
    output.unlink(missing_ok=True)
    csv_output.unlink(missing_ok=True)
