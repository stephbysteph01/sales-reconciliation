from __future__ import annotations

import random
from datetime import date
from pathlib import Path

import pandas as pd
from faker import Faker

BRANCHES = {
    "roma": {"rows": 82, "prefix": "ROM"},
    "milano": {"rows": 77, "prefix": "MIL"},
    "firenze": {"rows": 86, "prefix": "FIR"},
}


def _random_date(fake: Faker, start: str = "2024-01-01", end: str = "2025-01-31") -> str:
    start_date = date.fromisoformat(start)
    end_date = date.fromisoformat(end)
    value = fake.date_between(start_date=start_date, end_date=end_date)
    variants = [
        value.isoformat(),
        value.strftime("%d/%m/%Y"),
        value.strftime("%Y/%m/%d"),
        value.strftime("%d-%m-%Y"),
        "",
    ]
    return random.choice(variants)


def _make_customer_name(fake: Faker) -> str:
    name = fake.name()
    variants = [
        name,
        name.lower(),
        name.upper(),
        "  ".join(name.split()),
        name.replace(" ", ""),
        name.replace(" ", " "),
    ]
    return random.choice(variants)


def _sale_id_for(branch_prefix: str, index: int) -> str:
    return f"{branch_prefix}-{index:04d}"


def generate_sales_files(output_dir: Path) -> list[Path]:
    fake = Faker("it_IT")
    file_paths: list[Path] = []

    duplicate_ids = {
        "roma": ["ROM-0025", "ROM-0110"],
        "milano": ["MIL-0043", "MIL-0098"],
        "firenze": ["FIR-0064", "FIR-0140"],
    }

    for branch_name, config in BRANCHES.items():
        rows: list[dict[str, object]] = []
        branch_prefix = config["prefix"]
        for index in range(1, config["rows"] + 1):
            sale_id = _sale_id_for(branch_prefix, index)
            customer = _make_customer_name(fake)
            data = _random_date(fake)
            importo = round(random.uniform(65.00, 3200.00), 2)
            metodo = random.choice(["Contanti", "Carta", "Bonifico", "PayPal", "Assegno"])
            rows.append(
                {
                    "ID_Vendita": sale_id,
                    "Data": data,
                    "Cliente": customer,
                    "Importo": importo,
                    "Metodo_Pagamento": metodo,
                }
            )

        duplicate_target = {
            "roma": "ROM-0048",
            "milano": "MIL-0061",
            "firenze": "FIR-0073",
        }[branch_name]
        duplicate_source = next(item for item in rows if item["ID_Vendita"] == duplicate_target)
        rows.append({
            "ID_Vendita": duplicate_target,
            "Data": _random_date(fake),
            "Cliente": duplicate_source["Cliente"],
            "Importo": round(float(duplicate_source["Importo"]) + random.uniform(-20.0, 20.0), 2),
            "Metodo_Pagamento": duplicate_source["Metodo_Pagamento"],
        })

        df = pd.DataFrame(rows)
        file_path = output_dir / f"vendite_{branch_name}.xlsx"
        df.to_excel(file_path, index=False)
        file_paths.append(file_path)

    return file_paths


def generate_cash_file(sales_dir: Path, output_path: Path) -> Path:
    fake = Faker("it_IT")
    all_sales: list[pd.DataFrame] = []
    for sales_file in sorted(sales_dir.glob("vendite_*.xlsx")):
        all_sales.append(pd.read_excel(sales_file))

    sales_df = pd.concat(all_sales, ignore_index=True)
    sale_ids = sales_df["ID_Vendita"].tolist()
    missing_ids = set(sale_ids[::7])
    duplicate_payment_ids = [sale_ids[12], sale_ids[40], sale_ids[88]]

    payment_rows: list[dict[str, object]] = []
    seen_ids: set[str] = set()

    for _, row in sales_df.iterrows():
        sale_id = str(row["ID_Vendita"]).strip()
        if sale_id in missing_ids:
            continue

        base_amount = float(row["Importo"])
        amount = round(base_amount + random.uniform(-25.0, 25.0), 2)
        if random.random() < 0.35:
            amount = round(base_amount * random.uniform(0.97, 1.02), 2)

        date_value = row["Data"]
        if pd.isna(date_value):
            date_value = ""
        elif isinstance(date_value, str):
            date_value = date_value.strip() or ""

        if not date_value:
            date_value = _random_date(fake)

        if random.random() < 0.2:
            date_value = date_value.replace("-", "/")

        payment_rows.append(
            {
                "ID_Vendita": sale_id,
                "Data_Incasso": date_value,
                "Importo_Incassato": round(amount, 2),
            }
        )
        seen_ids.add(sale_id)

    for sale_id in duplicate_payment_ids:
        amount = round(float(random.uniform(100.0, 1500.0)), 2)
        payment_rows.append(
            {
                "ID_Vendita": sale_id,
                "Data_Incasso": fake.date_between(
                    start_date=date(2024, 1, 1),
                    end_date=date(2025, 1, 31),
                ).strftime("%d/%m/%Y"),
                "Importo_Incassato": amount,
            }
        )

    payment_df = pd.DataFrame(payment_rows)
    payment_df.to_excel(output_path, index=False)
    return output_path


def generate_demo_data(project_dir: str | Path = ".") -> dict[str, Path]:
    project_path = Path(project_dir)
    if project_path.name == "raw":
        raw_dir = project_path
    else:
        raw_dir = project_path / "data" / "raw"
    raw_dir.mkdir(parents=True, exist_ok=True)

    sales_files = generate_sales_files(raw_dir)
    cash_file = generate_cash_file(raw_dir, raw_dir / "incassi.xlsx")
    print(f"Creati {len(sales_files)} file vendite e 1 file incassi in {raw_dir}")
    return {"sales_files": sales_files, "cash_file": cash_file}


if __name__ == "__main__":
    generate_demo_data(Path(__file__).resolve().parents[2])
