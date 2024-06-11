import pandas as pd


def boats_filter(row):
    if row["Schema"] != "Sjösättning 2024":
        return False
    if row["Datum"][:4] != "2024":
        return False
    if pd.isna(row["Medlem (fullt namn)"]):
        return False
    return True


def get_all_boats(df: pd.DataFrame) -> list:
    return [_ for i, _ in df.iterrows() if boats_filter(_)]


def read_excel_file(file: str, filter_func) -> list:
    df = pd.read_excel(file)
    return [_ for i, _ in df.iterrows() if filter_func(_)]


# TODO: Get names from parameter
member_file = "report/Medlemmar_2023_24_20240609_1806.xlsx"
boat_file = "report/Torrsättning_2023_20240609_1806.xlsx"

members = read_excel_file(member_file, lambda x: True)
boats = read_excel_file(boat_file, boats_filter)

platser = {b["Plats"] for b in boats}
on_land = [m for m in members if m["Plats"] not in platser]

for i, m in enumerate(on_land, 1):
    print(
        f'{i:2}) {m["Medlemsnr"]:3} [{m["Plats"]:>3}] {m["Förnamn"]} {m["Efternamn"]}'
    )
