import json
import os
import sys
from datetime import timedelta
from typing import Tuple
from urllib.parse import quote

import pandas as pd
import requests
from pandas.core.frame import DataFrame
from pandas.core.series import Series
from tdqm import tdqm

if getattr(sys, "frozen", False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)


def get_coordinates(location: str) -> Tuple[float, float]:
    payload = json.loads(
        requests.get(
            f"https://nominatim.openstreetmap.org/search?q={quote(location)}&format=json"
        ).content
    )
    if payload == []:
        raise ValueError(f"Unable to parse the coordinates from {location}")
    data = payload[0]

    longitude = float(data.get("lon"))
    latitude = float(data.get("lat"))
    return (longitude, latitude)


def fetch_distance_duration(
    origin_longitude, origin_latitude, destination_longitude, destination_latitude
) -> Tuple[float, int]:
    raw_payload = requests.get(
        f"https://routing.openstreetmap.de/routed-car/route/v1/driving/{origin_longitude},{origin_latitude};{destination_longitude},{destination_latitude}?overview=false&geometries=polyline&steps=true"
    )
    payload = json.loads(raw_payload.content)
    if payload.get("message") == "Invalid coordinate value.":
        print(origin_longitude, origin_latitude)
        print(destination_longitude, destination_latitude)
        raise ValueError("Wrong coordinates")
    total_distance = payload.get("routes")[0]["distance"]
    total_duration = payload.get("routes")[0]["duration"]
    return round(total_distance / 1000, 2), int(total_duration)


headers = [
    "Indice",
    "Origine",
    "Origine (longitude)",
    "Origine (latitude)",
    "Destination",
    "Destination (longitude)",
    "Destination (latitude)",
    "Distance",
    "Durée",
]


def read_excel(path: str):
    return pd.read_excel(
        path,
        sheet_name="A->B",
        engine="openpyxl",
        usecols=headers,
        index_col=0,
        dtype=str,
    )


def fill_missing_coordinates(serie: Series):
    for prefix in ["Origine", "Destination"]:
        if pd.isna(serie[f"{prefix} (longitude)"]) or pd.isna(
            serie[f"{prefix} (latitude)"]
        ):
            longitude, latitude = get_coordinates(serie[prefix])
            serie[f"{prefix} (longitude)"] = longitude
            serie[f"{prefix} (latitude)"] = latitude


def fill_missing_duration_distance(df: DataFrame):
    for _, row in tdqm(df.iterrows()):
        fill_missing_coordinates(row)
        if pd.isna(row["Distance"]) or pd.isna(row["Durée"]):
            distance, duration = fetch_distance_duration(
                row["Origine (longitude)"],
                row["Origine (latitude)"],
                row["Destination (longitude)"],
                row["Destination (latitude)"],
            )
            row["Distance"] = distance
            row["Durée"] = str(timedelta(seconds=duration))
            # print("Pas de durée pour", row["Origine"], "->", row["Destination"])

        # print(row["Origine"], row["Destination"]):
    print(df)


def write_excel(path: str, df: DataFrame):
    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name="A->B")

    # Auto-adjust columns' width
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column)) + 5
        col_idx = df.columns.get_loc(column)
        writer.sheets["A->B"].set_column(col_idx, col_idx, column_width)

    writer.save()


def list_excel_files():
    print(f"{application_path}/data")
    filenames = next(os.walk(f"{application_path}/data/"), (None, None, []))[2]
    return [filename for filename in filenames if filename.endswith(".xlsx")]


def display_and_choose_excel_files():
    filenames = list_excel_files()
    while not filenames:
        print(
            "Aucun fichier Excel trouvé dans le dossier ./data/\nVerifies ce dossier puis appuies sur Entrée..."
        )
        input()
        filenames = list_excel_files()
    for i, _ in enumerate(filenames):
        print(i + 1, filenames[i])
    input_ok = False
    while not input_ok:
        try:
            print("Choisis l'indice du fichier à traiter:", end="")
            index = int(input()) - 1
            filename = filenames[index]
            input_ok = True
        except IndexError:
            print("L'indice choisi ne correspond pas aux indices proposés")
        except ValueError as value_error:
            print("La valeur tapée n'est pas un nombre")

    return filename


if __name__ == "__main__":
    filename = display_and_choose_excel_files()
    path = f"{application_path}/data/{filename}"

    df = read_excel(path)
    fill_missing_duration_distance(df)
    write_excel(path, df)
