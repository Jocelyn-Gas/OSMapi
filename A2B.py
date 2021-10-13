import json
import os
import sys
import time
from datetime import timedelta
from typing import Tuple
from urllib.parse import quote

import easygui
import pandas as pd
import requests
from pandas.core.frame import DataFrame
from pandas.core.series import Series
from tqdm import tqdm

from src.utils import read_excel

if getattr(sys, "frozen", False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)


def main(filename: str):
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
    print("Ouverture du fichier...", end="\r")
    df = read_excel(filename, "A->B", headers)
    print("Ouverture terminée     ", end="\r")

    for _, row in tqdm(df.iterrows(), total=len(df.index)):
        fill_missing_coordinates(row)
        fill_missing_distance_duration(row)

    write_excel(filename, df)
    open_excel(filename)


def write_excel(path: str, df: DataFrame, sheet: str = "A->B"):
    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name=sheet)
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column)) + 5
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheet].set_column(col_idx, col_idx, column_width)

    writer.save()


def open_excel(path):
    os.startfile(path)


def fill_missing_distance_duration(row: Series):
    if pd.isna(row["Distance"]) or pd.isna(row["Durée"]):
        distance, duration = fetch_distance_duration(
            row["Origine (longitude)"],
            row["Origine (latitude)"],
            row["Destination (longitude)"],
            row["Destination (latitude)"],
        )
        row["Distance"] = distance
        row["Durée"] = str(timedelta(seconds=duration))


def fill_missing_coordinates(serie: Series):
    for prefix in ["Origine", "Destination"]:
        if pd.isna(serie[f"{prefix} (longitude)"]) or pd.isna(
            serie[f"{prefix} (latitude)"]
        ):
            longitude, latitude = get_coordinates(serie[prefix])
            serie[f"{prefix} (longitude)"] = longitude
            serie[f"{prefix} (latitude)"] = latitude


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


if __name__ == "__main__":
    filename = easygui.fileopenbox(
        msg="Veuillez sélectionner le fichier excel d'entrée.",
        title="Sélection de fichier",
        default=application_path,
        filetypes="*.xlsx",
    )
    main(filename)
