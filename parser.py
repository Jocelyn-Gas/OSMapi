import json
import os
import sys
from dataclasses import dataclass
from datetime import timedelta
from typing import Tuple
from urllib.parse import quote

import pandas as pd
import requests
from pandas.core.frame import DataFrame
from pandas.core.series import Series


@dataclass(eq=True, frozen=True)
class Location:
    description: str
    longitude: float
    latitude: float


class Route:
    def __init__(self, origin: Location, destination: Location) -> None:
        self.origin = origin
        self.destination = destination
        self.distance, self.duration = self.fetch_distance_duration()

    def fetch_distance_duration(self) -> Tuple[float, int]:
        raw_payload = requests.get(
            f"https://routing.openstreetmap.de/routed-car/route/v1/driving/{self.origin.longitude},{self.origin.latitude};{self.destination.longitude},{self.destination.latitude}?overview=false&geometries=polyline&steps=true"
        )
        payload = json.loads(raw_payload.content)
        if payload.get("message") == "Invalid coordinate value.":
            print(self.origin.longitude, self.origin.latitude)
            print(self.destination.longitude, self.destination.latitude)
            raise ValueError("Wrong coordinates")
        total_distance = payload.get("routes")[0]["distance"]
        total_duration = payload.get("routes")[0]["duration"]
        return round(total_distance / 1000, 2), int(total_duration)

    def to_dict(self):
        return {
            "depart": self.origin.description,
            "depart_lon": self.origin.longitude,
            "depart_lat": self.origin.latitude,
            "arrivee": self.destination.description,
            "arrivee_lon": self.destination.longitude,
            "arrivee_lat": self.destination.latitude,
            "distance": self.distance,
            "duree": timedelta(seconds=self.duration),
        }

    def __str__(self) -> str:
        return f"{self.origin.description} -> {self.destination.description} | {self.distance}km | {str(timedelta(seconds=self.duration))}"

    def __lt__(self, other: "Route"):
        return self.duration < other.duration

    def __le__(self, other: "Route"):
        return self.duration <= other.duration

    def __gt__(self, other: "Route"):
        return self.duration > other.duration

    def __ge__(self, other: "Route"):
        return self.duration >= other.duration

    def __eq__(self, other: "Route"):
        return self.duration == other.duration

    def __ne__(self, other: "Route"):
        return self.duration != other.duration


class LocationOrderer:
    def __init__(self, locations: list[Location]) -> None:
        self.unordered_locations = locations
        self.routes = self.create_routes()

    def order_locations(
        self,
        origin: Location = None,
    ):
        if origin is None:
            origin = self.unordered_locations[0]
        locations = self.unordered_locations.copy()
        locations.remove(origin)
        ordered_locations = [origin]
        while len(locations) >= 1:
            routes = [
                route
                for route in self.routes.get(origin)
                if route.destination not in ordered_locations
            ]
            shortest: Route = min(routes)
            ordered_locations.append(shortest.destination)
            locations.remove(shortest.destination)
            origin = shortest.destination

        return ordered_locations

    def create_routes(self) -> dict[Location, list[Route]]:
        return {
            origin: [
                Route(origin, destination)
                for destination in self.unordered_locations
                if destination != origin
            ]
            for origin in self.unordered_locations
        }

    def get_routes(self, locations: list[Location]) -> list[Route]:
        routes = []
        for i, location in enumerate(locations[:-1]):
            _routes = self.routes.get(location)
            route = next(
                (route for route in _routes if route.destination == locations[i + 1])
            )
            routes.append(route)
        # print([str(route) for route in routes])
        return routes


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


headers = {
    "A->B": [
        "Indice",
        "Origine",
        "Origine (longitude)",
        "Origine (latitude)",
        "Destination",
        "Destination (longitude)",
        "Destination (latitude)",
        "Distance",
        "Durée",
    ],
    "A ordonner": ["Description", "Longitude", "Latitude"],
}


def read_excel(path: str, mode: str):
    return pd.read_excel(
        path,
        sheet_name=mode,
        engine="openpyxl",
        usecols=headers.get(mode),
        index_col=0,
        dtype=str,
    )


def order_locations(df: DataFrame):
    locations = []
    for description, row in df.iterrows():
        row["Longitude"], row["Latitude"] = get_coordinates(description)
        locations.append(Location(description, row["Longitude"], row["Latitude"]))

    orderer = LocationOrderer(locations)
    ordered_locations = orderer.order_locations()
    routes = orderer.get_routes(ordered_locations)
    return pd.DataFrame.from_dict([route.to_dict() for route in routes])


def fill_missing_coordinates(serie: Series):
    for prefix in ["Origine", "Destination"]:
        if pd.isna(serie[f"{prefix} (longitude)"]) or pd.isna(
            serie[f"{prefix} (latitude)"]
        ):
            longitude, latitude = get_coordinates(serie[prefix])
            serie[f"{prefix} (longitude)"] = longitude
            serie[f"{prefix} (latitude)"] = latitude


def fill_missing_duration_distance(df: DataFrame):
    for _, row in df.iterrows():
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


def write_excel(path: str, sheet: str, df: DataFrame, mode: int):
    if mode == 1:
        writer = pd.ExcelWriter(path)
        df.to_excel(writer, sheet_name=sheet)

        # Auto-adjust columns' width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column)) + 5
            col_idx = df.columns.get_loc(column)
            writer.sheets[sheet].set_column(col_idx, col_idx, column_width)
    elif mode == 2:
        writer = pd.ExcelWriter(path, mode="a", engine="openpyxl")
        df.to_excel(writer, sheet_name=sheet)

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
        except ValueError:
            print("La valeur tapée n'est pas un nombre")

    return filename


def choose_mode():
    choices = ["A->B", "A ordonner"]
    print("Choisis un mode de calcul")

    for i, choice in enumerate(choices):
        print(f"{choice} ({i+1})")
        input_ok = False
    while not input_ok:
        try:
            index = int(input()) - 1
            mode = choices[index]
            input_ok = True
        except IndexError:
            print("L'indice choisi ne correspond pas aux indices proposés")
        except ValueError:
            print("La valeur tapée n'est pas un nombre")
    return mode


if __name__ == "__main__":
    filename = display_and_choose_excel_files()
    mode = choose_mode()
    if mode == "A->B":
        path = f"{application_path}/data/{filename}"

        df = read_excel(path, mode)
        fill_missing_duration_distance(df)
        write_excel(path, mode, df, 1)
        print("Appuies sur n'importe quelle touche pour fermer le programme")
        input()
    else:
        path = f"{application_path}/data/{filename}"

        df = read_excel(path, mode)
        data = order_locations(df)
        write_excel(path, "Résultats", data, 2)
        print("Appuies sur n'importe quelle touche pour fermer le programme")
        input()
