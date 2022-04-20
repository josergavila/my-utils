import pickle
from datetime import date


__all__ = ("read_from_pickle", "save_to_pickle")

DATE = date.today().strftime("%d-%m-%y")


def read_from_pickle(filename: str):
    with open(filename, "rb") as f:
        data = pickle.load(f)
    return data


def save_to_pickle(data, filename: str) -> None:
    with open(f"{filename}_{DATE}.pickle", "wb") as f:
        pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
