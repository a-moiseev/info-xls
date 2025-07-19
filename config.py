import os
from pathlib import Path

import yaml


class Config:
    """Class to handle configuration settings for the application."""

    def __init__(self):
        params = self.get_config()

        self.current_path = Path(os.getcwd())

        self.info_path = Path(
            params.get("info_path") or self.current_path / "zp_file" / "Расчет ЗП.xlsx"
        )
        self.from_files_path = Path(
            params.get("files_path") or self.current_path / "files"
        )
        self.to_files_path = Path(
            params.get("files_new_path") or self.current_path / "files_new"
        )

        self.passwd = str(params.get("password"))
        self.similarity_ratio = 0.8

    @staticmethod
    def get_config() -> dict:
        try:
            with open("config.yaml", "r") as file:
                params = yaml.safe_load(file)
        except FileNotFoundError:
            params = {}
        return params

    def update_param(self, key: str, value: str) -> None:
        """Update a configuration parameter and save to the config file."""
        params = self.get_config()
        params[key] = value

        match key:
            case "info_path":
                self.info_path = Path(value)
            case "files_path":
                self.from_files_path = Path(value)
            case "files_new_path":
                self.to_files_path = Path(value)
            case "password":
                self.passwd = str(value)
            case "similarity_ratio":
                self.similarity_ratio = float(value)

        with open("config.yaml", "w") as file:
            yaml.dump(params, file, allow_unicode=True)
