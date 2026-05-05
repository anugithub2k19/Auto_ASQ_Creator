from __future__ import annotations

from pathlib import Path
from typing import Literal, Optional

import yaml
from pydantic import BaseModel, Field


AuthMode = Literal["credentials", "storage_state"]


class AppConfig(BaseModel):
    url: str


class AuthConfig(BaseModel):
    mode: AuthMode = "credentials"
    username: Optional[str] = None
    password: Optional[str] = None
    storage_state_path: str = "secrets/storage_state.json"


class RunConfig(BaseModel):
    headless: bool = False
    slow_mo_ms: int = 0
    timeout_ms: int = 30_000
    sheet_name: str = "Tickets"
    grant_geolocation: bool = True
    geolocation_lat: float = 17.3850
    geolocation_lon: float = 78.4867


class OutputConfig(BaseModel):
    results_xlsx: str = "output/results.xlsx"
    artifacts_dir: str = "output/artifacts"


class Config(BaseModel):
    app: AppConfig
    auth: AuthConfig = Field(default_factory=AuthConfig)
    run: RunConfig = Field(default_factory=RunConfig)
    output: OutputConfig = Field(default_factory=OutputConfig)


def load_config(path: str | Path) -> Config:
    p = Path(path)
    data = yaml.safe_load(p.read_text(encoding="utf-8"))
    return Config.model_validate(data)

