import json
import os

DEFAULT_SETTINGS = {
    "ortopedia": {
        "payment_pct": 0.70
    },
    "urologia": {
        "payment_pct": 0.90
    },
    "cardiologia": {
        "consulta_pct": 0.90,
        "procedimiento_pct": 0.36
    },
    "gastroenterologia": {
        "payment_pct": 0.70
    },
}


def get_app_config_dir():
    """
    Devuelve una carpeta persistente de configuración para la aplicación.
    En Windows usará:
    C:\\Users\\USUARIO\\AppData\\Local\\HONORARIOS
    """
    local_appdata = os.environ.get("LOCALAPPDATA")

    if local_appdata:
        return os.path.join(local_appdata, "HONORARIOS")

    return os.path.join(os.path.expanduser("~"), ".honorarios")


CONFIG_DIR = get_app_config_dir()
SETTINGS_FILE = os.path.join(CONFIG_DIR, "settings.json")


def ensure_settings_file():
    os.makedirs(CONFIG_DIR, exist_ok=True)
    if not os.path.exists(SETTINGS_FILE):
        save_settings(DEFAULT_SETTINGS)


def load_settings():
    ensure_settings_file()

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = DEFAULT_SETTINGS.copy()

    for key, value in DEFAULT_SETTINGS.items():
        if key not in data:
            data[key] = value.copy()
        else:
            for sub_key, sub_value in value.items():
                if sub_key not in data[key]:
                    data[key][sub_key] = sub_value

    return data


def save_settings(data):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def get_specialty_settings(specialty_name: str) -> dict:
    data = load_settings()
    return data.get(specialty_name, DEFAULT_SETTINGS.get(specialty_name, {})).copy()


def set_specialty_settings(specialty_name: str, settings_dict: dict):
    data = load_settings()
    current = data.get(specialty_name, {}).copy()
    current.update(settings_dict)
    data[specialty_name] = current
    save_settings(data)


def get_specialty_base_pct(specialty_name: str) -> float:
    settings = get_specialty_settings(specialty_name)

    if specialty_name == "cardiologia":
        return settings.get("consulta_pct", 0.90)

    return settings.get("payment_pct", 0.70)


def set_specialty_base_pct(specialty_name: str, pct: float):
    if specialty_name == "cardiologia":
        current = get_specialty_settings(specialty_name)
        set_specialty_settings(
            specialty_name,
            {
                "consulta_pct": pct,
                "procedimiento_pct": current.get("procedimiento_pct", 0.36)
            }
        )
    else:
        set_specialty_settings(
            specialty_name,
            {"payment_pct": pct}
        )