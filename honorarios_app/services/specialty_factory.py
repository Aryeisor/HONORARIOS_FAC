from honorarios_app.specialties.cardiologia import CardiologiaRules
from honorarios_app.specialties.ortopedia import OrtopediaRules
from honorarios_app.specialties.urologia import UrologiaRules
from honorarios_app.specialties.gastroenterologia import GastroenterologiaRules


def get_specialty(name: str):
    name = (name or "").strip().lower()

    specialties = {
        "ortopedia": OrtopediaRules,
        "urologia": UrologiaRules,
        "cardiologia": CardiologiaRules,
        "gastroenterologia": GastroenterologiaRules,
    }

    if name not in specialties:
        raise ValueError(f"Especialidad no soportada: {name}")

    return specialties[name]()