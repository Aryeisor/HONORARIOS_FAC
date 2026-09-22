from abc import ABC, abstractmethod

from honorarios_app.core.common import norm, normalize_fac


class SpecialtyBase(ABC):
    name = "base"
    contratos_excluir = {"PARTICULAR525"}

    def apply_common_exclusion_rules(
        self,
        no_fac_value,
        contrato_value,
        cod_proc_value,
        tipo_fac_value=None,
        id_paciente_value=None,
        nombre_paciente_value=None,
    ):
        no_fac = normalize_fac(no_fac_value)
        contrato = norm(contrato_value)
        cod_proc = norm(cod_proc_value)
        tipo_fac = normalize_fac(tipo_fac_value)
        id_paciente = normalize_fac(id_paciente_value)
        nombre_paciente = norm(nombre_paciente_value)

        if tipo_fac == "17":
            return False, "TIPO FAC excluido: 17"

        if not id_paciente:
            return False, "Identificación paciente vacía"

        if not nombre_paciente:
            return False, "Nombre paciente vacío"

        if no_fac.startswith("999"):
            return False, f"NO FAC excluido (empieza por 999): {no_fac}"

        if no_fac == "7":
            return False, "NO FAC excluido: 7"

        if contrato in self.contratos_excluir:
            return False, f"Contrato excluido: {contrato}"

        if cod_proc.endswith("-F"):
            return False, f"Código con -F: {cod_proc}"

        return True, ""

    @abstractmethod
    def get_calc_headers(self):
        pass

    @abstractmethod
    def process_row(self, **kwargs):
        pass