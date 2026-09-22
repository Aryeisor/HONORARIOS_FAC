import re
from datetime import date, datetime

from honorarios_app.core.common import get_col, norm, normalize_fac, to_number
from honorarios_app.specialties.cardiologia import CardiologiaRules


class FonoaudiologiaRules(CardiologiaRules):
    """Reglas de Fonoaudiologia, sin dependencia de HOSVIREPORT."""

    name = "fonoaudiologia"
    requires_hosvireport = False
    default_payment_pct = 0.70
    contratos_aplican = {"COOSALUD00125", "COOSALUD00225"}
    tarifa_codigo = "937001"
    terapia_re = re.compile(r"\bTERAPIA\s+FONOAUDIOL[ÓO]GICA\b", re.IGNORECASE)
    interconsulta_re = re.compile(r"\bINTERCONSULTA\b", re.IGNORECASE)

    def get_calc_headers(self, payment_pct=None):
        payment_pct = self.default_payment_pct if payment_pct is None else payment_pct
        return [
            "Valor COOSALUD",
            f"{int(round(payment_pct * 100))}% COOS",
            "TOTAL COOSALUD",
            "DIF SALDO - COOSALUD",
        ]

    def filter_duplicate_consult_rows(
        self, pre_rows_coo, pre_rows_other, pedir_rows_coo, pedir_rows_other,
        base_headers,
    ):
        col_map = {
            norm(header): column
            for column, header in enumerate(base_headers, start=1)
            if norm(header)
        }
        col_id = get_col(
            col_map, "IDENTIFICACION PACIENTE", "IDENTIFICACION_PACIENTE",
            "ID PACIENTE", "DOC PACIENTE", "DOCUMENTO PACIENTE",
            "IDENTIFICACION",
        )
        col_fecha = get_col(
            col_map, "FECHA DEL SERVICIO", "FECHA SERVICIO", "FECHA_PROCED",
            "FECHA PROCED", "FECHA",
        )
        col_proc = get_col(col_map, "PROCEDIMIENTO")
        col_saldo = get_col(col_map, "SALDO")
        if any(column is None for column in (col_id, col_fecha, col_proc, col_saldo)):
            return (
                list(pre_rows_coo) + list(pre_rows_other),
                list(pedir_rows_coo) + list(pedir_rows_other),
                [],
            )

        idx_id, idx_fecha = col_id - 1, col_fecha - 1
        idx_proc, idx_saldo = col_proc - 1, col_saldo - 1

        def date_key(value):
            if isinstance(value, datetime):
                return value.date().isoformat()
            if isinstance(value, date):
                return value.isoformat()
            if value is None:
                return ""
            return str(value).strip().split("T", 1)[0].split(" ", 1)[0]

        pre_rows = list(pre_rows_coo) + list(pre_rows_other)
        pedir_rows = list(pedir_rows_coo) + list(pedir_rows_other)
        if len(pre_rows) != len(pedir_rows):
            raise ValueError("Las filas PRE y PEDIR FACT no estan sincronizadas")

        groups = {}
        row_types = {}
        for row_index, row in enumerate(pre_rows):
            procedure = norm(row[idx_proc])
            if self.terapia_re.search(procedure):
                row_type = "terapia"
            elif self.interconsulta_re.search(procedure):
                row_type = "interconsulta"
            else:
                continue
            patient = normalize_fac(row[idx_id])
            service_date = date_key(row[idx_fecha])
            if not patient or not service_date:
                continue
            groups.setdefault((patient, service_date), []).append(row_index)
            row_types[row_index] = row_type

        duplicate_reasons = {}
        for row_indexes in groups.values():
            therapies = [i for i in row_indexes if row_types[i] == "terapia"]
            interconsults = [i for i in row_indexes if row_types[i] == "interconsulta"]
            if therapies:
                keep = max(therapies, key=lambda i: to_number(pre_rows[i][idx_saldo]))
                for row_index in therapies:
                    if row_index != keep:
                        duplicate_reasons[row_index] = "TERAPIA DUPLICADA MISMO PACIENTE Y FECHA; MENOR SALDO"
                for row_index in interconsults:
                    duplicate_reasons[row_index] = "INTERCONSULTA DUPLICADA POR TERAPIA MISMO PACIENTE Y FECHA"
            elif len(interconsults) > 1:
                keep = max(interconsults, key=lambda i: to_number(pre_rows[i][idx_saldo]))
                for row_index in interconsults:
                    if row_index != keep:
                        duplicate_reasons[row_index] = "INTERCONSULTA DUPLICADA MISMO PACIENTE Y FECHA; MENOR SALDO"

        final_rows = [row for i, row in enumerate(pre_rows) if i not in duplicate_reasons]
        final_pedir = [row for i, row in enumerate(pedir_rows) if i not in duplicate_reasons]
        duplicates = [
            pre_rows[i][:len(base_headers)] + [duplicate_reasons[i]]
            for i in range(len(pre_rows)) if i in duplicate_reasons
        ]
        return final_rows, final_pedir, duplicates

    def process_row(
        self, row_values, contrato_value, procedimiento_value, cod_proc_value,
        no_fac_value, saldo_value, tarifario, hosvi_by_fact, hosvi_by_alt,
        ws_original, row_number, col_map, idx_vlr_aut, hosvi_cursor_fact,
        hosvi_cursor_alt, payment_pct=None,
    ):
        payment_pct = self.default_payment_pct if payment_pct is None else payment_pct
        contract_applies = norm(contrato_value) in self.contratos_aplican
        is_therapy = normalize_fac(cod_proc_value) == self.tarifa_codigo
        base_value = payment_value = total = difference = 0.0
        new_authorization_value = None

        if contract_applies and is_therapy:
            tariff = tarifario.get(self.tarifa_codigo, {"cir": 0.0, "ayd2": 0.0})
            base_value = tariff["cir"]
            payment_value = base_value * payment_pct
            total = payment_value
            difference = to_number(saldo_value) - total
            if idx_vlr_aut is not None:
                new_authorization_value = total

        return {
            "destino": "pre",
            "calc_values": [base_value, payment_value, total, difference],
            "nuevo_vlr_aut": new_authorization_value,
            "is_coo_proced": contract_applies and is_therapy,
        }
