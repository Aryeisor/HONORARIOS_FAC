import re
from datetime import date, datetime

from honorarios_app.core.common import (
    get_col,
    norm,
    normalize_fac,
    to_number,
)
from honorarios_app.specialties.base import SpecialtyBase


class CardiologiaRules(SpecialtyBase):
    name = "cardiologia"

    default_consulta_pct = 0.90
    default_procedimiento_pct = 0.36

    contratos_aplican = {"COOSALUD00225", "COOSALUD00125"}

    consulta_re = re.compile(r"\b(CONSULTA|INTERCONSULTA)\b", re.IGNORECASE)
    cuidados_re = re.compile(r"\bCUIDAD(O|OS)?\b", re.IGNORECASE)
    junta_medica_re = re.compile(
        r"\bPARTICIPACION EN JUNTA MEDICA O EQUIPO INTERDISCIPLINARIO\b",
        re.IGNORECASE,
    )

    def get_calc_headers(self, payment_cfg=None):
        payment_cfg = payment_cfg or {}
        consulta_pct = payment_cfg.get("consulta_pct", self.default_consulta_pct)
        proc_pct = payment_cfg.get("procedimiento_pct", self.default_procedimiento_pct)

        consulta_label = int(round(consulta_pct * 100))
        proc_label = int(round(proc_pct * 100))

        return [
            "100% COOSALUD FORMULADO",
            f"{consulta_label}% DE COOSALUD",
            f"{proc_label}% PROCED.COOSALUD",
            "TOTAL COOSALUD",
            "DIFE. HOSV Y COOSALUD",
            "OBS",
        ]

    def get_extra_sheets(self):
        return {
            "DUPLICADOS": "RAZON DUPLICADO",
        }

    def filter_duplicate_consult_rows(
        self,
        pre_rows_coo,
        pre_rows_other,
        pedir_rows_coo,
        pedir_rows_other,
        base_headers,
    ):
        pre_col_map = {
            norm(header): column
            for column, header in enumerate(base_headers, start=1)
            if norm(header)
        }
        col_id_pac = get_col(
            pre_col_map,
            "IDENTIFICACION PACIENTE",
            "IDENTIFICACION_PACIENTE",
            "ID PACIENTE",
            "DOC PACIENTE",
            "DOCUMENTO PACIENTE",
            "IDENTIFICACION",
        )
        col_fecha = get_col(
            pre_col_map,
            "FECHA DEL SERVICIO",
            "FECHA SERVICIO",
            "FECHA_PROCED",
            "FECHA PROCED",
            "FECHA",
        )
        col_proc = get_col(pre_col_map, "PROCEDIMIENTO")
        col_saldo = get_col(pre_col_map, "SALDO")

        if any(c is None for c in [col_id_pac, col_fecha, col_proc, col_saldo]):
            return (
                list(pre_rows_coo) + list(pre_rows_other),
                list(pedir_rows_coo) + list(pedir_rows_other),
                [],
            )

        idx_id_pac = col_id_pac - 1
        idx_fecha = col_fecha - 1
        idx_proc = col_proc - 1
        idx_saldo = col_saldo - 1

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
        for row_index, row in enumerate(pre_rows):
            procedimiento = norm(row[idx_proc])
            if not (
                self.consulta_re.search(procedimiento)
                or self.cuidados_re.search(procedimiento)
            ):
                continue

            paciente = normalize_fac(row[idx_id_pac])
            fecha = date_key(row[idx_fecha])
            if not paciente or not fecha:
                continue

            groups.setdefault((paciente, fecha), []).append(row_index)

        duplicate_indexes = set()
        for row_indexes in groups.values():
            if len(row_indexes) < 2:
                continue
            row_to_keep = max(
                row_indexes,
                key=lambda row_index: to_number(pre_rows[row_index][idx_saldo]),
            )
            duplicate_indexes.update(
                row_index for row_index in row_indexes if row_index != row_to_keep
            )

        reason = "CONSULTA DUPLICADA MISMO PACIENTE Y FECHA"
        final_rows = [
            row for row_index, row in enumerate(pre_rows)
            if row_index not in duplicate_indexes
        ]
        final_pedir_rows = [
            row for row_index, row in enumerate(pedir_rows)
            if row_index not in duplicate_indexes
        ]
        duplicate_rows = [
            pre_rows[row_index][:len(base_headers)] + [reason]
            for row_index in range(len(pre_rows))
            if row_index in duplicate_indexes
        ]
        return final_rows, final_pedir_rows, duplicate_rows

    def detect_anulable_rows(self, ws_original, header_row, col_map):
        col_no_fac = get_col(col_map, "NO FAC", "NUM FAC", "N° FAC", "FACTURA")
        col_proc = get_col(col_map, "PROCEDIMIENTO")
        col_saldo = get_col(col_map, "SALDO")
        col_id_pac = get_col(
            col_map,
            "IDENTIFICACION PACIENTE", "IDENTIFICACION_PACIENTE",
            "ID PACIENTE", "DOC PACIENTE", "DOCUMENTO PACIENTE", "IDENTIFICACION"
        )

        anulable_rows = set()

        if all(c is not None for c in [col_no_fac, col_proc, col_saldo, col_id_pac]):
            groups = {}

            for r in range(header_row + 1, ws_original.max_row + 1):
                no_fac_val = ws_original.cell(r, col_no_fac).value
                proc_val = ws_original.cell(r, col_proc).value
                id_pac_val = ws_original.cell(r, col_id_pac).value
                saldo_val = ws_original.cell(r, col_saldo).value

                key = (
                    normalize_fac(no_fac_val),
                    normalize_fac(id_pac_val),
                    norm(proc_val),
                )

                if not any(key):
                    continue

                groups.setdefault(key, []).append({
                    "row": r,
                    "saldo": to_number(saldo_val),
                })

            for _, items in groups.items():
                if len(items) < 2:
                    continue

                saldo_sum = sum(it["saldo"] for it in items)

                # Anulado correcto: mismo NO FAC + paciente + procedimiento
                # y neto saldo = 0
                if abs(saldo_sum) < 0.01:
                    for it in items:
                        anulable_rows.add(it["row"])

            # Excepcion: la anulacion puede haberse registrado con otra factura.
            # Solo participan filas que no fueron anuladas por la regla principal.
            groups_without_fac = {}
            for r in range(header_row + 1, ws_original.max_row + 1):
                if r in anulable_rows:
                    continue

                proc_val = ws_original.cell(r, col_proc).value
                id_pac_val = ws_original.cell(r, col_id_pac).value
                saldo_val = ws_original.cell(r, col_saldo).value

                key = (
                    normalize_fac(id_pac_val),
                    norm(proc_val),
                )

                if not all(key):
                    continue

                groups_without_fac.setdefault(key, []).append({
                    "row": r,
                    "saldo": to_number(saldo_val),
                })

            for items in groups_without_fac.values():
                if len(items) < 2:
                    continue

                saldo_sum = sum(it["saldo"] for it in items)
                if abs(saldo_sum) < 0.01:
                    for it in items:
                        anulable_rows.add(it["row"])

        return anulable_rows

    def _is_consulta_or_cuidados(self, proc_text):
        return (
            bool(self.consulta_re.search(proc_text))
            or bool(self.cuidados_re.search(proc_text))
            or bool(self.junta_medica_re.search(proc_text))
        )

    def _is_cardiologia_valid(self, contrato):
        return contrato in self.contratos_aplican

    def _get_tarifa_base(self, cod_proc_value, tarifario):
        cod = normalize_fac(cod_proc_value)
        tarifa = tarifario.get(cod, {"cir": 0.0, "ayd2": 0.0})
        return tarifa["cir"]

    def process_row(
        self,
        row_values,
        contrato_value,
        procedimiento_value,
        cod_proc_value,
        no_fac_value,
        saldo_value,
        tarifario,
        hosvi_by_fact,
        hosvi_by_alt,
        ws_original,
        row_number,
        col_map,
        idx_vlr_aut,
        hosvi_cursor_fact,
        hosvi_cursor_alt,
        payment_cfg=None,
    ):
        payment_cfg = payment_cfg or {}
        consulta_pct = payment_cfg.get("consulta_pct", self.default_consulta_pct)
        procedimiento_pct = payment_cfg.get("procedimiento_pct", self.default_procedimiento_pct)

        contrato = norm(contrato_value)
        proc_text = norm(procedimiento_value)
        saldo_val = to_number(saldo_value)

        aplica = self._is_cardiologia_valid(contrato)
        es_consulta_cuidado = self._is_consulta_or_cuidados(proc_text)

        base_coosalud = 0.0
        valor_90 = 0.0
        valor_36 = 0.0
        total = 0.0
        diff = 0.0
        obs = "NO APLICA CARDIOLOGIA"

        if aplica:
            base_coosalud = self._get_tarifa_base(cod_proc_value, tarifario)

            if base_coosalud == 0.0:
                obs = "Sin tarifa COOSALUD"
            else:
                obs = ""
                if es_consulta_cuidado:
                    valor_90 = base_coosalud * consulta_pct
                else:
                    valor_36 = base_coosalud * procedimiento_pct

                total = valor_90 + valor_36
                diff = saldo_val - total

        calc_values = [
            base_coosalud,
            valor_90,
            valor_36,
            total,
            diff,
            obs,
        ]

        nuevo_vlr_aut = None
        if aplica and idx_vlr_aut is not None:
            nuevo_vlr_aut = total

        return {
            "destino": "pre",
            "calc_values": calc_values,
            "nuevo_vlr_aut": nuevo_vlr_aut,
            "is_coo_proced": aplica,
        }
