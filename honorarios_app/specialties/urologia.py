import re

from honorarios_app.core.common import (
    get_col,
    norm,
    normalize_datetime_key,
    normalize_fac,
    to_number,
)
from honorarios_app.specialties.base import SpecialtyBase


class UrologiaRules(SpecialtyBase):
    name = "urologia"
    default_payment_pct = 0.90

    contratos_aplican = {"COOSALUD00225", "COOSALUD00125"}

    consulta_re = re.compile(r"\b(CONSULTA|INTERCONSULTA)\b", re.IGNORECASE)
    cuidados_re = re.compile(r"\bCUIDAD(O|OS)?\b", re.IGNORECASE)

    def get_calc_headers(self, payment_pct=None):
        payment_pct = payment_pct if payment_pct is not None else self.default_payment_pct
        pct_label = int(round(payment_pct * 100))

        return [
            "ROL (HOSVI)",
            "PORCENTAJE PAGO (HOSVI/OVERRIDE)",
            "100% CIRUJANO COOSALUD",
            "100% AYUD.2 COOSALUD",
            "100% COOPSA",
            "75% COOPSA",
            "60% COOPSA",
            "50% COOPSA",
            "TOTAL COOSALUD",
            f"{pct_label}% CLINICA",
            "DIF. HOSVI Y COOSALUD",
            "OBS",
        ]

    def get_extra_sheets(self):
        return {
            "DUPLICADOS": "RAZON DUPLICADO",
        }

    def detect_anulable_rows(self, ws_original, header_row, col_map):
        col_cod_proc = get_col(col_map, "CODIGO PROCEDIMIENTO", "COD PROCEDIMIENTO", "CODIGO PROC", "COD_PROCED", "COD")
        col_proc = get_col(col_map, "PROCEDIMIENTO")
        col_saldo = get_col(col_map, "SALDO")
        col_id_pac = get_col(
            col_map,
            "IDENTIFICACION PACIENTE", "IDENTIFICACION_PACIENTE",
            "ID PACIENTE", "DOC PACIENTE", "DOCUMENTO PACIENTE", "IDENTIFICACION"
        )
        col_fecha_serv = get_col(
            col_map,
            "FECHA DEL SERVICIO", "FECHA SERVICIO", "FECHA_PROCED", "FECHA PROCED", "FECHA"
        )

        anulable_rows = set()

        def _safe_date(v):
            if v is None:
                return ""
            return str(v)

        if col_saldo is not None:
            groups = {}

            for r in range(header_row + 1, ws_original.max_row + 1):
                cod_proc_val = ws_original.cell(r, col_cod_proc).value
                proc_val = ws_original.cell(r, col_proc).value if col_proc else ""
                id_pac_val = ws_original.cell(r, col_id_pac).value if col_id_pac else ""
                fecha_val = ws_original.cell(r, col_fecha_serv).value if col_fecha_serv else ""

                # IMPORTANTE:
                # Para anulados NO usar NO FAC, porque puede cambiar por refacturación.
                key = (
                    normalize_fac(id_pac_val),
                    normalize_fac(cod_proc_val),
                    norm(proc_val),
                    _safe_date(fecha_val),
                )

                groups.setdefault(key, []).append({
                    "row": r,
                    "saldo": to_number(ws_original.cell(r, col_saldo).value),
                })

            for _, items in groups.items():
                if len(items) < 2:
                    continue

                saldo_sum = sum(it["saldo"] for it in items)

                # Si el neto da 0, se consideran anulados/refacturados
                if abs(saldo_sum) < 0.01:
                    for it in items:
                        anulable_rows.add(it["row"])

        return anulable_rows

    def detect_duplicate_rows(self, ws_original, header_row, col_map):
        col_nombre_pac = (
            get_col(col_map, "NOMBRE DEL PACIENTE")
            or get_col(col_map, "NOMBRE PACIENTE")
            or get_col(col_map, "PACIENTE")
            or get_col(col_map, "NOMBRE")
        )
        col_proc = get_col(col_map, "PROCEDIMIENTO")
        col_fecha_serv = get_col(
            col_map,
            "FECHA DEL SERVICIO", "FECHA SERVICIO", "FECHA_PROCED", "FECHA PROCED", "FECHA"
        )

        if any(c is None for c in [col_nombre_pac, col_proc, col_fecha_serv]):
            return {}

        seen = {}
        duplicate_rows = {}

        def normalize_only_date(value):
            if value is None:
                return ""
            s = str(value).strip()
            return s.split(" ")[0]

        for r in range(header_row + 1, ws_original.max_row + 1):
            nombre_pac = norm(ws_original.cell(r, col_nombre_pac).value)
            procedimiento = norm(ws_original.cell(r, col_proc).value)
            fecha_serv = normalize_only_date(ws_original.cell(r, col_fecha_serv).value)

            if not nombre_pac or not procedimiento or not fecha_serv:
                continue

            # SOLO consultas / interconsultas
            if not self.consulta_re.search(procedimiento):
                continue

            key = (nombre_pac, procedimiento, fecha_serv)

            if key in seen:
                duplicate_rows[r] = "Consulta/interconsulta duplicada en la misma fecha del servicio"
            else:
                seen[key] = r

        return duplicate_rows

    def _is_urologia_procedure(self, contrato, proc_text):
        is_aplica = contrato in self.contratos_aplican
        is_consulta = bool(self.consulta_re.search(proc_text))
        is_cuidados = bool(self.cuidados_re.search(proc_text))
        return is_aplica and (not is_consulta) and (not is_cuidados), is_consulta, is_cuidados

    def _get_hosvi_match(
        self,
        no_fac_value,
        cod_proc_value,
        ws_original,
        row_number,
        col_map,
        hosvi_by_fact,
        hosvi_by_alt,
        hosvi_cursor_fact,
        hosvi_cursor_alt,
    ):
        col_id_pac = get_col(
            col_map,
            "IDENTIFICACION PACIENTE", "IDENTIFICACION_PACIENTE",
            "ID PACIENTE", "DOC PACIENTE", "DOCUMENTO PACIENTE", "IDENTIFICACION"
        )
        col_fecha_serv = get_col(
            col_map,
            "FECHA DEL SERVICIO", "FECHA SERVICIO", "FECHA_PROCED", "FECHA PROCED", "FECHA"
        )

        factura = normalize_fac(no_fac_value)
        cod = normalize_fac(cod_proc_value)
        doc_pac = normalize_fac(ws_original.cell(row_number, col_id_pac).value) if col_id_pac else ""
        fecha_proc = normalize_datetime_key(ws_original.cell(row_number, col_fecha_serv).value) if col_fecha_serv else ""

        rol = ""
        pct = None

        key_fact = (factura, cod)
        hos_list = hosvi_by_fact.get(key_fact, [])
        pos = hosvi_cursor_fact.get(key_fact, 0)

        if pos < len(hos_list):
            hos = hos_list[pos]
            hosvi_cursor_fact[key_fact] = pos + 1
            rol = hos.get("rol", "")
            pct = hos.get("pct", None)
        else:
            key_alt = (doc_pac, cod, fecha_proc)
            hos_list_alt = hosvi_by_alt.get(key_alt, [])
            pos_alt = hosvi_cursor_alt.get(key_alt, 0)

            if pos_alt < len(hos_list_alt):
                hos = hos_list_alt[pos_alt]
                hosvi_cursor_alt[key_alt] = pos_alt + 1
                rol = hos.get("rol", "")
                pct = hos.get("pct", None)

        return rol, pct

    def _get_base_value(self, cod_proc_value, rol, tarifario):
        cod = normalize_fac(cod_proc_value)
        tarifa = tarifario.get(cod, {"cir": 0.0, "ayd2": 0.0})

        if "CIR" in rol:
            return tarifa["cir"]
        if "AYUD" in rol and "2" in rol:
            return tarifa["ayd2"]

        return 0.0

    def _calculate_distribution(self, base_val, pct):
        b100 = b75 = b60 = b50 = 0.0
        obs = ""

        if base_val != 0.0 and pct is not None:
            if abs(pct - 1.75) < 1e-9:
                b100 = base_val
                b75 = base_val * 0.75
            elif abs(pct - 1.05) < 1e-9:
                b100 = base_val
                b50 = base_val * 0.50
            else:
                fact = base_val * pct

                if abs(pct - 1.0) < 1e-9:
                    b100 = fact
                elif abs(pct - 0.75) < 1e-9:
                    b75 = fact
                elif abs(pct - 0.60) < 1e-9:
                    b60 = fact
                elif abs(pct - 0.50) < 1e-9:
                    b50 = fact
                else:
                    obs = f"Porcentaje no esperado: {pct}"

        total = b100 + b75 + b60 + b50
        return b100, b75, b60, b50, total, obs

    def _calculate_clinic_value(self, total, payment_pct):
        return total * payment_pct

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
        payment_pct=None,
    ):
        payment_pct = payment_pct if payment_pct is not None else self.default_payment_pct

        contrato = norm(contrato_value)
        proc_text = norm(procedimiento_value)
        saldo_val = to_number(saldo_value)

        is_uro_proced, is_consulta, is_cuidados = self._is_urologia_procedure(contrato, proc_text)

        rol = ""
        pct = None
        v_cir = 0.0
        v_ayd2 = 0.0
        b100 = b75 = b60 = b50 = 0.0
        total = 0.0
        valor_clinica = 0.0
        diff = 0.0

        obs = "NO APLICA UROLOGIA"
        if contrato in self.contratos_aplican and (is_consulta or is_cuidados):
            obs = "UROLOGIA (consulta/cuidados) - sin cálculo"

        if is_uro_proced:
            rol, pct = self._get_hosvi_match(
                no_fac_value=no_fac_value,
                cod_proc_value=cod_proc_value,
                ws_original=ws_original,
                row_number=row_number,
                col_map=col_map,
                hosvi_by_fact=hosvi_by_fact,
                hosvi_by_alt=hosvi_by_alt,
                hosvi_cursor_fact=hosvi_cursor_fact,
                hosvi_cursor_alt=hosvi_cursor_alt,
            )

            cod = normalize_fac(cod_proc_value)
            tarifa = tarifario.get(cod, {"cir": 0.0, "ayd2": 0.0})
            v_cir = tarifa["cir"]
            v_ayd2 = tarifa["ayd2"]

            base_val = self._get_base_value(cod_proc_value, rol, tarifario)

            obs = ""
            if base_val == 0.0:
                obs = "Sin tarifa (COOSALUD) o rol no reconocido"
            if pct is None:
                obs = (obs + " | " if obs else "") + "Sin porcentaje (HOSVI / override)"

            if base_val != 0.0 and pct is not None:
                b100, b75, b60, b50, total, obs_pct = self._calculate_distribution(base_val, pct)
                if obs_pct:
                    obs = (obs + " | " if obs else "") + obs_pct

            valor_clinica = self._calculate_clinic_value(total, payment_pct)
            diff = saldo_val - valor_clinica

        calc_values = [
            rol,
            pct if pct is not None else "",
            v_cir,
            v_ayd2,
            b100,
            b75,
            b60,
            b50,
            total,
            valor_clinica,
            diff,
            obs,
        ]

        nuevo_vlr_aut = None
        if is_uro_proced and idx_vlr_aut is not None:
            nuevo_vlr_aut = valor_clinica

        return {
            "destino": "pre",
            "calc_values": calc_values,
            "nuevo_vlr_aut": nuevo_vlr_aut,
            "is_coo_proced": is_uro_proced,
        }