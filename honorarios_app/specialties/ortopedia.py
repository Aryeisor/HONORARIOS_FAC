from honorarios_app.core.common import (
    get_col,
    norm,
    normalize_datetime_key,
    normalize_fac,
    to_number,
)
from honorarios_app.specialties.base import SpecialtyBase


class OrtopediaRules(SpecialtyBase):
    contratos_aplican = {"COOSALUD00225", "COOSALUD00125"}
    default_payment_pct = 0.70

    def get_calc_headers(self, payment_pct=None):
        p = payment_pct if payment_pct is not None else self.default_payment_pct
        pct_txt = f"{int(round(p * 100))}% CLINICA"

        return [
            "ROL HOSVI",
            "% HOSVI",
            "VALOR CIRUJANO COOSALUD",
            "VALOR AYUDANTE 2 COOSALUD",
            "100% COOPSA",
            "75% COOPSA",
            "70% COOPSA",
            "60% COOPSA",
            "50% COOPSA",
            "TOTAL COOSALUD",
            pct_txt,
            "DIF SALDO-COOSALUD",
            "OBS",
        ]

    def detect_anulable_rows(self, ws_original, header_row, col_map):
        col_no_fac = get_col(col_map, "NO FAC", "NUM FAC", "N° FAC", "FACTURA")
        col_cod_proc = get_col(
            col_map,
            "CODIGO PROCEDIMIENTO",
            "COD PROCEDIMIENTO",
            "CODIGO PROC",
            "COD_PROCED",
            "COD",
        )
        col_proc = get_col(col_map, "PROCEDIMIENTO")
        col_saldo = get_col(col_map, "SALDO")
        col_estado = get_col(col_map, "ESTADO")
        col_id_pac = get_col(
            col_map,
            "IDENTIFICACION PACIENTE",
            "IDENTIFICACION_PACIENTE",
            "ID PACIENTE",
            "DOC PACIENTE",
            "DOCUMENTO PACIENTE",
            "IDENTIFICACION",
        )
        col_fecha_serv = get_col(
            col_map,
            "FECHA DEL SERVICIO",
            "FECHA SERVICIO",
            "FECHA_PROCED",
            "FECHA PROCED",
            "FECHA",
        )

        anulable_rows = set()

        if None in [col_no_fac, col_cod_proc, col_proc, col_saldo, col_estado]:
            return anulable_rows

        groups = {}

        for r in range(header_row + 1, ws_original.max_row + 1):
            no_fac_val = ws_original.cell(r, col_no_fac).value
            cod_proc_val = ws_original.cell(r, col_cod_proc).value
            proc_val = ws_original.cell(r, col_proc).value
            saldo_val = ws_original.cell(r, col_saldo).value
            estado_val = ws_original.cell(r, col_estado).value
            id_pac_val = ws_original.cell(r, col_id_pac).value if col_id_pac else ""
            fecha_val = ws_original.cell(r, col_fecha_serv).value if col_fecha_serv else ""

            key = (
                normalize_fac(no_fac_val),
                normalize_fac(id_pac_val),
                normalize_fac(cod_proc_val),
                norm(proc_val),
                normalize_datetime_key(fecha_val),
            )

            groups.setdefault(key, []).append(
                {
                    "row": r,
                    "saldo": to_number(saldo_val),
                    "estado": norm(estado_val),
                }
            )

        for _, items in groups.items():
            if len(items) < 2:
                continue

            saldo_sum = sum(it["saldo"] for it in items)
            estados = {it["estado"] for it in items if it["estado"]}

            has_anulado = any("ANUL" in e for e in estados)
            has_base = any("ANUL" not in e for e in estados) if estados else True

            if has_anulado and has_base and abs(saldo_sum) < 0.01:
                for it in items:
                    anulable_rows.add(it["row"])

        return anulable_rows

    def _is_ortopedia_procedure(self, contrato, proc_text):
        is_aplica = contrato in self.contratos_aplican
        proc_upper = norm(proc_text)

        is_consulta = "CONSULTA" in proc_upper or "INTERCONSULTA" in proc_upper
        is_cuidados = "CUIDADOS" in proc_upper

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
            "IDENTIFICACION PACIENTE",
            "IDENTIFICACION_PACIENTE",
            "ID PACIENTE",
            "DOC PACIENTE",
            "DOCUMENTO PACIENTE",
            "IDENTIFICACION",
        )
        col_fecha_serv = get_col(
            col_map,
            "FECHA DEL SERVICIO",
            "FECHA SERVICIO",
            "FECHA_PROCED",
            "FECHA PROCED",
            "FECHA",
        )

        factura = normalize_fac(no_fac_value)
        cod = normalize_fac(cod_proc_value)
        doc_pac = normalize_fac(ws_original.cell(row_number, col_id_pac).value) if col_id_pac else ""
        fecha_proc = (
            normalize_datetime_key(ws_original.cell(row_number, col_fecha_serv).value)
            if col_fecha_serv
            else ""
        )

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

        if "CIR" in norm(rol):
            return tarifa["cir"]
        if "AYUD" in norm(rol) and "2" in norm(rol):
            return tarifa["ayd2"]

        return 0.0

    def _calculate_distribution(self, base_val, pct):
        b100 = b75 = b70 = b60 = b50 = 0.0
        obs = ""

        if base_val != 0.0 and pct is not None:
            fact = base_val * pct

            if abs(pct - 1.00) < 1e-9:
                b100 = fact
            elif abs(pct - 0.75) < 1e-9:
                b75 = fact
            elif abs(pct - 0.70) < 1e-9:
                b70 = fact
            elif abs(pct - 0.60) < 1e-9:
                b60 = fact
            elif abs(pct - 0.50) < 1e-9:
                b50 = fact
            else:
                obs = f"Porcentaje no esperado: {pct}"

        total = b100 + b75 + b70 + b60 + b50
        return b100, b75, b70, b60, b50, total, obs

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

        is_orto_proced, is_consulta, is_cuidados = self._is_ortopedia_procedure(contrato, proc_text)

        rol = ""
        pct = None
        v_cir = 0.0
        v_ayd2 = 0.0
        b100 = b75 = b70 = b60 = b50 = 0.0
        total = 0.0
        valor_clinica = 0.0
        diff = 0.0

        obs = "NO APLICA ORTOPEDIA"
        if contrato in self.contratos_aplican and (is_consulta or is_cuidados):
            obs = "ORTOPEDIA (consulta/cuidados) - sin cálculo"

        if is_orto_proced:
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
                b100, b75, b70, b60, b50, total, obs_pct = self._calculate_distribution(base_val, pct)
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
            b70,
            b60,
            b50,
            total,
            valor_clinica,
            diff,
            obs,
        ]

        nuevo_vlr_aut = None
        if is_orto_proced and idx_vlr_aut is not None:
            nuevo_vlr_aut = valor_clinica

        return {
            "destino": "pre",
            "calc_values": calc_values,
            "nuevo_vlr_aut": nuevo_vlr_aut,
            "is_coo_proced": is_orto_proced,
        }