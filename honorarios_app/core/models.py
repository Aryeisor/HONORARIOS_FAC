from dataclasses import dataclass


@dataclass
class ProcessResult:
    output_path: str
    pre_rows: int
    pedir_rows: int
    amarillo_rows: int
    anulados_rows: int