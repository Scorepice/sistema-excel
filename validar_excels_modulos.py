import argparse
import os
import shutil
import sqlite3
import tempfile
from dataclasses import dataclass
from typing import Dict, List, Optional

from app import (
    DB_NAME,
    MODULES,
    align_dataframe_to_existing_table,
    leer_excel_completo,
    normalizar_texto,
    should_parse_as_date,
    parse_date_like_value,
    table_exists,
    table_name_for_module,
)


@dataclass
class ModuleValidationResult:
    module_key: str
    module_name: str
    excel_path: Optional[str]
    status: str
    details: List[str]


def find_excel_for_module(module_name: str, excel_files: List[str]) -> Optional[str]:
    target = normalizar_texto(module_name)
    matches = []

    for path in excel_files:
        file_name = os.path.splitext(os.path.basename(path))[0]
        if target in normalizar_texto(file_name):
            matches.append(path)

    if not matches:
        return None

    # Prefiere coincidencia exacta por nombre entre parentesis, luego por ruta corta.
    exact_priority = []
    for path in matches:
        file_name = os.path.splitext(os.path.basename(path))[0]
        nfile = normalizar_texto(file_name)
        if nfile.endswith(target) or f"({target})" in nfile:
            exact_priority.append(path)

    if exact_priority:
        return sorted(exact_priority, key=len)[0]

    return sorted(matches, key=len)[0]


def get_excel_files(folder: str) -> List[str]:
    files = []
    for root, _, names in os.walk(folder):
        for name in names:
            if name.lower().endswith(".xlsx") and not name.startswith("~$"):
                files.append(os.path.join(root, name))
    return sorted(files)


def get_table_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
    cur = conn.cursor()
    cur.execute(f'PRAGMA table_info("{table_name}")')
    return [row[1] for row in cur.fetchall()]


def validate_module(module_key: str, module_name: str, excel_path: Optional[str]) -> ModuleValidationResult:
    details: List[str] = []

    if not excel_path:
        return ModuleValidationResult(
            module_key=module_key,
            module_name=module_name,
            excel_path=None,
            status="MISSING_FILE",
            details=["No se encontro archivo .xlsx asociado al modulo."],
        )

    try:
        df = leer_excel_completo(excel_path)
        if df.empty:
            return ModuleValidationResult(
                module_key=module_key,
                module_name=module_name,
                excel_path=excel_path,
                status="EMPTY_EXCEL",
                details=["El Excel no contiene filas validas despues de leer todas las hojas."],
            )

        original_columns = list(df.columns)
        original_rows = len(df)
        details.append(f"Filas leidas del Excel: {original_rows}")
        details.append(f"Columnas leidas del Excel: {len(original_columns)}")

        for col in df.columns:
            if should_parse_as_date(col):
                df[col] = df[col].apply(parse_date_like_value)

        module_table = table_name_for_module(module_key)

        with tempfile.TemporaryDirectory() as tmp_dir:
            temp_db = os.path.join(tmp_dir, "database_copy.db")
            shutil.copy2(DB_NAME, temp_db)

            conn = sqlite3.connect(temp_db)
            try:
                before_exists = table_exists(conn, module_table)
                before_columns = get_table_columns(conn, module_table) if before_exists else []
                before_count = (
                    conn.execute(f'SELECT COUNT(*) FROM "{module_table}"').fetchone()[0]
                    if before_exists
                    else 0
                )

                if before_exists:
                    df_to_insert = align_dataframe_to_existing_table(df.copy(), conn, module_table)
                else:
                    df_to_insert = df.copy()

                inserted_rows_expected = len(df_to_insert)

                df_to_insert.to_sql(module_table, conn, if_exists="append", index=False)

                after_columns = get_table_columns(conn, module_table)
                after_count = conn.execute(f'SELECT COUNT(*) FROM "{module_table}"').fetchone()[0]
            finally:
                conn.close()

        inserted_rows_real = after_count - before_count
        details.append(f"Filas insertadas esperadas: {inserted_rows_expected}")
        details.append(f"Filas insertadas reales: {inserted_rows_real}")

        if inserted_rows_real != inserted_rows_expected:
            details.append("Diferencia en insercion detectada: posible perdida de filas.")
            return ModuleValidationResult(module_key, module_name, excel_path, "ROW_MISMATCH", details)

        if before_exists:
            new_cols = [c for c in after_columns if c not in before_columns]
            missing_cols = [c for c in before_columns if c not in df_to_insert.columns]
            if new_cols:
                details.append(f"Columnas nuevas detectadas y agregadas: {len(new_cols)}")
            if missing_cols:
                details.append(
                    "Columnas de tabla sin valor en este Excel (se rellenan como vacio): "
                    f"{len(missing_cols)}"
                )

        details.append("Validacion OK: no se detecto perdida de filas en la carga simulada.")
        return ModuleValidationResult(module_key, module_name, excel_path, "OK", details)

    except Exception as exc:
        details.append(f"Error al validar: {exc}")
        return ModuleValidationResult(module_key, module_name, excel_path, "ERROR", details)


def print_report(results: List[ModuleValidationResult]) -> int:
    print("=" * 80)
    print("VALIDACION DE MODULOS VS EXCEL")
    print("=" * 80)

    status_priority = {"ERROR": 0, "ROW_MISMATCH": 1, "MISSING_FILE": 2, "EMPTY_EXCEL": 3, "OK": 4}
    sorted_results = sorted(results, key=lambda r: status_priority.get(r.status, 99))

    for result in sorted_results:
        print("-" * 80)
        print(f"Modulo: {result.module_name} ({result.module_key})")
        print(f"Estado: {result.status}")
        print(f"Excel: {result.excel_path or 'NO ENCONTRADO'}")
        for line in result.details:
            print(f"  - {line}")

    print("-" * 80)
    total = len(results)
    ok = sum(1 for r in results if r.status == "OK")
    failed = total - ok
    print(f"Resumen: OK={ok} | Con observaciones={failed} | Total modulos={total}")

    return 0 if failed == 0 else 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Valida integridad de carga Excel por modulo.")
    parser.add_argument(
        "--excel-dir",
        default=os.getcwd(),
        help="Carpeta donde estan los Excel (se busca recursivamente).",
    )
    args = parser.parse_args()

    excel_files = get_excel_files(args.excel_dir)
    by_module: Dict[str, Optional[str]] = {}
    for key, info in MODULES.items():
        by_module[key] = find_excel_for_module(info["nombre"], excel_files)

    results = []
    for key, info in MODULES.items():
        result = validate_module(key, info["nombre"], by_module[key])
        results.append(result)

    return print_report(results)


if __name__ == "__main__":
    raise SystemExit(main())
