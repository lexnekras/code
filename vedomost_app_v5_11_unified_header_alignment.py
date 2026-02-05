# -*- coding: utf-8 -*-
"""
Ведомость (без сметной части) — редактирование в таблице + подбор ресурса + расчёты

Что реализовано по вашему ТЗ (этап 1, без сметы):
1) Редактирование ячеек прямо в таблице.
   - Добавление Группы/Работы/Ресурса добавляет ПУСТУЮ строку (без окна), далее заполняете двойным кликом.
2) Столбец "Ресурс (ресурсы)" (E):
   - Двойной клик по ячейке ресурса в строке "Ресурс" открывает окно подбора из базы "Ресурсы.xlsx"
     (поиск по наименованию).
   - Можно отменить подбор и ввести вручную.
   - При выборе из базы автоматически заполняются поля: Ед. изм. (ресурсы), Тип ресурса, Статья БДР,
     Способ расчета (затраты) (+ при желании копируем в "Способ расчета (ресурсы)").
   - Если ресурс введён вручную и такого нет в базе — при сохранении добавляем его в "Ресурсы объекта.xlsx"
     (рядом со скриптом) по форме, как в "Ресурсы.xlsx".
3) Расчёт полей по логике из "Расчет формул.xlsx" + агрегаты, которые вы описали:
   - объем_работы_L = L у родительской "Работы"
   - сумма_трудозатрат_AC = сумма AC по ресурсам в текущей работе
   - объем_работы_N = N у родительской "Работы"
   - сумма_трудозатрат_AF = сумма AF по ресурсам в текущей работе
   - общие_трудозатраты_AC/AF = сумма AC/AF по всей ведомости
   - сумма_материалов_T = сумма T по всей ведомости (исключая ресурсы со способом расчета H="По сумме материалов")
   - сумма_материалов_V = сумма V по всей ведомости (исключая ресурсы со способом расчета J="По сумме материалов")
   + пересчет итогов по работам/группам (суммирование дочерних).

Округление (как вы просили ранее):
- Цена/Стоимость -> 2
- Норма расхода  -> 4
- Прочие числа   -> 3

Файлы рядом со скриптом (желательно):
- Ресурсы.xlsx
- Расчет формул.xlsx
- Логика заполнения.xlsx (для дефолтов коэффициентов; если нет — берём 1.0)
- Выпадающие списки встроены в скрипт (без внешних файлов)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd

try:
    import openpyxl
except Exception:
    openpyxl = None


# ===================== Columns (fixed) =====================
COLUMNS: list[str] = [
    "Unnamed: 0",
    "Тип",
    "Наименование\n(затраты + смета)",
    "Ед. изм.\n(затраты + смета)",
    "Ресурс\n(ресурсы)",
    "Ед. изм.\n(ресурсы)",
    "Гр.",
    "Способ расчета\n(затраты)",
    "Норма расхода\n(затраты)",
    "Способ расчета\n(ресурсы)",
    "Норма расхода\n(ресурсы)",
    "Объем\n(затраты)",     # L
    "Объем\n(смета)",       # M (не используем)
    "Объем\n(ресурсы)",     # N
    "Цена материала\n(затраты)",  # O
    "Тип коэффициента",
    "Коэф. на материал \nБазовый:",  # Q
    "Цена материала\n(смета)",      # R (не используем)
    "Цена материала\n(ресурсы)",    # S
    "Стоимость материала\n(затраты)",  # T
    "Стоимость материала\n(смета)",    # U (не используем)
    "Стоимость материала\n(ресурсы)",  # V
    "Цена за работу\n(затраты)",       # W
    "Тип коэффициента.1",
    "Коэф. на работу \nБазовый:",      # Y
    "Цена за работу без распределения\n(смета)",
    "Цена за работу\n(смета)",
    "Цена за работу\n(ресурсы)",       # AB
    "Стоимость работ\n(затраты)",      # AC
    "Стоимость работ без распределения\n(смета)",
    "Стоимость работ\n(смета)",
    "Стоимость работ\n(ресурсы)",      # AF
    "Общая стоимость\n(затраты)",      # AG
    "Общая стоимость\n(смета)",
    "Общая стоимость\n(ресурсы)",      # AI
    "Отображение в смете",
    "Тип ресурса",
    "Статья БДР",
    "Комментарий",
]

# Column shortcuts
COL_TIP = "Тип"
COL_NAME = "Наименование\n(затраты + смета)"
COL_RES_NAME = "Ресурс\n(ресурсы)"
COL_RES_UNIT = "Ед. изм.\n(ресурсы)"
COL_UNIT = "Ед. изм.\n(затраты + смета)"
COL_METHOD_Z = "Способ расчета\n(затраты)"   # H
COL_NORM_Z = "Норма расхода\n(затраты)"      # I
COL_METHOD_R = "Способ расчета\n(ресурсы)"   # J
COL_NORM_R = "Норма расхода\n(ресурсы)"      # K

COL_VOL_Z = "Объем\n(затраты)"               # L
COL_VOL_R = "Объем\n(ресурсы)"               # N

COL_PRICE_MAT_Z = "Цена материала\n(затраты)"  # O
COL_COST_MAT_Z = "Стоимость материала\n(затраты)"  # T
COL_PRICE_WORK_Z = "Цена за работу\n(затраты)"  # W
COL_COST_WORK_Z = "Стоимость работ\n(затраты)"  # AC
COL_TOTAL_Z = "Общая стоимость\n(затраты)"      # AG

COL_PRICE_MAT_R = "Цена материала\n(ресурсы)"   # S
COL_COST_MAT_R = "Стоимость материала\n(ресурсы)"  # V
COL_PRICE_WORK_R = "Цена за работу\n(ресурсы)"  # AB
COL_COST_WORK_R = "Стоимость работ\n(ресурсы)"  # AF
COL_TOTAL_R = "Общая стоимость\n(ресурсы)"      # AI

COL_COEF_MAT = "Коэф. на материал \nБазовый:"
COL_COEF_WORK = "Коэф. на работу \nБазовый:"
COL_RES_TYPE = "Тип ресурса"
COL_BDR = "Статья БДР"
COL_COMMENT = "Комментарий"


# ===================== UI palette =====================
C_APP_BG = "#F3F1E5"
C_RIBBON_BG = "#E5E0BD"
C_TABLE_HDR = "#ECECEC"
C_ROW_ODD = "#FFFFFF"
C_ROW_EVEN = "#F6F7FB"
C_SELECT_BG = "#2B579A"
C_SELECT_FG = "#FFFFFF"
C_HILITE_YELLOW = "#F7E7A8"

ROWHEIGHT = 26


# ===================== Paths (relative to script) =====================
SCRIPT_DIR = Path(__file__).resolve().parent
RESOURCES_DB_FILE = SCRIPT_DIR / "Ресурсы.xlsx"
RESOURCES_OBJECT_FILE = SCRIPT_DIR / "Ресурсы объекта.xlsx"
FORMULAS_FILE = SCRIPT_DIR / "Расчет формул.xlsx"
FILL_LOGIC_FILE = SCRIPT_DIR / "Логика заполнения.xlsx"
TEMPLATE_XLSM = SCRIPT_DIR / "мкрс 6.xlsm"


# ===================== Built-in dropdown lists (no external dependency) =====================
DROPDOWNS: dict[str, list[str]] = {
    # Calculation methods
    "Способ расчета\n(затраты)": [
        "На объем работы",
        "По норме расхода",
        "По сумме трудозатрат",
        "По общим трудозатратам",
        "По сумме материалов",
        "По расходу денег",
    ],
    "Способ расчета\n(ресурсы)": [
        "На объем работы",
        "По норме расхода",
        "По сумме трудозатрат",
        "По общим трудозатратам",
        "По сумме материалов",
        "По расходу денег",
    ],
    # Coefficient types (if you use these columns later)
    "Тип коэффициента": [
        "Твердый",
        "Плавающий",
        "Зависимый",
    ],
    "Тип коэффициента.1": [
        "Твердый",
        "Плавающий",
        "Зависимый",
    ],
    # Display in estimate (kept for future even if estimate part disabled)
    "Отображение в смете": [
        "Показывать",
        "Распределять",
        "Не учитывать",
    ],
    # Resource types
    "Тип ресурса": [
        "Материал",
        "Трудозатраты",
        "Спецтехника",
        "Субподряд",
        "Прочее",
    ],
}


# ===================== Header labels and colors (from 'Строки заголовков.xlsx' / основная) =====================
DISPLAY_HEADERS: dict[str, str] = {
    'Unnamed: 0': '№',
    'Тип': 'Тип',
    'Наименование\n(затраты + смета)': 'Наименование',
    'Ед. изм.\n(затраты + смета)': 'Ед.изм.',
    'Ресурс\n(ресурсы)': 'Наименование',
    'Ед. изм.\n(ресурсы)': 'Ед.изм.',
    'Гр.': 'Гр.',
    'Способ расчета\n(затраты)': 'Способ расчета',
    'Норма расхода\n(затраты)': 'Норма расхода',
    'Способ расчета\n(ресурсы)': 'Способ расчета',
    'Норма расхода\n(ресурсы)': 'Норма расхода',
    'Объем\n(затраты)': 'Объем',
    'Объем\n(смета)': 'Объем',
    'Объем\n(ресурсы)': 'Объем',
    'Цена материала\n(затраты)': 'Цена МР',
    'Тип коэффициента': 'К тип МР',
    'Коэф. на материал \nБазовый:': 'К МР',
    'Цена материала\n(смета)': 'Цена МР',
    'Цена материала\n(ресурсы)': 'Цена МР',
    'Стоимость материала\n(затраты)': 'Итог МР',
    'Стоимость материала\n(смета)': 'Итог МР',
    'Стоимость материала\n(ресурсы)': 'Итог МР',
    'Цена за работу\n(затраты)': 'ЦЕНА ЗП',
    'Тип коэффициента.1': 'К тип ЗП',
    'Коэф. на работу \nБазовый:': 'К тип ЗП',
    'Цена за работу без распределения\n(смета)': 'Цена без распр.',
    'Цена за работу\n(смета)': 'ЦЕНА ЗП',
    'Цена за работу\n(ресурсы)': 'ЦЕНА ЗП',
    'Стоимость работ\n(затраты)': 'Итог ЗП',
    'Стоимость работ без распределения\n(смета)': 'Итог без распр.',
    'Стоимость работ\n(смета)': 'Итог ЗП',
    'Стоимость работ\n(ресурсы)': 'Итог ЗП',
    'Общая стоимость\n(затраты)': 'ИТОГО',
    'Общая стоимость\n(смета)': 'ИТОГО',
    'Общая стоимость\n(ресурсы)': 'ИТОГО',
    'Отображение в смете': 'Видимость',
    'Тип ресурса': 'Тип ресурса',
    'Статья БДР': 'Статья БДР',
    'Комментарий': 'Комментарий',
}

HEADER_CATEGORY: dict[str, str] = {
    'Unnamed: 0': 'Общая',
    'Тип': 'Общая',
    'Наименование\n(затраты + смета)': 'Смета',
    'Ед. изм.\n(затраты + смета)': 'Смета',
    'Ресурс\n(ресурсы)': 'Ресурсы',
    'Ед. изм.\n(ресурсы)': 'Ресурсы',
    'Гр.': 'Общая',
    'Способ расчета\n(затраты)': 'Затраты',
    'Норма расхода\n(затраты)': 'Затраты',
    'Способ расчета\n(ресурсы)': 'Ресурсы',
    'Норма расхода\n(ресурсы)': 'Ресурсы',
    'Объем\n(затраты)': 'Затраты',
    'Объем\n(смета)': 'Смета',
    'Объем\n(ресурсы)': 'Ресурсы',
    'Цена материала\n(затраты)': 'Затраты',
    'Тип коэффициента': 'Смета',
    'Коэф. на материал \nБазовый:': 'Смета',
    'Цена материала\n(смета)': 'Смета',
    'Цена материала\n(ресурсы)': 'Ресурсы',
    'Стоимость материала\n(затраты)': 'Затраты',
    'Стоимость материала\n(смета)': 'Смета',
    'Стоимость материала\n(ресурсы)': 'Ресурсы',
    'Цена за работу\n(затраты)': 'Затраты',
    'Тип коэффициента.1': 'Смета',
    'Коэф. на работу \nБазовый:': 'Смета',
    'Цена за работу без распределения\n(смета)': 'Смета',
    'Цена за работу\n(смета)': 'Смета',
    'Цена за работу\n(ресурсы)': 'Ресурсы',
    'Стоимость работ\n(затраты)': 'Затраты',
    'Стоимость работ без распределения\n(смета)': 'Смета',
    'Стоимость работ\n(смета)': 'Смета',
    'Стоимость работ\n(ресурсы)': 'Ресурсы',
    'Общая стоимость\n(затраты)': 'Затраты',
    'Общая стоимость\n(смета)': 'Смета',
    'Общая стоимость\n(ресурсы)': 'Ресурсы',
    'Отображение в смете': 'Смета',
    'Тип ресурса': 'Затраты',
    'Статья БДР': 'Ресурсы',
    'Комментарий': 'Общая',
}

HEADER_COLORS: dict[str, str] = {
    'Общая': '#E7E6E6',
    'Затраты': '#D9E1F2',
    'Смета': '#D9E1F2',
    'Ресурсы': '#FFF2CC',
}

HEADER_THEME: dict[str, str | None] = {
    'Unnamed: 0': None,
    'Тип': 'theme:7',
    'Наименование\n(затраты + смета)': 'theme:9',
    'Ед. изм.\n(затраты + смета)': 'theme:9',
    'Ресурс\n(ресурсы)': 'theme:7',
    'Ед. изм.\n(ресурсы)': 'theme:7',
    'Гр.': 'theme:7',
    'Способ расчета\n(затраты)': 'theme:9',
    'Норма расхода\n(затраты)': 'theme:9',
    'Способ расчета\n(ресурсы)': 'theme:7',
    'Норма расхода\n(ресурсы)': 'theme:7',
    'Объем\n(затраты)': 'theme:9',
    'Объем\n(смета)': 'theme:4',
    'Объем\n(ресурсы)': 'theme:7',
    'Цена материала\n(затраты)': 'theme:9',
    'Тип коэффициента': 'theme:4',
    'Коэф. на материал \nБазовый:': 'theme:4',
    'Цена материала\n(смета)': 'theme:4',
    'Цена материала\n(ресурсы)': 'theme:7',
    'Стоимость материала\n(затраты)': 'theme:9',
    'Стоимость материала\n(смета)': 'theme:4',
    'Стоимость материала\n(ресурсы)': 'theme:7',
    'Цена за работу\n(затраты)': 'theme:9',
    'Тип коэффициента.1': 'theme:4',
    'Коэф. на работу \nБазовый:': 'theme:4',
    'Цена за работу без распределения\n(смета)': 'theme:4',
    'Цена за работу\n(смета)': 'theme:4',
    'Цена за работу\n(ресурсы)': 'theme:7',
    'Стоимость работ\n(затраты)': 'theme:9',
    'Стоимость работ без распределения\n(смета)': 'theme:4',
    'Стоимость работ\n(смета)': 'theme:4',
    'Стоимость работ\n(ресурсы)': 'theme:7',
    'Общая стоимость\n(затраты)': 'theme:9',
    'Общая стоимость\n(смета)': 'theme:4',
    'Общая стоимость\n(ресурсы)': 'theme:7',
    'Отображение в смете': 'theme:4',
    'Тип ресурса': 'theme:9',
    'Статья БДР': 'theme:7',
    'Комментарий': 'theme:7',
}

HEADER_THEME_COLORS: dict[str, str] = {
    'theme:4': '#DAE3F3',
    'theme:7': '#FFF2CC',
    'theme:9': '#E2EFDA',
    None: '#E7E6E6',
}



# ===================== Rounding / formatting =====================
def decimals_for_column(col: str) -> Optional[int]:
    if "Норма расхода" in col:
        return 4
    if "Цена" in col or "Стоимость" in col:
        return 2
    non_numeric = {
        COL_TIP, COL_METHOD_Z, COL_METHOD_R,
        "Тип коэффициента", "Тип коэффициента.1",
        "Отображение в смете", COL_RES_TYPE, COL_BDR, COL_COMMENT,
        COL_UNIT, COL_RES_UNIT, COL_NAME, COL_RES_NAME, "Гр.",
    }
    if col in non_numeric:
        return None
    return 3


def try_parse_number(x: Any) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, (int, float)) and not (isinstance(x, float) and pd.isna(x)):
        return float(x)
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return None
        s = s.replace("\xa0", " ").replace(" ", "")
        s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return None
    return None


def round_value(col: str, value: Any) -> Any:
    dec = decimals_for_column(col)
    if dec is None:
        return value
    num = try_parse_number(value)
    if num is None:
        return value
    return round(num, dec)


def is_money_column(col: str) -> bool:
    """
    Денежные колонки = те, что отображаются с 2 знаками (Цена/Стоимость/Общая стоимость и т.п.).
    """
    return decimals_for_column(col) == 2


def format_number_ru_fixed(num: float, dec: int) -> str:
    """Всегда фиксированное число знаков (деньги)."""
    s = format(num, f",.{dec}f")  # 1,234,567.89
    s = s.replace(",", " ").replace(".", ",")
    return s


def format_number_ru_trim(num: float, max_dec: int) -> str:
    """До max_dec знаков, без лишних нулей в конце дробной части."""
    s = format(num, f",.{max_dec}f")  # 1,234,567.890
    if "." in s:
        int_part, frac = s.split(".", 1)
        frac = frac.rstrip("0")
        s = int_part + ("." + frac if frac else "")
    s = s.replace(",", " ").replace(".", ",")
    return s



def format_cell_for_view(col: str, value: Any) -> Any:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    dec = decimals_for_column(col)
    num = try_parse_number(value)
    if dec is None or num is None:
        return value
    num_r = round(num, dec)
    if is_money_column(col):
        # деньги всегда с 2 знаками
        return format_number_ru_fixed(num_r, 2)
    # остальные числа: до dec знаков, но без лишних нулей
    return format_number_ru_trim(num_r, dec)



def num0(x: Any) -> float:
    v = try_parse_number(x)
    return float(v) if v is not None else 0.0



# ===================== Column visibility modes (Excel-like) =====================
def excel_col_to_num(col: str) -> int:
    """A -> 1, Z -> 26, AA -> 27, ..."""
    col = col.strip().upper()
    n = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            continue
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def parse_excel_range(rng: str) -> tuple[int, int]:
    """
    "E:F" -> (5,6), "AL:AL" -> (38,38)
    """
    a, b = rng.split(":")
    return excel_col_to_num(a), excel_col_to_num(b)


COLUMN_MODE_RANGES: dict[str, list[str]] = {
    "общая": [],
    "затраты": [
        "E:F",
        "J:K",
        "M:N",
        "P:S",
        "U:V",
        "X:AB",
        "AD:AF",
        "AH:AJ",
        "AL:AL",
    ],
    "смета": [
        "E:F",
        "H:L",
        "N:O",
        "S:T",
        "V:W",
        "AB:AC",
        "AF:AG",
        "AI:AI",
        "AK:AL",
    ],
    "ресурсы": [
        "C:D",
        "H:I",
        "L:M",
        "O:R",
        "T:U",
        "W:AA",
        "AC:AE",
        "AG:AH",
        "AJ:AK",
    ],
}

# ===================== Excel helpers =====================
def read_excel_any(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheet_name = xls.sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def align_df_to_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    df2 = df.copy()
    for c in columns:
        if c not in df2.columns:
            df2[c] = None
    df2 = df2[[c for c in columns]]
    return df2


# ===================== Dropdowns (optional) =====================
def _col_values(ws, col_letter: str, rows: range) -> list[str]:
    out: list[str] = []
    for r in rows:
        v = ws[f"{col_letter}{r}"].value
        if v is None:
            continue
        s = str(v).strip()
        if s:
            out.append(s)
    return out


def extract_dropdown_lists(template_path: Path) -> dict[str, list[str]]:
    defaults = {
        "Тип коэффициента": ["Твердый", "Плавающий", "Зависимый"],
        "Тип коэффициента.1": ["Твердый", "Плавающий", "Зависимый"],
        COL_RES_TYPE: ["Материал", "Субподряд", "Спецтехника", "Трудозатраты", "Другое"],
        "Отображение в смете": ["Показывать", "Распределять", "Не учитывать"],
        COL_METHOD_Z: [
            "На объем работы", "По норме расхода", "По сумме трудозатрат",
            "По общим трудозатратам", "По сумме материалов", "По расходу денег",
        ],
        COL_METHOD_R: [
            "На объем работы", "По норме расхода", "По сумме трудозатрат",
            "По общим трудозатратам", "По сумме материалов", "По расходу денег",
        ],
    }
    if openpyxl is None or not template_path.exists():
        return defaults
    try:
        wb = openpyxl.load_workbook(template_path, data_only=True, keep_vba=True)
        if "Лист1" not in wb.sheetnames:
            return defaults
        ws = wb["Лист1"]

        coef = _col_values(ws, "A", range(1, 30))
        res_type = _col_values(ws, "C", range(1, 30))
        show = _col_values(ws, "E", range(1, 30))
        calc = _col_values(ws, "G", range(1, 80))

        def uniq(seq):
            seen = set()
            out = []
            for x in seq:
                k = x.strip()
                if k and k not in seen:
                    seen.add(k)
                    out.append(k)
            return out

        coef = uniq(coef)[:40] or defaults["Тип коэффициента"]
        res_type = uniq(res_type)[:60] or defaults[COL_RES_TYPE]
        show = uniq(show)[:40] or defaults["Отображение в смете"]
        calc = uniq(calc)[:60] or defaults[COL_METHOD_Z]

        return {
            "Тип коэффициента": coef,
            "Тип коэффициента.1": coef,
            COL_RES_TYPE: res_type,
            "Отображение в смете": show,
            COL_METHOD_Z: calc,
            COL_METHOD_R: calc,
        }
    except Exception:
        return defaults


# ===================== Defaults from "Логика заполнения.xlsx" =====================
def read_defaults_from_fill_logic(path: Path) -> dict[str, Any]:
    defaults = {COL_COEF_MAT: 1.0, COL_COEF_WORK: 1.0}
    if not path.exists():
        return defaults
    try:
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl", header=None)
        # row 1 has defaults at specific columns in your file
        # Find column indices by matching header row 0
        hdr = df.iloc[0].tolist()
        def col_idx(name: str) -> Optional[int]:
            for i, v in enumerate(hdr):
                if str(v).strip() == name:
                    return i
            return None
        i_mat = col_idx(COL_COEF_MAT)
        i_work = col_idx(COL_COEF_WORK)
        if i_mat is not None:
            v = df.iloc[1, i_mat]
            if pd.notna(v):
                defaults[COL_COEF_MAT] = float(v)
        if i_work is not None:
            v = df.iloc[1, i_work]
            if pd.notna(v):
                defaults[COL_COEF_WORK] = float(v)
        return defaults
    except Exception:
        return defaults


# ===================== Calculation rules from "Расчет формул.xlsx" =====================
@dataclass
class MethodRule:
    """Rules for one calculation method."""
    method: str
    cell_map: dict[str, str]  # column_name -> instruction/formula text


def load_method_rules(path: Path) -> dict[str, MethodRule]:
    """
    Loads rules from your "Расчет формул.xlsx" (sheet 0).
    Falls back to a minimal hardcoded set if file not found/readable.
    """
    def fallback() -> dict[str, MethodRule]:
        return {
            "На объем работы": MethodRule("На объем работы", {
                COL_COST_MAT_Z: "L*O if (L!=0 or O!=0) else пусто",
                COL_COST_WORK_Z: "L*W if (L!=0 or W!=0) else пусто",
                COL_TOTAL_Z: "T + AC",
            }),
            "По норме расхода": MethodRule("По норме расхода", {
                COL_VOL_Z: "I*объем_работы_L",
                COL_COST_MAT_Z: "L*O",
                COL_COST_WORK_Z: "L*W",
                COL_TOTAL_Z: "T + AC",
            }),
            "По сумме трудозатрат": MethodRule("По сумме трудозатрат", {
                COL_COST_MAT_Z: "Сумма_трудозатрат_AC*I",
                COL_TOTAL_Z: "T + AC",
            }),
            "По общим трудозатратам": MethodRule("По общим трудозатратам", {
                COL_COST_MAT_Z: "Общие_трудозатраты_AC*I",
                COL_TOTAL_Z: "T + AC",
            }),
            "По сумме материалов": MethodRule("По сумме материалов", {
                COL_COST_MAT_Z: "Сумма_материалов_T*I",
                COL_TOTAL_Z: "T + AC",
            }),
            "По расходу денег": MethodRule("По расходу денег", {
                COL_COST_MAT_Z: "I*объем_работы_L",
                COL_TOTAL_Z: "T + AC",
            }),
        }

    if not path.exists():
        return fallback()

    try:
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl", header=None)
        header_row = df.iloc[0].tolist()
        # Column 7 should be "Способ расчета (затраты)" in your file.
        # Rows 1..6 correspond to methods.
        rules: dict[str, MethodRule] = {}
        # Build column index -> column name for non-empty headers
        idx_to_colname = {i: header_row[i] for i in range(len(header_row)) if pd.notna(header_row[i])}
        for r in range(1, df.shape[0]):
            method = df.iloc[r, 7]
            if pd.isna(method):
                continue
            method = str(method).strip()
            if not method:
                continue
            cell_map: dict[str, str] = {}
            for i, colname in idx_to_colname.items():
                if i < 7:
                    continue
                if colname == "Способ расчета\n(затраты)":
                    continue
                txt = df.iloc[r, i]
                if pd.isna(txt):
                    continue
                cell_map[str(colname)] = str(txt).strip()
            rules[method] = MethodRule(method, cell_map)
        if rules:
            return rules
        return fallback()
    except Exception:
        return fallback()


# ===================== Model (tree) =====================
@dataclass
class Node:
    node_type: str  # "Группа" | "Работа" | "Ресурс" | "ROOT"
    row: dict[str, Any] = field(default_factory=dict)
    parent: Optional["Node"] = None
    children: list["Node"] = field(default_factory=list)

    def add_child(self, child: "Node") -> None:
        child.parent = self
        self.children.append(child)

    def remove_child(self, child: "Node") -> None:
        self.children = [c for c in self.children if c is not child]

    def ancestors(self):
        n = self.parent
        while n is not None:
            yield n
            n = n.parent

    def is_ancestor_of(self, other: "Node") -> bool:
        for a in other.ancestors():
            if a is self:
                return True
        return False


class VedomostModel:
    def __init__(self, df: pd.DataFrame, columns: list[str]):
        self.columns = list(columns)
        self.display_col = COL_NAME if COL_NAME in columns else columns[2]
        self.root = Node("ROOT", row={})
        self._build_tree_from_df(df)

    def _build_tree_from_df(self, df: pd.DataFrame) -> None:
        current_group: Optional[Node] = None
        current_work: Optional[Node] = None

        for _, row in df.iterrows():
            typ = row.get(COL_TIP, None)
            if pd.isna(typ):
                continue
            typ = str(typ).strip()
            if not typ:
                continue

            row_dict: dict[str, Any] = {}
            for c in self.columns:
                v = row.get(c, None)
                if pd.isna(v):
                    v = None
                row_dict[c] = round_value(c, v)

            node = Node(node_type=typ, row=row_dict)

            if typ == "Группа":
                self.root.add_child(node)
                current_group = node
                current_work = None
            elif typ == "Работа":
                if current_group is None:
                    self.root.add_child(node)
                else:
                    current_group.add_child(node)
                current_work = node
            elif typ == "Ресурс":
                if current_work is not None:
                    current_work.add_child(node)
                elif current_group is not None:
                    current_group.add_child(node)
                else:
                    self.root.add_child(node)
            else:
                self.root.add_child(node)

    def flatten_rows(self) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []

        def walk(n: Node):
            for ch in n.children:
                d = {c: None for c in self.columns}
                for c in self.columns:
                    d[c] = round_value(c, ch.row.get(c, None))
                d[COL_TIP] = ch.node_type
                out.append(d)
                walk(ch)

        walk(self.root)
        return out


# ===================== Resource DB =====================
class ResourceDB:
    def __init__(self, path: Path):
        self.path = path
        self.df: Optional[pd.DataFrame] = None

    def load(self) -> None:
        if not self.path.exists():
            self.df = pd.DataFrame(columns=[
                "Наименование", "Код", "Тип ресурса", "Единица измерения",
                "Способ расчета", "Трудозатраты", "Статья БДР", "Комментарий"
            ])
            return
        self.df = pd.read_excel(self.path, sheet_name=0, engine="openpyxl")

    def ensure_loaded(self):
        if self.df is None:
            self.load()

    def find_by_name(self, name: str) -> Optional[pd.Series]:
        self.ensure_loaded()
        if self.df is None or self.df.empty:
            return None
        s = name.strip().lower()
        m = self.df["Наименование"].astype(str).str.strip().str.lower() == s
        hits = self.df[m]
        if hits.empty:
            return None
        return hits.iloc[0]

    def max_code(self) -> int:
        self.ensure_loaded()
        if self.df is None or self.df.empty or "Код" not in self.df.columns:
            return 0
        codes = pd.to_numeric(self.df["Код"], errors="coerce").dropna()
        if codes.empty:
            return 0
        return int(codes.max())


def upsert_object_resources(base_db: ResourceDB, object_path: Path, model: VedomostModel) -> None:
    """
    Adds new manually-entered resources into 'Ресурсы объекта.xlsx' if not present in base.
    """
    base_db.ensure_loaded()

    if object_path.exists():
        obj_df = pd.read_excel(object_path, sheet_name=0, engine="openpyxl")
    else:
        obj_df = pd.DataFrame(columns=[
            "Наименование", "Код", "Тип ресурса", "Единица измерения",
            "Способ расчета", "Трудозатраты", "Статья БДР", "Комментарий"
        ])

    # Build set of existing names (base + object)
    existing = set()
    if base_db.df is not None and not base_db.df.empty:
        existing |= set(base_db.df["Наименование"].astype(str).str.strip().str.lower().tolist())
    if not obj_df.empty:
        existing |= set(obj_df["Наименование"].astype(str).str.strip().str.lower().tolist())

    next_code = max(base_db.max_code(), int(pd.to_numeric(obj_df["Код"], errors="coerce").dropna().max() or 0)) + 1

    new_rows = []
    # iterate nodes
    def walk(n: Node):
        nonlocal next_code
        for ch in n.children:
            if ch.node_type == "Ресурс":
                name = str(ch.row.get(COL_RES_NAME, "") or "").strip()
                if not name:
                    continue
                key = name.lower()
                if key in existing:
                    continue
                # if this resource was selected from base, we keep hidden code and skip
                if ch.row.get("__resource_code") is not None:
                    continue

                unit = str(ch.row.get(COL_RES_UNIT, "") or "").strip()
                rtype = str(ch.row.get(COL_RES_TYPE, "") or "").strip()
                method = str(ch.row.get(COL_METHOD_Z, "") or "").strip()
                bdr = str(ch.row.get(COL_BDR, "") or "").strip()
                comm = str(ch.row.get(COL_COMMENT, "") or "").strip()

                new_rows.append({
                    "Наименование": name,
                    "Код": next_code,
                    "Тип ресурса": rtype,
                    "Единица измерения": unit,
                    "Способ расчета": method,
                    "Трудозатраты": "Нет",
                    "Статья БДР": bdr,
                    "Комментарий": comm,
                })
                existing.add(key)
                next_code += 1
            walk(ch)

    walk(model.root)

    if new_rows:
        obj_df = pd.concat([obj_df, pd.DataFrame(new_rows)], ignore_index=True)
        obj_df.to_excel(object_path, index=False)


# ===================== Calculation Engine =====================
class CalcEngine:
    """
    Recalculates computed columns for Resources and totals for Works/Groups.
    """
    def __init__(self, rules: dict[str, MethodRule]):
        self.rules = rules

    def recalc_all(self, model: VedomostModel):
        # Pass 1: compute "direct" formulas for resource rows (doesn't need sums)
        for work in self._iter_works(model.root):
            work_L = num0(work.row.get(COL_VOL_Z))
            work_N = num0(work.row.get(COL_VOL_R))

            # direct for each resource
            for res in [c for c in work.children if c.node_type == "Ресурс"]:
                self._calc_resource_direct(res, work_L, work_N)

        # Pass 1 totals to compute sums
        sums = self._compute_sums(model)

        # Pass 2: apply sum-based methods
        for work in self._iter_works(model.root):
            work_L = num0(work.row.get(COL_VOL_Z))
            work_N = num0(work.row.get(COL_VOL_R))

            sum_AC_work = sums["sum_AC_work"].get(id(work), 0.0)
            sum_AF_work = sums["sum_AF_work"].get(id(work), 0.0)

            for res in [c for c in work.children if c.node_type == "Ресурс"]:
                self._calc_resource_sum_based(
                    res=res,
                    work_L=work_L,
                    work_N=work_N,
                    sum_AC_work=sum_AC_work,
                    sum_AF_work=sum_AF_work,
                    sum_AC_all=sums["sum_AC_all"],
                    sum_AF_all=sums["sum_AF_all"],
                    sum_T_all=sums["sum_T_all_excl_sum_materials"],
                    sum_V_all=sums["sum_V_all_excl_sum_materials"],
                )

        # Totals for work/group after resources final
        self._recalc_totals_recursive(model.root)

    def _iter_works(self, node: Node):
        for ch in node.children:
            if ch.node_type == "Работа":
                yield ch
            yield from self._iter_works(ch)

    def _iter_resources(self, node: Node):
        for ch in node.children:
            if ch.node_type == "Ресурс":
                yield ch
            yield from self._iter_resources(ch)

    def _compute_sums(self, model: VedomostModel) -> dict[str, Any]:
        sum_AC_work: dict[int, float] = {}
        sum_AF_work: dict[int, float] = {}

        # per work sums
        for work in self._iter_works(model.root):
            ac = 0.0
            af = 0.0
            for res in [c for c in work.children if c.node_type == "Ресурс"]:
                ac += num0(res.row.get(COL_COST_WORK_Z))
                af += num0(res.row.get(COL_COST_WORK_R))
            sum_AC_work[id(work)] = ac
            sum_AF_work[id(work)] = af

        # global sums
        sum_AC_all = 0.0
        sum_AF_all = 0.0

        sum_T_all_excl_sum_materials = 0.0
        sum_V_all_excl_sum_materials = 0.0

        for res in self._iter_resources(model.root):
            sum_AC_all += num0(res.row.get(COL_COST_WORK_Z))
            sum_AF_all += num0(res.row.get(COL_COST_WORK_R))

            # exclude "По сумме материалов" as described
            mZ = str(res.row.get(COL_METHOD_Z, "") or "").strip()
            if mZ.lower() != "по сумме материалов":
                sum_T_all_excl_sum_materials += num0(res.row.get(COL_COST_MAT_Z))

            mR = str(res.row.get(COL_METHOD_R, "") or "").strip()
            if mR.lower() != "по сумме материалов":
                sum_V_all_excl_sum_materials += num0(res.row.get(COL_COST_MAT_R))

        return {
            "sum_AC_work": sum_AC_work,
            "sum_AF_work": sum_AF_work,
            "sum_AC_all": sum_AC_all,
            "sum_AF_all": sum_AF_all,
            "sum_T_all_excl_sum_materials": sum_T_all_excl_sum_materials,
            "sum_V_all_excl_sum_materials": sum_V_all_excl_sum_materials,
        }

    def _calc_resource_direct(self, res: Node, work_L: float, work_N: float) -> None:
        """
        Direct formulas for resource row based on method, without sum-based variables.
        """
        methodZ = str(res.row.get(COL_METHOD_Z, "") or "").strip()
        methodR = str(res.row.get(COL_METHOD_R, "") or "").strip()
        if not methodR:
            methodR = methodZ  # default

        # --- ZATRATY section ---
        if methodZ == "По норме расхода":
            I = num0(res.row.get(COL_NORM_Z))
            res.row[COL_VOL_Z] = round_value(COL_VOL_Z, I * work_L)
        # Other methods: volume is manual or blocked -> we don't touch here.

        # Costs for "На объем работы" and "По норме расхода" are direct
        if methodZ in ("На объем работы", "По норме расхода"):
            self._compute_costs_z(res, conditional=(methodZ == "На объем работы"))

        if methodZ == "По расходу денег":
            I = num0(res.row.get(COL_NORM_Z))
            res.row[COL_COST_MAT_Z] = round_value(COL_COST_MAT_Z, I * work_L)
            res.row[COL_TOTAL_Z] = round_value(COL_TOTAL_Z, num0(res.row.get(COL_COST_MAT_Z)) + num0(res.row.get(COL_COST_WORK_Z)))

        # --- RESOURCES section (symmetry) ---
        if methodR == "По норме расхода":
            K = num0(res.row.get(COL_NORM_R))
            res.row[COL_VOL_R] = round_value(COL_VOL_R, K * work_N)

        if methodR in ("На объем работы", "По норме расхода"):
            self._compute_costs_r(res, conditional=(methodR == "На объем работы"))

        if methodR == "По расходу денег":
            K = num0(res.row.get(COL_NORM_R))
            res.row[COL_COST_MAT_R] = round_value(COL_COST_MAT_R, K * work_N)
            res.row[COL_TOTAL_R] = round_value(COL_TOTAL_R, num0(res.row.get(COL_COST_MAT_R)) + num0(res.row.get(COL_COST_WORK_R)))

    def _calc_resource_sum_based(
        self,
        res: Node,
        work_L: float,
        work_N: float,
        sum_AC_work: float,
        sum_AF_work: float,
        sum_AC_all: float,
        sum_AF_all: float,
        sum_T_all: float,
        sum_V_all: float,
    ) -> None:
        """
        Sum-based formulas for resource row.
        """
        methodZ = str(res.row.get(COL_METHOD_Z, "") or "").strip()
        methodR = str(res.row.get(COL_METHOD_R, "") or "").strip()
        if not methodR:
            methodR = methodZ

        # ZATRATY
        if methodZ == "По сумме трудозатрат":
            I = num0(res.row.get(COL_NORM_Z))
            res.row[COL_COST_MAT_Z] = round_value(COL_COST_MAT_Z, sum_AC_work * I)
        elif methodZ == "По общим трудозатратам":
            I = num0(res.row.get(COL_NORM_Z))
            res.row[COL_COST_MAT_Z] = round_value(COL_COST_MAT_Z, sum_AC_all * I)
        elif methodZ == "По сумме материалов":
            I = num0(res.row.get(COL_NORM_Z))
            res.row[COL_COST_MAT_Z] = round_value(COL_COST_MAT_Z, sum_T_all * I)

        # total Z always: T + AC
        if methodZ in ("По сумме трудозатрат", "По общим трудозатратам", "По сумме материалов"):
            res.row[COL_TOTAL_Z] = round_value(COL_TOTAL_Z, num0(res.row.get(COL_COST_MAT_Z)) + num0(res.row.get(COL_COST_WORK_Z)))

        # RESOURCES
        if methodR == "По сумме трудозатрат":
            K = num0(res.row.get(COL_NORM_R))
            res.row[COL_COST_MAT_R] = round_value(COL_COST_MAT_R, sum_AF_work * K)
        elif methodR == "По общим трудозатратам":
            K = num0(res.row.get(COL_NORM_R))
            res.row[COL_COST_MAT_R] = round_value(COL_COST_MAT_R, sum_AF_all * K)
        elif methodR == "По сумме материалов":
            K = num0(res.row.get(COL_NORM_R))
            res.row[COL_COST_MAT_R] = round_value(COL_COST_MAT_R, sum_V_all * K)

        if methodR in ("По сумме трудозатрат", "По общим трудозатратам", "По сумме материалов"):
            res.row[COL_TOTAL_R] = round_value(COL_TOTAL_R, num0(res.row.get(COL_COST_MAT_R)) + num0(res.row.get(COL_COST_WORK_R)))

    def _compute_costs_z(self, res: Node, conditional: bool):
        L = num0(res.row.get(COL_VOL_Z))
        O = num0(res.row.get(COL_PRICE_MAT_Z))
        W = num0(res.row.get(COL_PRICE_WORK_Z))

        if conditional and abs(L) < 1e-12 and abs(O) < 1e-12:
            res.row[COL_COST_MAT_Z] = None
        else:
            res.row[COL_COST_MAT_Z] = round_value(COL_COST_MAT_Z, L * O)

        if conditional and abs(L) < 1e-12 and abs(W) < 1e-12:
            res.row[COL_COST_WORK_Z] = None
        else:
            res.row[COL_COST_WORK_Z] = round_value(COL_COST_WORK_Z, L * W)

        res.row[COL_TOTAL_Z] = round_value(COL_TOTAL_Z, num0(res.row.get(COL_COST_MAT_Z)) + num0(res.row.get(COL_COST_WORK_Z)))

    def _compute_costs_r(self, res: Node, conditional: bool):
        N = num0(res.row.get(COL_VOL_R))
        S = num0(res.row.get(COL_PRICE_MAT_R))
        AB = num0(res.row.get(COL_PRICE_WORK_R))

        if conditional and abs(N) < 1e-12 and abs(S) < 1e-12:
            res.row[COL_COST_MAT_R] = None
        else:
            res.row[COL_COST_MAT_R] = round_value(COL_COST_MAT_R, N * S)

        if conditional and abs(N) < 1e-12 and abs(AB) < 1e-12:
            res.row[COL_COST_WORK_R] = None
        else:
            res.row[COL_COST_WORK_R] = round_value(COL_COST_WORK_R, N * AB)

        res.row[COL_TOTAL_R] = round_value(COL_TOTAL_R, num0(res.row.get(COL_COST_MAT_R)) + num0(res.row.get(COL_COST_WORK_R)))

    def _recalc_totals_recursive(self, node: Node):
        # post-order
        for ch in node.children:
            self._recalc_totals_recursive(ch)

        if node.node_type in ("Работа", "Группа"):
            # sum from children (resources and nested)
            t = 0.0
            ac = 0.0
            v = 0.0
            af = 0.0

            for ch in node.children:
                if ch.node_type == "Ресурс":
                    t += num0(ch.row.get(COL_COST_MAT_Z))
                    ac += num0(ch.row.get(COL_COST_WORK_Z))
                    v += num0(ch.row.get(COL_COST_MAT_R))
                    af += num0(ch.row.get(COL_COST_WORK_R))
                else:
                    # take totals from child group/work (already computed)
                    t += num0(ch.row.get(COL_COST_MAT_Z))
                    ac += num0(ch.row.get(COL_COST_WORK_Z))
                    v += num0(ch.row.get(COL_COST_MAT_R))
                    af += num0(ch.row.get(COL_COST_WORK_R))

            node.row[COL_COST_MAT_Z] = round_value(COL_COST_MAT_Z, t)
            node.row[COL_COST_WORK_Z] = round_value(COL_COST_WORK_Z, ac)
            node.row[COL_TOTAL_Z] = round_value(COL_TOTAL_Z, t + ac)

            node.row[COL_COST_MAT_R] = round_value(COL_COST_MAT_R, v)
            node.row[COL_COST_WORK_R] = round_value(COL_COST_WORK_R, af)
            node.row[COL_TOTAL_R] = round_value(COL_TOTAL_R, v + af)


# ===================== Resource Search Dialog =====================
class ResourceSearchDialog(tk.Toplevel):
    def __init__(self, master: tk.Widget, db: ResourceDB):
        super().__init__(master)
        self.db = db
        self.db.ensure_loaded()
        self.result: Optional[dict[str, Any]] = None
        self.manual = False

        self.title("Подбор ресурса")
        self.geometry("900x600")
        self.minsize(700, 450)
        self.configure(bg="#FFFFFF")
        self.transient(master)
        self.grab_set()

        top = tk.Frame(self, bg="#FFFFFF")
        top.pack(fill="x", padx=10, pady=10)

        tk.Label(top, text="Поиск:", bg="#FFFFFF", fg="#222222",
                 font=("Segoe UI", 10)).pack(side="left")
        self.e = tk.Entry(top, font=("Segoe UI", 10))
        self.e.pack(side="left", fill="x", expand=True, padx=8)
        self.e.bind("<KeyRelease>", lambda e: self._apply_filter())
        self.e.focus_set()

        btn_manual = tk.Button(top, text="Ввести вручную", command=self._manual)
        btn_manual.pack(side="right")

        mid = tk.Frame(self, bg="#FFFFFF")
        mid.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        mid.grid_rowconfigure(0, weight=1)
        mid.grid_columnconfigure(0, weight=1)

        cols = ["Код", "Наименование", "Тип ресурса", "Единица измерения", "Способ расчета", "Трудозатраты", "Статья БДР"]
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for c in cols:
            self.tree.heading(c, text=c, anchor="w")
            self.tree.column(c, anchor="w", width=140 if c != "Наименование" else 300, stretch=True)

        self.tree.bind("<Double-1>", lambda e: self._choose())
        self.tree.bind("<Return>", lambda e: self._choose())
        self.tree.bind("<Escape>", lambda e: self._cancel())

        bottom = tk.Frame(self, bg="#FFFFFF")
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(bottom, text="Отмена", command=self._cancel).pack(side="right", padx=(6, 0))
        tk.Button(bottom, text="Выбрать", command=self._choose).pack(side="right")

        self._rows_view = self.db.df.copy() if self.db.df is not None else pd.DataFrame()
        self._populate()

    def _populate(self):
        for iid in self.tree.get_children(""):
            self.tree.delete(iid)
        if self._rows_view is None or self._rows_view.empty:
            return
        # show only rows where Наименование not empty
        for _, r in self._rows_view.iterrows():
            name = str(r.get("Наименование", "") or "").strip()
            if not name:
                continue
            self.tree.insert("", "end", values=[
                r.get("Код", ""),
                name,
                r.get("Тип ресурса", ""),
                r.get("Единица измерения", ""),
                r.get("Способ расчета", ""),
                r.get("Трудозатраты", ""),
                r.get("Статья БДР", ""),
            ])

    def _apply_filter(self):
        q = self.e.get().strip().lower()
        if self.db.df is None:
            return
        if not q:
            self._rows_view = self.db.df
        else:
            s = self.db.df["Наименование"].astype(str).str.lower()
            self._rows_view = self.db.df[s.str.contains(q, na=False)]
        self._populate()

    def _manual(self):
        self.manual = True
        self.result = None
        self.destroy()

    def _choose(self):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        # reconstruct dict by matching to db row (by code + name)
        code = vals[0]
        name = vals[1]
        row = None
        if self.db.df is not None and not self.db.df.empty:
            m = (self.db.df["Наименование"].astype(str).str.strip() == str(name).strip())
            if "Код" in self.db.df.columns:
                m = m & (self.db.df["Код"].astype(str).str.strip() == str(code).strip())
            hits = self.db.df[m]
            if not hits.empty:
                row = hits.iloc[0]
        if row is None:
            return
        self.result = {
            "Код": row.get("Код", None),
            "Наименование": row.get("Наименование", ""),
            "Тип ресурса": row.get("Тип ресурса", ""),
            "Единица измерения": row.get("Единица измерения", ""),
            "Способ расчета": row.get("Способ расчета", ""),
            "Трудозатраты": row.get("Трудозатраты", ""),
            "Статья БДР": row.get("Статья БДР", ""),
            "Комментарий": row.get("Комментарий", ""),
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.manual = False
        self.destroy()


# ===================== Main App =====================
class OneCLikeApp(tk.Tk):
    def __init__(
        self,
        model: VedomostModel,
        dropdowns: dict[str, list[str]],
        current_path: Optional[Path],
        resource_db: ResourceDB,
        defaults: dict[str, Any],
        calc: CalcEngine,
    ):
        super().__init__()
        self.model = model
        self._header_redraw_after_id = None
        self.dropdowns = dropdowns
        self.current_path = current_path
        self.resource_db = resource_db
        self.defaults = defaults
        self.calc = calc

        self.title("Ресурсная ведомость (без сметы)")
        self.geometry("1750x900")
        self.minsize(1250, 680)
        self.configure(bg=C_APP_BG)

        self.node_by_iid: dict[str, Node] = {}
        self._drag_iid: Optional[str] = None

        self._editor: Optional[tk.Widget] = None
        self._edit_iid: Optional[str] = None
        self._edit_col: Optional[str] = None

        self._build_layout()
        self._recalc_and_refresh(full=True)

        self.bind_all("<Control-s>", lambda e: self._save())
        self.bind_all("<F5>", lambda e: self._recalc_and_refresh(full=False))
        self.bind_all("<Escape>", lambda e: self._close_editor())

    # -------- UI ----------
    def _style_ttk(self):
        style = ttk.Style(self)
        for theme in ("vista", "xpnative", "clam", "alt", "default"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure(".", font=("Segoe UI", 10))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=ROWHEIGHT,
                        background=C_ROW_ODD, fieldbackground=C_ROW_ODD, borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background=C_TABLE_HDR)
        style.map("Treeview", background=[("selected", C_SELECT_BG)], foreground=[("selected", C_SELECT_FG)])

    def _build_layout(self):
        self._style_ttk()
        main = tk.Frame(self, bg=C_APP_BG)
        main.pack(fill="both", expand=True)

        ribbon = tk.Frame(main, bg=C_RIBBON_BG, height=54)
        ribbon.pack(fill="x")
        ribbon.pack_propagate(False)

        def ribbon_btn(text: str, command, primary: bool = False):
            bg = "#F2D64A" if primary else "#F7F7F7"
            abg = "#E9C93F" if primary else "#ECECEC"
            return tk.Button(ribbon, text=text,
                             font=("Segoe UI", 10, "bold" if primary else "normal"),
                             bg=bg, activebackground=abg,
                             relief="solid", bd=1, padx=10, pady=6,
                             cursor="hand2", command=command)

        ribbon_btn("Записать", self._save, primary=True).pack(side="left", padx=(10, 6), pady=10)
        ribbon_btn("Сохранить как…", self._save_as).pack(side="left", padx=6, pady=10)
        ribbon_btn("Пересчитать (F5)", lambda: self._recalc_and_refresh(full=False)).pack(side="left", padx=6, pady=10)

        self.lbl_title = tk.Label(main, text="", bg=C_APP_BG, fg="#1e1e1e",
                                  font=("Segoe UI", 14, "bold"), anchor="w", padx=12, pady=10)
        self.lbl_title.pack(fill="x")

        toolbar = tk.Frame(main, bg=C_APP_BG)
        toolbar.pack(fill="x", padx=10, pady=(0, 6))

        def tool_btn(text: str, cmd):
            b = tk.Button(toolbar, text=text, font=("Segoe UI", 10),
                          bg="#F7F7F7", activebackground="#ECECEC",
                          relief="solid", bd=1, padx=10, pady=4, command=cmd)
            b.pack(side="left", padx=(0, 6))
            return b

        tool_btn("Добавить группу", self._add_group)
        tool_btn("Добавить работу", self._add_work)
        tool_btn("Добавить ресурс", self._add_resource)
        tool_btn("Удалить", self._delete_selected)

        # Режим отображения колонок (как в VBA: общая/затраты/смета/ресурсы)
        mode_bar = tk.Frame(main, bg=C_APP_BG)
        mode_bar.pack(fill="x", padx=10, pady=(0, 4))

        tk.Label(mode_bar, text="Режим колонок:", bg=C_APP_BG, fg="#333333",
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        self.col_mode = tk.StringVar(value="общая")

        def add_mode(text_ru: str, value: str):
            rb = tk.Radiobutton(
                mode_bar, text=text_ru, value=value, variable=self.col_mode,
                bg=C_APP_BG, fg="#333333", selectcolor=C_APP_BG,
                activebackground=C_APP_BG, font=("Segoe UI", 10),
                command=self._apply_column_mode
            )
            rb.pack(side="left", padx=(10, 0))

        add_mode("Общая", "общая")
        add_mode("Затраты", "затраты")
        add_mode("Смета", "смета")
        add_mode("Ресурсы", "ресурсы")

        info = info = ("Редактирование: двойной клик по ячейке. "
                "Ресурс (E): двойной клик -> подбор из 'Ресурсы.xlsx' (или ручной ввод). "
                "DnD: Ресурс→Работа; Работа/Группа→Группа. Ctrl+S: сохранить.")
        tk.Label(main, text=info, bg=C_APP_BG, fg="#444444", anchor="w", padx=12).pack(fill="x")

        table_frame = tk.Frame(main, bg=C_APP_BG)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        table_frame.grid_rowconfigure(1, weight=1)  # row 0 = header, row 1 = tree
        table_frame.grid_columnconfigure(0, weight=1)

        # Кастомная "шапка" (Canvas), чтобы красить заголовки по колонкам
        self.header_canvas = tk.Canvas(
            table_frame, height=ROWHEIGHT + 6, bg=C_TABLE_HDR, highlightthickness=0
        )
        self.header_canvas.grid(row=0, column=0, sticky="ew")

        # Заглушка справа под вертикальный скроллбар (чтобы шапка не "заезжала")
        self._hdr_pad = tk.Frame(table_frame, width=16, bg=C_TABLE_HDR)
        self._hdr_pad.grid(row=0, column=1, sticky="ns")

        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        hsb = ttk.Scrollbar(table_frame, orient="horizontal")
        self._hsb = hsb

        self.visible_columns = [c for c in self.model.columns if c != self.model.display_col]

        self.tree = ttk.Treeview(
            table_frame,
            columns=self.visible_columns,
            show="tree",
            yscrollcommand=vsb.set,
            xscrollcommand=self._on_xscroll,
            selectmode="browse",
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self._on_hscroll)

        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")

        self.header_canvas.bind("<Configure>", lambda e: self._schedule_header_redraw())
        self.tree.bind("<Configure>", lambda e: self._schedule_header_redraw())


        self.tree.heading("#0", text=DISPLAY_HEADERS.get(self.model.display_col, self.model.display_col.replace("\n", " ")), anchor="w")
        self.tree.column("#0", anchor="w", width=520, minwidth=260, stretch=True)

        for col in self.visible_columns:
            internal = col
            anchor = "e" if is_money_column(internal) else "w"
            self.tree.heading(col, text=DISPLAY_HEADERS.get(internal, internal.replace("\n", " ")), anchor=anchor)
            self.tree.column(col, anchor=anchor, width=160, minwidth=90, stretch=True)

        # Сохраняем стандартные ширины, чтобы восстанавливать после "скрытия" колонок
        self._default_col_conf = {}
        self._default_col_conf["#0"] = dict(
            width=self.tree.column("#0", "width"),
            minwidth=self.tree.column("#0", "minwidth"),
            stretch=self.tree.column("#0", "stretch"),
            anchor=self.tree.column("#0", "anchor"),
        )
        for col in self.visible_columns:
            self._default_col_conf[col] = dict(
                width=self.tree.column(col, "width"),
                minwidth=self.tree.column(col, "minwidth"),
                stretch=self.tree.column(col, "stretch"),
                anchor=self.tree.column(col, "anchor"),
            )

        self.after(0, self._schedule_header_redraw)

        self.tree.tag_configure("odd", background=C_ROW_ODD)
        self.tree.tag_configure("even", background=C_ROW_EVEN)
        self.tree.tag_configure("hilite", background=C_HILITE_YELLOW)
        self.tree.tag_configure("group", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("work", font=("Segoe UI", 10, "bold"))

        self.tree.bind("<ButtonPress-1>", self._on_press)
        self.tree.bind("<B1-Motion>", self._on_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_release)
        self.tree.bind("<Double-1>", self._on_double_click)

        # применим режим колонок по умолчанию
        self.after(0, self._apply_column_mode)

    def _apply_column_mode(self):
        """
        Hides/shows columns according to COLUMN_MODE_RANGES, analogous to VBA Hidden=True.
        Реализовано через width=0 + minwidth=0 + stretch=False.
        """
        mode = (self.col_mode.get() or "общая").strip().lower()
        ranges = COLUMN_MODE_RANGES.get(mode, [])

        # Build hidden set of excel indices (1-based)
        hidden_idx = set()
        for rng in ranges:
            a, b = parse_excel_range(rng)
            for i in range(a, b + 1):
                hidden_idx.add(i)

        # Map excel index -> tree column id
        # Excel columns correspond to COLUMNS list order:
        # index=1 -> COLUMNS[0], ...
        def col_id_by_excel_index(idx: int) -> Optional[str]:
            if idx < 1 or idx > len(self.model.columns):
                return None
            col_name = self.model.columns[idx - 1]
            if col_name == self.model.display_col:
                return "#0"
            return col_name

        # Apply show/hide
        for idx in range(1, len(self.model.columns) + 1):
            cid = col_id_by_excel_index(idx)
            if cid is None:
                continue
            # If a column isn't in the tree (e.g., display col handled by #0), ok.
            if cid != "#0" and cid not in self._default_col_conf:
                continue

            if idx in hidden_idx:
                # hide
                self.tree.column(cid, width=0, minwidth=0, stretch=False)
                if cid != "#0":
                    self.tree.heading(cid, text="")
                else:
                    self.tree.heading("#0", text="")
            else:
                # show restore
                conf = self._default_col_conf.get(cid)
                if conf:
                    self.tree.column(cid, anchor=conf.get("anchor","w"), width=conf["width"], minwidth=conf["minwidth"], stretch=conf["stretch"])
                # restore heading text
                if cid == "#0":
                    self.tree.heading("#0", text=DISPLAY_HEADERS.get(self.model.display_col, self.model.display_col.replace("\n", " ")), anchor="w")
                else:
                    a = "e" if is_money_column(cid) else "w"
                    self.tree.heading(cid, text=DISPLAY_HEADERS.get(cid, cid.replace("\n", " ")), anchor=a)
        self._schedule_header_redraw()


    # -------- render / refresh ----------
    def _tags_for_node(self, node: Node, stripe_idx: int) -> tuple[str, ...]:
        if node.node_type == "Группа":
            return ("hilite", "group")
        if node.node_type == "Работа":
            return ("hilite", "work")
        base = "even" if stripe_idx % 2 == 0 else "odd"
        return (base,)

    def _node_text(self, node: Node) -> str:
        v = node.row.get(self.model.display_col, "")
        if v is None:
            v = ""
        s = str(v).strip()
        return s if s else f"[{node.node_type}]"

    def _node_values(self, node: Node) -> list[Any]:
        out = []
        for c in self.visible_columns:
            if c == COL_TIP:
                out.append(node.node_type)
            else:
                v = node.row.get(c, None)
                out.append(format_cell_for_view(c, v))
        return out

    def _render_tree(self):
        for iid in self.tree.get_children(""):
            self.tree.delete(iid)
        self.node_by_iid.clear()

        def render(parent_iid: str, node: Node, stripe_state: list[int]):
            for ch in node.children:
                iid = self.tree.insert(parent_iid, "end",
                                       text=self._node_text(ch),
                                       values=self._node_values(ch),
                                       tags=self._tags_for_node(ch, stripe_state[0]),
                                       open=True)
                self.node_by_iid[iid] = ch
                stripe_state[0] += 1
                render(iid, ch, stripe_state)

        render("", self.model.root, [0])

    def _refresh_rows_inplace(self):
        # update values/text for existing nodes
        # (simple approach: rebuild tree, preserving selection if possible)
        sel_iid = self._selected_iid()
        sel_node = self.node_by_iid.get(sel_iid) if sel_iid else None
        open_nodes = {iid for iid in self.tree.get_children("") if self.tree.item(iid, "open")}
        # full rebuild to keep consistent ordering
        self._render_tree()
        # restore selection by node identity
        if sel_node is not None:
            for iid, n in self.node_by_iid.items():
                if n is sel_node:
                    self.tree.selection_set(iid)
                    self.tree.see(iid)
                    break

    def _recalc_and_refresh(self, full: bool):
        self._close_editor()
        # recalc
        self.calc.recalc_all(self.model)
        if full:
            self._render_tree()
        else:
            self._refresh_rows_inplace()
        self._update_title()

    def _update_title(self):
        name = self.current_path.name if self.current_path else "Новая ведомость (не сохранена)"
        self.lbl_title.configure(text=f"Ресурсная ведомость (без сметы)   {name}")

    # -------- selection helpers ----------
    def _selected_iid(self) -> Optional[str]:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _selected_node(self) -> Optional[Node]:
        iid = self._selected_iid()
        return self.node_by_iid.get(iid) if iid else None

    def _nearest_ancestor_of_type(self, node: Node, typ: str) -> Optional[Node]:
        if node.node_type == typ:
            return node
        for a in node.ancestors():
            if a.node_type == typ:
                return a
        return None

    def _iid_of_node(self, node: Node) -> Optional[str]:
        for iid, n in self.node_by_iid.items():
            if n is node:
                return iid
        return None

    # -------- add rows (empty) ----------
    def _blank_row(self, node_type: str) -> dict[str, Any]:
        row = {c: None for c in self.model.columns}
        row[COL_TIP] = node_type
        # defaults for coefficients (if present)
        row[COL_COEF_MAT] = self.defaults.get(COL_COEF_MAT, 1.0)
        row[COL_COEF_WORK] = self.defaults.get(COL_COEF_WORK, 1.0)
        return row

    def _add_group(self):
        self._close_editor()
        selected = self._selected_node()
        parent = self.model.root
        if selected and selected.node_type == "Группа":
            parent = selected

        node = Node("Группа", row=self._blank_row("Группа"))
        parent.add_child(node)
        self._recalc_and_refresh(full=True)
        # select new
        iid = self._iid_of_node(node)
        if iid:
            self.tree.selection_set(iid)
            self.tree.see(iid)

    def _add_work(self):
        self._close_editor()
        selected = self._selected_node()
        if selected is None:
            messagebox.showwarning("Добавить работу", "Выберите группу (или элемент внутри группы).")
            return
        grp = self._nearest_ancestor_of_type(selected, "Группа")
        if grp is None:
            messagebox.showwarning("Добавить работу", "Работу можно добавить только внутри группы.")
            return
        node = Node("Работа", row=self._blank_row("Работа"))
        grp.add_child(node)
        self._recalc_and_refresh(full=True)
        iid = self._iid_of_node(node)
        if iid:
            self.tree.selection_set(iid)
            self.tree.see(iid)

    def _add_resource(self):
        self._close_editor()
        selected = self._selected_node()
        if selected is None:
            messagebox.showwarning("Добавить ресурс", "Выберите работу (или ресурс внутри работы).")
            return
        work = self._nearest_ancestor_of_type(selected, "Работа")
        if work is None:
            messagebox.showwarning("Добавить ресурс", "Ресурс можно добавить только внутри работы.")
            return
        node = Node("Ресурс", row=self._blank_row("Ресурс"))
        work.add_child(node)
        self._recalc_and_refresh(full=True)
        iid = self._iid_of_node(node)
        if iid:
            self.tree.selection_set(iid)
            self.tree.see(iid)

    # -------- delete ----------
    def _delete_selected(self):
        self._close_editor()
        iid = self._selected_iid()
        if not iid:
            return
        node = self.node_by_iid.get(iid)
        if node is None or node.parent is None:
            return
        if node.children:
            if not messagebox.askyesno("Удалить", "У элемента есть вложенные элементы. Удалить вместе с ними?"):
                return
        node.parent.remove_child(node)
        self._recalc_and_refresh(full=True)

    # -------- drag & drop ----------
    def _on_press(self, event):
        self._close_editor()
        iid = self.tree.identify_row(event.y)
        if not iid:
            self._drag_iid = None
            return
        self.tree.selection_set(iid)
        self._drag_iid = iid

    def _on_motion(self, event):
        if not self._drag_iid:
            return
        target = self.tree.identify_row(event.y)
        if target:
            self.tree.focus(target)

    def _on_release(self, event):
        if not self._drag_iid:
            return
        source_iid = self._drag_iid
        self._drag_iid = None

        target_iid = self.tree.identify_row(event.y)
        if not target_iid or target_iid == source_iid:
            return

        src_node = self.node_by_iid.get(source_iid)
        tgt_node = self.node_by_iid.get(target_iid)
        if src_node is None or tgt_node is None:
            return

        if src_node.node_type == "Ресурс":
            container = self._nearest_ancestor_of_type(tgt_node, "Работа")
            if container is None:
                messagebox.showwarning("Перемещение", "Ресурс можно перемещать только в 'Работа'.")
                return
        elif src_node.node_type in ("Работа", "Группа"):
            container = self._nearest_ancestor_of_type(tgt_node, "Группа")
            if container is None:
                messagebox.showwarning("Перемещение", "Работу/группу можно перемещать только в 'Группа'.")
                return
        else:
            return

        if src_node is container or src_node.is_ancestor_of(container):
            messagebox.showwarning("Перемещение", "Нельзя переместить элемент внутрь самого себя (или потомка).")
            return

        # move in model
        if src_node.parent is None:
            return
        src_node.parent.remove_child(src_node)
        container.add_child(src_node)

        self._recalc_and_refresh(full=True)

    # -------- saving ----------
    def _save(self):
        self._close_editor()
        # update resources-object db before save
        upsert_object_resources(self.resource_db, RESOURCES_OBJECT_FILE, self.model)

        if self.current_path is None:
            self._save_as()
            return
        self._save_to_path(self.current_path)

    def _save_as(self):
        self._close_editor()
        path_str = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path_str:
            return
        path = Path(path_str)
        ok = self._save_to_path(path)
        if ok:
            self.current_path = path
            self._update_title()

    def _save_to_path(self, path: Path) -> bool:
        rows = self.model.flatten_rows()
        out_df = pd.DataFrame(rows, columns=self.model.columns)
        try:
            out_df.to_excel(path, index=False)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл:\n{path}\n\n{e}")
            return False
        return True

    # -------- cell edit restrictions ----------
    def _is_computed_cell(self, node: Node, col: str) -> bool:
        """
        Very first version of locking:
        - For Work/Group: totals are computed (T, AC, AG, V, AF, AI) -> lock
        - For Resource: computed fields depend on method:
            - cost and total fields always computed: T, AC, AG, V, AF, AI
            - volume L/N computed when method "По норме расхода" (for corresponding section)
            - material cost T/V computed for sum-based methods and "По расходу денег" (for corresponding section)
        """
        if col in (COL_TIP,):
            return True

        if node.node_type in ("Работа", "Группа"):
            if col in (COL_COST_MAT_Z, COL_COST_WORK_Z, COL_TOTAL_Z, COL_COST_MAT_R, COL_COST_WORK_R, COL_TOTAL_R):
                return True
            return False

        if node.node_type != "Ресурс":
            return False

        # always computed
        if col in (COL_COST_MAT_Z, COL_COST_WORK_Z, COL_TOTAL_Z, COL_COST_MAT_R, COL_COST_WORK_R, COL_TOTAL_R):
            return True

        mZ = str(node.row.get(COL_METHOD_Z, "") or "").strip()
        mR = str(node.row.get(COL_METHOD_R, "") or "").strip()

        if col == COL_VOL_Z and mZ == "По норме расхода":
            return True
        if col == COL_VOL_R and (mR == "По норме расхода" or (not mR and mZ == "По норме расхода")):
            return True

        # for methods where T is computed from sums/money, it is computed
        if col == COL_COST_MAT_Z and mZ in ("По сумме трудозатрат", "По общим трудозатратам", "По сумме материалов", "По расходу денег"):
            return True

        return False

    # -------- editing ----------
    def _on_double_click(self, event):
        self._close_editor()

        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        col_id = self.tree.identify_column(event.x)
        region = self.tree.identify("region", event.x, event.y)
        if region not in ("tree", "cell"):
            return

        node = self.node_by_iid.get(iid)
        if node is None:
            return

        if col_id == "#0":
            col_name = "#0"
            current = self.tree.item(iid, "text")
        else:
            idx = int(col_id[1:]) - 1
            if idx < 0 or idx >= len(self.visible_columns):
                return
            col_name = self.visible_columns[idx]
            if col_name == COL_TIP:
                return
            current = self.tree.set(iid, col_name)

        # Special: resource column (E) on resource rows -> open search dialog
        if node.node_type == "Ресурс" and col_name == COL_RES_NAME:
            self._open_resource_picker(node)
            return

        # Lock computed cells
        if col_name != "#0" and self._is_computed_cell(node, col_name):
            messagebox.showinfo("Поле рассчитывается", "Это поле рассчитывается автоматически и не редактируется.")
            return

        bbox = self.tree.bbox(iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        self._edit_iid = iid
        self._edit_col = col_name

        # Dropdown support
        options = None
        if col_name != "#0":
            options = self.dropdowns.get(col_name)

        if options:
            cb = ttk.Combobox(self.tree, values=options, state="readonly")
            cb.set(str(current) if current is not None else "")
            cb.place(x=x, y=y, width=w, height=h)
            cb.focus_set()
            cb.bind("<<ComboboxSelected>>", lambda e: self._commit_editor())
            cb.bind("<Return>", lambda e: self._commit_editor())
            cb.prevent_default = True
            cb.bind("<Escape>", lambda e: self._close_editor())
            cb.bind("<FocusOut>", lambda e: self._commit_editor())
            self._editor = cb
        else:
            ent = tk.Entry(self.tree, relief="solid", bd=1)
            ent.insert(0, str(current) if current is not None else "")
            ent.place(x=x, y=y, width=w, height=h)
            ent.focus_set()
            ent.bind("<Return>", lambda e: self._commit_editor())
            ent.bind("<Escape>", lambda e: self._close_editor())
            ent.bind("<FocusOut>", lambda e: self._commit_editor())
            self._editor = ent

    def _open_resource_picker(self, node: Node):
        dlg = ResourceSearchDialog(self, self.resource_db)
        self.wait_window(dlg)
        if dlg.manual:
            # fall back to manual editor on the cell
            # emulate double click but open entry editor directly
            iid = self._iid_of_node(node)
            if iid:
                self.tree.selection_set(iid)
            # open entry in E column
            self._open_manual_editor_for_cell(node, COL_RES_NAME)
            return
        if dlg.result is None:
            return

        data = dlg.result
        # apply mapping
        node.row["__resource_code"] = data.get("Код", None)
        node.row[COL_RES_NAME] = data.get("Наименование", "")
        node.row[COL_RES_UNIT] = data.get("Единица измерения", "")
        # keep D unit in sync if empty
        if not str(node.row.get(COL_UNIT, "") or "").strip():
            node.row[COL_UNIT] = data.get("Единица измерения", "")

        node.row[COL_RES_TYPE] = data.get("Тип ресурса", "")
        node.row[COL_BDR] = data.get("Статья БДР", "")
        if data.get("Комментарий"):
            node.row[COL_COMMENT] = data.get("Комментарий", "")

        method = str(data.get("Способ расчета", "") or "").strip()
        if method:
            node.row[COL_METHOD_Z] = method
            # заполним и J (ресурсы) по умолчанию так же
            if not str(node.row.get(COL_METHOD_R, "") or "").strip():
                node.row[COL_METHOD_R] = method

        self._recalc_and_refresh(full=False)

    def _open_manual_editor_for_cell(self, node: Node, col: str):
        iid = self._iid_of_node(node)
        if iid is None:
            return
        # find column id
        if col == self.model.display_col:
            col_id = "#0"
        else:
            if col not in self.visible_columns:
                return
            idx = self.visible_columns.index(col) + 1
            col_id = f"#{idx}"
        bbox = self.tree.bbox(iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        self._edit_iid = iid
        self._edit_col = col
        ent = tk.Entry(self.tree, relief="solid", bd=1)
        cur = self.tree.set(iid, col) if col_id != "#0" else self.tree.item(iid, "text")
        ent.insert(0, str(cur))
        ent.place(x=x, y=y, width=w, height=h)
        ent.focus_set()
        ent.bind("<Return>", lambda e: self._commit_editor())
        ent.bind("<Escape>", lambda e: self._close_editor())
        ent.bind("<FocusOut>", lambda e: self._commit_editor())
        self._editor = ent

    def _commit_editor(self):
        if not self._editor or not self._edit_iid or not self._edit_col:
            self._close_editor()
            return

        iid = self._edit_iid
        col = self._edit_col
        node = self.node_by_iid.get(iid)
        if node is None:
            self._close_editor()
            return

        raw = self._editor.get().strip()

        if col == "#0":
            node.row[self.model.display_col] = raw
        else:
            # store rounded numeric if possible; keep strings as is
            v_num = try_parse_number(raw)
            if v_num is not None and decimals_for_column(col) is not None:
                node.row[col] = round_value(col, v_num)
            else:
                node.row[col] = raw

            # if user changed method, clear computed columns to avoid stale
            if col in (COL_METHOD_Z, COL_METHOD_R):
                # clean totals/costs for resource to be recalculated
                if node.node_type == "Ресурс":
                    for c in (COL_VOL_Z, COL_COST_MAT_Z, COL_COST_WORK_Z, COL_TOTAL_Z,
                              COL_VOL_R, COL_COST_MAT_R, COL_COST_WORK_R, COL_TOTAL_R):
                        # keep manual volume for "На объем работы": don't wipe L/N
                        if c in (COL_VOL_Z, COL_VOL_R) and str(raw).strip() == "На объем работы":
                            continue
                        node.row[c] = node.row.get(c, None)

        self._close_editor()
        self._recalc_and_refresh(full=False)

    def _close_editor(self):
        if self._editor is not None:
            try:
                self._editor.destroy()
            except Exception:
                pass
        self._editor = None
        self._edit_iid = None
        self._edit_col = None




    # -------- header sync helpers (robust) ----------
    def _schedule_header_redraw(self):
        """
        Debounced redraw of custom header (если используется).
        """
        try:
            if getattr(self, "_header_redraw_after_id", None) is not None:
                self.after_cancel(self._header_redraw_after_id)
        except Exception:
            pass
        self._header_redraw_after_id = self.after(30, self._redraw_header_safe)

    def _redraw_header_safe(self):
        self._header_redraw_after_id = None
        fn = getattr(self, "_redraw_header", None)
        if callable(fn):
            try:
                fn()
            except Exception:
                pass
# ===================== Startup flow =====================
def start_flow() -> tuple[Optional[Path], pd.DataFrame]:
    root = tk.Tk()
    root.withdraw()

    choice = messagebox.askyesnocancel(
        "Ведомость — старт",
        "Открыть существующий файл ведомости?\n\n"
        "Да  — открыть файл\n"
        "Нет — создать новую\n"
        "Отмена — выход",
        parent=root,
    )

    if choice is None:
        root.destroy()
        raise SystemExit(0)

    if choice is True:
        file_path = filedialog.askopenfilename(
            parent=root,
            title="Открыть ведомость",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not file_path:
            root.destroy()
            raise SystemExit(0)

        path = Path(file_path)
        try:
            df = read_excel_any(path)
        except Exception as e:
            messagebox.showerror("Ошибка чтения", f"Не удалось прочитать файл:\n{path}\n\n{e}", parent=root)
            root.destroy()
            raise SystemExit(2)

        df = align_df_to_columns(df, COLUMNS)
        root.destroy()
        return path, df

    # Create new
    df = pd.DataFrame(columns=COLUMNS)
    df.loc[0, COL_TIP] = "Группа"
    df.loc[0, COL_NAME] = "Работы ресурсной ведомости"
    root.destroy()
    return None, df


def main() -> int:
    current_path, df = start_flow()

    dropdowns = DROPDOWNS
    defaults = read_defaults_from_fill_logic(FILL_LOGIC_FILE)

    model = VedomostModel(df, columns=COLUMNS)

    # resource db
    rdb = ResourceDB(RESOURCES_DB_FILE)
    rdb.load()

    rules = load_method_rules(FORMULAS_FILE)
    calc = CalcEngine(rules)

    app = OneCLikeApp(model=model, dropdowns=dropdowns, current_path=current_path,
                      resource_db=rdb, defaults=defaults, calc=calc)
    app.mainloop()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        # Показываем понятную ошибку даже при запуске "двойным кликом"
        import traceback
        err = traceback.format_exc()
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror("Ошибка приложения", err)
            r.destroy()
        except Exception:
            print(err)
        raise