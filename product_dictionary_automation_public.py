# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from rapidfuzz import fuzz, process
except ImportError:
    raise SystemExit("Install rapidfuzz first: pip install rapidfuzz")

# openpyxl فقط برای تنظیم recalculation در فایل خروجی لازم است
# (pandas نمی‌تواند این ویژگی را روی ورک‌بوک ست کند)
from openpyxl import load_workbook

HIGH_SCORE = 97.0
GOOD_SCORE = 90.0
REVIEW_SCORE = 80.0
HIGH_MARGIN = 8.0
GOOD_MARGIN = 5.0

DIGITS_TRANS = str.maketrans(
    "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
    "01234567890123456789",
)

REPLACEMENTS = {
    # Portfolio version: company-specific product aliases removed.
    # Add project-specific normalization rules locally if needed.
}


def normalize_text(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""

    s = str(value).translate(DIGITS_TRANS)
    s = (
        s.replace("\u200c", " ")
        .replace("\u200f", " ")
        .replace("\u200e", " ")
        .replace("ي", "ی")
        .replace("ى", "ی")
        .replace("ك", "ک")
        .replace("ۀ", "ه")
        .replace("ة", "ه")
        .lower()
    )

    for old, new in REPLACEMENTS.items():
        s = s.replace(old, new)

    s = re.sub(r"(?<=\d)\s*ع\b", " عددی", s)
    s = re.sub(
        r"(\d+)\s*/\s*5\s*(لیتر|ليتر)",
        lambda m: str(int(m.group(1)) * 1000 + 500) + " میلی لیتر",
        s,
    )
    s = re.sub(r"(\d+)\s*/\s*(\d+)", r"\1.\2", s)
    s = re.sub(r"[^0-9a-zآ-ی]+", " ", s)

    return re.sub(r"\s+", " ", s).strip()


def confidence_status(score: float, margin: float) -> str:
    if score >= HIGH_SCORE and margin >= HIGH_MARGIN:
        return "خیلی مطمئن"
    if score >= GOOD_SCORE and margin >= GOOD_MARGIN:
        return "مطمئن"
    if score >= REVIEW_SCORE:
        return "بررسی شود"
    return "بررسی دقیق"


def clean_code(value) -> str:
    """کد را به رشته تبدیل می‌کند بدون افزودن .0 که هنگام خواندن اعداد اکسل توسط pandas پیش می‌آید."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def update_dictionary(dict_path: str, master_path: str, output_path: str):
    df_dict = pd.read_excel(dict_path, sheet_name=0, dtype=object)
    df_master = pd.read_excel(master_path, sheet_name=0, dtype=object)

    required_d = ["نام کالای نماینده", "کد کالای خودمان", "نام کالای خودمان"]
    required_m = ["کد", "عنوان"]

    for h in required_d:
        if h not in df_dict.columns:
            raise ValueError(f"Missing Dictionary column: {h}")
    for h in required_m:
        if h not in df_master.columns:
            raise ValueError(f"Missing master column: {h}")

    # --- ساخت دیکشنری master: کد -> عنوان ---
    df_master = df_master.copy()
    df_master["کد"] = df_master["کد"].map(clean_code)
    df_master["عنوان"] = df_master["عنوان"].astype(str).str.strip()
    df_master = df_master[(df_master["کد"] != "") & (df_master["عنوان"] != "")]

    master_by_code: dict[str, str] = dict(zip(df_master["کد"], df_master["عنوان"]))

    if not master_by_code:
        raise ValueError("No products found in master file.")

    # --- ساخت exemplars: برای هر کد، لیست نام‌های نرمال‌شده (نمونه‌های موجود + عنوان master) ---
    exemplars: dict[str, list[str]] = {code: [] for code in master_by_code}

    dict_codes_clean = df_dict["کد کالای خودمان"].map(clean_code)
    mask_existing = (
        (dict_codes_clean != "")
        & dict_codes_clean.isin(master_by_code.keys())
        & df_dict["نام کالای نماینده"].notna()
    )
    existing_count = int(mask_existing.sum())

    for code, rep_name in zip(
        dict_codes_clean[mask_existing], df_dict.loc[mask_existing, "نام کالای نماینده"]
    ):
        n = normalize_text(rep_name)
        if n:
            exemplars[code].append(n)

    for code, title in master_by_code.items():
        n = normalize_text(title)
        if n:
            exemplars[code].append(n)

    # --- تخت کردن exemplars برای امتیازدهی برداری با rapidfuzz.process.cdist ---
    flat_codes = []
    flat_examples = []
    for code, exs in exemplars.items():
        for ex in exs:
            flat_codes.append(code)
            flat_examples.append(ex)

    flat_codes_arr = np.array(flat_codes)
    flat_examples_arr = np.array(flat_examples, dtype=object)

    n_rows = len(df_dict)
    out_code = np.empty(n_rows, dtype=object)
    out_name = np.empty(n_rows, dtype=object)
    out_score = np.empty(n_rows, dtype=object)
    out_margin = np.empty(n_rows, dtype=object)
    out_status = np.empty(n_rows, dtype=object)

    new_codes = dict_codes_clean.to_numpy(dtype=object, copy=True)
    new_names = df_dict["نام کالای خودمان"].to_numpy(dtype=object, copy=True)

    filled_count = high_count = good_count = review_count = 0
    detailed_review_count = blank_count = 0

    for i in range(n_rows):
        current_code = dict_codes_clean.iat[i]
        rep_name = df_dict["نام کالای نماینده"].iat[i]

        # نگاشت موجود، معتبر می‌ماند
        if current_code != "":
            title = master_by_code.get(current_code, df_dict["نام کالای خودمان"].iat[i] or "")
            out_code[i] = current_code
            out_name[i] = title
            out_score[i] = 100
            out_margin[i] = ""
            out_status[i] = "موجود در دیکشنری"
            continue

        nt = normalize_text(rep_name)

        if not nt:
            out_code[i] = ""
            out_name[i] = ""
            out_score[i] = 0
            out_margin[i] = 0
            out_status[i] = "نام کالا خالی است"
            blank_count += 1
            continue

        if flat_examples_arr.size == 0:
            out_status[i] = "بدون گزینه تطبیق"
            out_code[i] = ""
            out_name[i] = ""
            out_score[i] = 0
            out_margin[i] = 0
            detailed_review_count += 1
            continue

        # امتیازدهی برداری این ردیف در برابر تمام نمونه‌ها با rapidfuzz.cdist
        scores = process.cdist([nt], flat_examples_arr, scorer=fuzz.WRatio)[0]

        # بهترین امتیاز هر کد با groupby روی pandas Series
        best_per_code = (
            pd.Series(scores, index=flat_codes_arr).groupby(level=0).max().sort_values(ascending=False)
        )

        c1 = best_per_code.index[0]
        s1 = float(best_per_code.iloc[0])
        s2 = float(best_per_code.iloc[1]) if len(best_per_code) > 1 else 0.0
        margin = s1 - s2
        status = confidence_status(s1, margin)
        title = master_by_code[c1]

        out_code[i] = c1
        out_name[i] = title
        out_score[i] = round(s1, 2)
        out_margin[i] = round(margin, 2)
        out_status[i] = status

        new_codes[i] = c1
        new_names[i] = title

        filled_count += 1
        if status == "خیلی مطمئن":
            high_count += 1
        elif status == "مطمئن":
            good_count += 1
        elif status == "بررسی شود":
            review_count += 1
        else:
            detailed_review_count += 1

    # --- نوشتن ستون‌های نهایی روی دیتافریم ---
    df_dict["کد کالای خودمان"] = new_codes
    df_dict["نام کالای خودمان"] = new_names
    df_dict["پیشنهاد کد خودمان"] = out_code
    df_dict["پیشنهاد نام خودمان"] = out_name
    df_dict["امتیاز تطبیق"] = out_score
    df_dict["فاصله از گزینه دوم"] = out_margin
    df_dict["وضعیت تطبیق"] = out_status

    df_dict.to_excel(output_path, index=False, engine="openpyxl")

    # اجبار به recalculation کامل هنگام باز شدن فایل در اکسل
    wb = load_workbook(output_path)
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(output_path)
    wb.close()

    print()
    print("SUCCESS")
    print(f"Output: {output_path}")
    print(f"Existing mappings kept: {existing_count}")
    print(f"New rows filled: {filled_count}")
    print(f"Very confident: {high_count}")
    print(f"Confident: {good_count}")
    print(f"Review: {review_count}")
    print(f"Detailed review: {detailed_review_count}")
    print(f"Blank names: {blank_count}")


if __name__ == "__main__":
    base = Path(__file__).resolve().parent

    dict_path = base / "Dictionary_v2_matched.xlsx"
    master_path = base / "نام کالا کد کالا.xlsx"
    output_path = base / "Dictionary_v3_matched_updated.xlsx"

    if not dict_path.exists():
        raise SystemExit(
            f"Dictionary file not found:\n{dict_path}\n"
            "Expected: Dictionary_v2_matched.xlsx"
        )

    if not master_path.exists():
        raise SystemExit(
            f"Master file not found:\n{master_path}\n"
            'Expected: نام کالا کد کالا.xlsx'
        )

    update_dictionary(
        str(dict_path),
        str(master_path),
        str(output_path),
    )
