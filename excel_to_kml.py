
# -*- coding: utf-8 -*-
"""
Excel → KML converter (RU headers) — v3
----------------------------------------
Требования: Python 3.9+, pandas, openpyxl

Новые дефолты:
- Если не указаны --excel/--in-dir, берём папку **reports** рядом со скриптом.
- В пакетном режиме по умолчанию сохраняем KML в **текущую папку** (корень запуска).
- Имена KML = имена Excel-файлов (без расширения) + ".kml".

Примеры:
  # Без параметров: возьмёт .\reports и положит KML в текущую папку
  python excel_to_kml.py

  # Явно папка
  python excel_to_kml.py --in-dir "C:/Data/reports"

  # Один файл (как раньше)
  python excel_to_kml.py --excel "input.xlsx" --sheet "Лист1" --out "gorod.kml"
"""
import argparse
import hashlib
import html
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pandas as pd
import xml.etree.ElementTree as ET


def ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def app_base() -> Path:
    # Если запущено из .exe (PyInstaller)
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # Обычный запуск .py
    return Path(__file__).resolve().parent

def ensure_float(s):
    """Конвертирует число с запятой/точкой в float. Возвращает None при ошибке."""
    if pd.isna(s):
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip().replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def in_range_lat_lon(lat: float, lon: float) -> bool:
    return lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180


def kml_color_from_district(district: str) -> str:
    """
    Цвет KML (aabbggrr) детерминированно из названия района.
    Альфа=FF, HSV -> RGB (s=0.7, v=0.95).
    """
    h = int(hashlib.md5(district.encode("utf-8")).hexdigest(), 16) % 360  # hue
    s, v = 0.7, 0.95
    c = v * s
    x = c * (1 - abs((h / 60) % 2 - 1))
    m = v - c
    if 0 <= h < 60:
        r, g, b = (c, x, 0)
    elif 60 <= h < 120:
        r, g, b = (x, c, 0)
    elif 120 <= h < 180:
        r, g, b = (0, c, x)
    elif 180 <= h < 240:
        r, g, b = (0, x, c)
    elif 240 <= h < 300:
        r, g, b = (x, 0, c)
    else:
        r, g, b = (c, 0, x)
    R = int((r + m) * 255)
    G = int((g + m) * 255)
    B = int((b + m) * 255)
    return f"FF{B:02X}{G:02X}{R:02X}"  # aabbggrr


def create_style(style_id: str, color: str) -> ET.Element:
    style = ET.Element("Style", id=style_id)
    icon_style = ET.SubElement(style, "IconStyle")
    ET.SubElement(icon_style, "color").text = color
    ET.SubElement(icon_style, "scale").text = "1.2"
    icon = ET.SubElement(icon_style, "Icon")
    ET.SubElement(icon, "href").text = "http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png"
    label_style = ET.SubElement(style, "LabelStyle")
    ET.SubElement(label_style, "scale").text = "0.9"
    return style


def build_description(addr: str, obj: str, topic: str, text: str, links: List[str]) -> str:
    parts = []
    if addr:
        parts.append(f"<b>Адрес:</b> {html.escape(str(addr))}")
    if obj:
        parts.append(f"<b>Объект:</b> {html.escape(str(obj))}")
    if topic:
        parts.append(f"<b>Проблема:</b> {html.escape(str(topic))}")
    if text:
        parts.append(f"<b>Текст:</b> {html.escape(str(text))}")
    if links:
        a = []
        for i, u in enumerate(links, 1):
            u = u.strip()
            if not u:
                continue
            a.append(f'<a href="{html.escape(u)}" target="_blank">Фото {i}</a>')
        if a:
            parts.append(f"<b>Фото:</b> " + " | ".join(a))
    html_block = "<br/>".join(parts)
    return f"<![CDATA[{html_block}]]>"


def excel_to_kml(excel_path: Path, sheet_name: str, out_path: Path) -> Tuple[int, int, List[str]]:
    """
    Конвертирует один Excel в KML.
    Возвращает: (total_rows, written_count, problems_log).
    """
    # .xls не поддерживаем (xlrd>=2.0), .xlsx и .xlsm — ок.
    if excel_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise SystemExit(f"{excel_path.name}: поддерживаются только .xlsx/.xlsm (сконвертируйте .xls в .xlsx)")

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=2, dtype=str)
    required_cols = [
        "Номер сообщения",
        "Округ",
        "Район",
        "Адрес",
        "Название объекта",
        "Проблемная тема",
        "Текст сообщения",
        "Ссылки на фотографии сообщения",
        "Широта",
        "Долгота",
    ]
    for col in required_cols:
        if col not in df.columns:
            raise SystemExit(f"{excel_path.name}: отсутствует обязательная колонка: {col}")

    # Normalize and validate
    df["Широта_num"] = df["Широта"].map(ensure_float)
    df["Долгота_num"] = df["Долгота"].map(ensure_float)

    problems: List[str] = []

    def valid_row(row) -> bool:
        if pd.isna(row["Номер сообщения"]) or str(row["Номер сообщения"]).strip() == "":
            problems.append(f"{excel_path.name}: пустой 'Номер сообщения' в строке Excel {row.name + 1}")
            return False
        lat, lon = row["Широта_num"], row["Долгота_num"]
        if not in_range_lat_lon(lat, lon):
            problems.append(
                f"{excel_path.name}: неверные координаты в строке Excel {row.name + 1}: lat={row['Широта']}, lon={row['Долгота']}"
            )
            return False
        return True

    df_valid = df[df.apply(valid_row, axis=1)].copy()

    # KML root
    kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
    doc = ET.SubElement(kml, "Document")
    ET.SubElement(doc, "name").text = out_path.name

    # Styles per district
    districts = sorted(df_valid["Район"].dropna().unique())
    style_ids: Dict[str, str] = {}
    for d in districts:
        sid = f"style_{hashlib.md5(d.encode('utf-8')).hexdigest()[:8]}"
        style_ids[d] = sid
        doc.append(create_style(sid, kml_color_from_district(d)))

    # Folders: Okrug -> Rayon
    folder_map: Dict[Tuple[str, str], ET.Element] = {}
    okrug_folders: Dict[str, ET.Element] = {}

    def get_okrug_folder(okrug: str) -> ET.Element:
        if okrug not in okrug_folders:
            f = ET.SubElement(doc, "Folder")
            ET.SubElement(f, "name").text = okrug
            okrug_folders[okrug] = f
        return okrug_folders[okrug]

    def get_pair_folder(okrug: str, rayon: str) -> ET.Element:
        key = (okrug, rayon)
        if key not in folder_map:
            of = get_okrug_folder(okrug)
            rf = ET.SubElement(of, "Folder")
            ET.SubElement(rf, "name").text = rayon
            folder_map[key] = rf
        return folder_map[key]

    written = 0
    for _, row in df_valid.iterrows():
        okrug = str(row["Округ"]).strip()
        rayon = str(row["Район"]).strip()
        folder = get_pair_folder(okrug, rayon)

        name = str(row["Номер сообщения"]).strip()
        addr = row["Адрес"]
        obj = row["Название объекта"]
        topic = row["Проблемная тема"]
        text_msg = row["Текст сообщения"]
        links_raw = row["Ссылки на фотографии сообщения"] or ""
        links = [u.strip() for u in str(links_raw).split(";") if u.strip()]

        lat = float(row["Широта_num"])
        lon = float(row["Долгота_num"])

        pm = ET.SubElement(folder, "Placemark")
        ET.SubElement(pm, "name").text = name
        sid = style_ids.get(rayon)
        if sid:
            ET.SubElement(pm, "styleUrl").text = f"#{sid}"
        desc = ET.SubElement(pm, "description")
        desc.text = build_description(addr, obj, topic, text_msg, links)
        point = ET.SubElement(pm, "Point")
        ET.SubElement(point, "coordinates").text = f"{lon:.8f},{lat:.8f},0"
        written += 1

    # Write file
    xml_bytes = ET.tostring(kml, encoding="utf-8", method="xml")
    text = xml_bytes.decode("utf-8").replace("><", ">\n<")
    Path(out_path).write_text(text, encoding="utf-8")

    return len(df), written, problems


def derive_out_name(xlsx_path: Path) -> str:
    return f"{xlsx_path.stem}.kml"


def collect_excel_files(in_dir: Path) -> List[Path]:
    files: List[Path] = []
    # Поддерживаем .xlsx и .xlsm; временные ~$.xlsx пропускаем
    files += [p for p in in_dir.glob("*.xlsx") if not p.name.startswith("~$")]
    files += [p for p in in_dir.glob("*.xlsm") if not p.name.startswith("~$")]
    return sorted(files)


def process_dir(in_dir: Path, out_dir: Optional[Path], sheet_name: str) -> int:
    """
    Обрабатывает все .xlsx/.xlsm в папке.
    Возвращает количество успешно созданных KML.
    """
    if not in_dir.exists():
        print(f"[ERR] Папка не найдена: {in_dir}", file=sys.stderr)
        return 0
    if out_dir is None:
        out_dir = Path.cwd()  # по умолчанию — корень запуска
    out_dir.mkdir(parents=True, exist_ok=True)

    excel_files = collect_excel_files(in_dir)
    if not excel_files:
        print(f"[WARN] В папке нет .xlsx/.xlsm файлов: {in_dir}")
        return 0

    ok_count = 0
    for x in excel_files:
        out_name = derive_out_name(x)
        out_path = out_dir / out_name
        try:
            total, written, problems = excel_to_kml(x, sheet_name, out_path)
            print(f"[OK] {x.name} → {out_path.name} (всего: {total}, записано: {written}, пропущено: {len(problems)})")
            if problems:
                for p in problems:
                    print(" - " + p)
            ok_count += 1
        except SystemExit as e:
            print(f"[ERR] {x.name}: {e}", file=sys.stderr)
        except Exception as e:
            print(f"[ERR] {x.name}: {e}", file=sys.stderr)
    return ok_count


def main():
    parser = argparse.ArgumentParser(description="Excel (.xlsx/.xlsm) → KML generator (single file or folder)")
    g = parser.add_mutually_exclusive_group(required=False)
    g.add_argument("--excel", help="Путь к одному Excel-файлу (.xlsx/.xlsm)")
    g.add_argument("--in-dir", help="Папка с Excel-файлами (.xlsx/.xlsm)")
    parser.add_argument("--sheet", default="Лист1", help="Имя листа (по умолчанию 'Лист1')")

    parser.add_argument("--out", help="Имя выходного KML-файла (только для --excel). Если не задан, gorod_YYYYMMDD_HHMM.kml")
    parser.add_argument("--out-dir", help="Папка для вывода (только для --in-dir). Если не задана, используется текущая папка.")

    args = parser.parse_args()

    if args.excel:
        excel = Path(args.excel)
        if not excel.exists():
            print(f"[ERR] Файл не найден: {excel}", file=sys.stderr)
            sys.exit(2)
        out = Path(args.out) if args.out else Path(f"gorod_{ts()}.kml")
        total, written, problems = excel_to_kml(excel, args.sheet, out)
        print(f"[OK] Готово: {out}  (всего строк: {total}, записано: {written}, пропущено: {len(problems)})")
        if problems:
            print("---- Пропуски / проблемы ----")
            for p in problems:
                print(" - " + p)
    else:
        # Дефолт: папка reports рядом со скриптом
        in_dir = Path(args.in_dir) if args.in_dir else (app_base() / "reports")
        out_dir = Path(args.out_dir) if args.out_dir else None  # None => текущая папка
        count = process_dir(in_dir, out_dir, args.sheet)
        print(f"[DONE] Успешно создано KML: {count} шт. (из {in_dir})")

if __name__ == "__main__":
    main()
