import base64
import csv
import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional, Tuple

try:
    import numpy as np
    from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageOps
    import treepoem
    import zxingcpp
    from openpyxl import load_workbook
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
except ImportError as exc:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Не хватает библиотек",
        "Запустите install_and_run.bat.\n\n"
        f"Ошибка импорта: {exc}",
    )
    raise SystemExit(1)


# treepoem запускает консольный Ghostscript дважды для каждого символа.
# В оконной сборке Windows эти процессы должны работать без всплывающего CMD.
_subprocess_run = subprocess.run


def _run_subprocess_hidden(*args, **kwargs):
    if os.name == "nt":
        kwargs["creationflags"] = kwargs.get("creationflags", 0) | subprocess.CREATE_NO_WINDOW
        startupinfo = kwargs.get("startupinfo") or subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs["startupinfo"] = startupinfo
    return _subprocess_run(*args, **kwargs)


class _HiddenSubprocessProxy:
    run = staticmethod(_run_subprocess_hidden)
    # treepoem checks this attribute before passing the flag to Ghostscript.
    # Keep the constants visible through the proxy; otherwise its own
    # CREATE_NO_WINDOW protection is disabled and a black console flashes.
    CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    STARTF_USESHOWWINDOW = getattr(subprocess, "STARTF_USESHOWWINDOW", 1)
    SW_HIDE = getattr(subprocess, "SW_HIDE", 0)
    STARTUPINFO = getattr(subprocess, "STARTUPINFO", None)


treepoem.subprocess = _HiddenSubprocessProxy()


APP_TITLE = "PROJECTZ GS1 DataMatrix — Честный Знак"
APP_VERSION = "4.0.7.5 — пакетные этикетки 100 × 73 мм"
GS = "\x1d"

# ВАЖНО: структура зафиксирована и не должна изменяться:
# {FNC1}01GTIN21SERIAL{GS}91KEY{GS}92CRYPTO
DEFAULT_CODE = (
    "[01]04602322013476"
    "[21]5n2wdl(iVFd8("
    "[91]EE11"
    "[92]TnOUjY60KSfpW290LByTGoAtveyffIm2zEmPZlbBrqY="
)

ALLOWED_SERIAL_RE = re.compile(r'^[\x20-\x7E]+$')
SAFE_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1F]')

LABEL_PRESETS = {
    "100 × 73 мм": (100.0, 73.0),
    "58 × 30 мм": (58.0, 30.0),
    "58 × 40 мм": (58.0, 40.0),
    "43 × 25 мм": (43.0, 25.0),
    "30 × 20 мм": (30.0, 20.0),
    "Пользовательский": None,
}


@dataclass
class CodeParts:
    gtin: str
    serial: str
    key91: str
    crypto92: str


@dataclass
class BatchItem:
    row_no: int
    source_code: str
    caption: str = ""
    filename: str = ""
    status: str = ""
    error: str = ""
    parts: Optional[CodeParts] = None


def find_ghostscript() -> Optional[str]:
    # A PyInstaller one-file build extracts the bundled Ghostscript beside
    # the Python payload. Put that private bin directory on PATH so both this
    # check and treepoem resolve the bundled executable without a system-wide
    # Ghostscript installation.
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    bundled_bin = bundle_root / "ghostscript" / "bin"
    bundled_exe = bundled_bin / "gswin64c.exe"
    if bundled_exe.is_file():
        current_path = os.environ.get("PATH", "")
        bundled_bin_text = str(bundled_bin)
        if bundled_bin_text not in current_path.split(os.pathsep):
            os.environ["PATH"] = bundled_bin_text + os.pathsep + current_path
        return str(bundled_exe)

    for name in ("gswin64c.exe", "gswin32c.exe", "gs.exe"):
        found = shutil.which(name)
        if found:
            return found

    roots = [
        Path(r"C:\Program Files\gs"),
        Path(r"C:\Program Files (x86)\gs"),
    ]
    candidates = []
    for root in roots:
        if root.exists():
            candidates.extend(root.glob(r"gs*\bin\gswin64c.exe"))
            candidates.extend(root.glob(r"gs*\bin\gswin32c.exe"))

    if candidates:
        return str(sorted(candidates, reverse=True)[0])
    return None


def normalize_code(text: str) -> str:
    """
    Принимает несколько представлений одного и того же кода:

    1. [01]GTIN[21]SERIAL[91]KEY[92]CRYPTO
    2. {FNC1}01GTIN21SERIAL{GS}91KEY{GS}92CRYPTO
    3. <FNC1>01GTIN21SERIAL<GS>91KEY<GS>92CRYPTO
    4. Строку с настоящим ASCII GS (код 29)
    5. Вариант со скобками: (01)...(21)...(91)...(92)...
    6. Сырую строку сканера с пробелами перед AI 91 и AI 92:
       01...21СЕРИЙНЫЙ 91КЛЮЧ 92КРИПТОХВОСТ

    Внутреннее представление всегда возвращается в квадратных скобках.
    """
    value = "".join(text.strip().splitlines()).strip()

    # Убираем AIM-префикс сканера, если он был скопирован вместе с данными.
    if value.startswith("]d2"):
        value = value[3:]

    # Some scanner exports display the leading FNC1 as a single non-ASCII
    # character (for example U+00E8, ``è``). Accept that representation only
    # when it is immediately followed by a complete GS1 AI 01 + GTIN + AI 21.
    # Do not strip arbitrary ASCII characters: a damaged code must remain an
    # error instead of being silently altered.
    if re.match(r"^[^\x00-\x7f]01[0-9]{14}21", value):
        value = value[1:]

    # Нормализуем обозначения AI со скобками.
    for ai in ("01", "21", "91", "92"):
        value = value.replace(f"({ai})", f"[{ai}]")

    # Уже нормализованный пользовательский формат.
    if re.fullmatch(
        r"\[01\]\d{14}\[21\].*?\[91\].{4}\[92\].+",
        value,
        re.DOTALL,
    ):
        return value

    # Видимое представление служебных символов.
    value = re.sub(r"^(?:\{FNC1\}|<FNC1>)", "", value, flags=re.IGNORECASE)

    gs_marker = r"(?:\{GS\}|<GS>|\x1d)"
    match = re.fullmatch(
        rf"01(\d{{14}})"
        rf"21(.*?)"
        rf"{gs_marker}91(.{{4}})"
        rf"{gs_marker}92(.+)",
        value,
        re.DOTALL | re.IGNORECASE,
    )
    if match:
        gtin, serial, key91, crypto92 = match.groups()
        return (
            f"[01]{gtin}"
            f"[21]{serial}"
            f"[91]{key91}"
            f"[92]{crypto92}"
        )

    # Сырая строка сканера, где пробелы отделяют AI 91 и AI 92.
    # Пример:
    # 0104603721560172215GOE8bBs%?Exu 91EE12 92XxR7...
    raw_space_match = re.fullmatch(
        r"01(\d{14})"
        r"21(.+?)"
        r"\s+91(.{4})"
        r"\s+92(.+)",
        value,
        re.DOTALL,
    )
    if raw_space_match:
        gtin, serial, key91, crypto92 = raw_space_match.groups()
        return (
            f"[01]{gtin}"
            f"[21]{serial.rstrip()}"
            f"[91]{key91}"
            f"[92]{crypto92.strip()}"
        )

    # Некоторые сканеры удаляют разделители полностью:
    # 01GTIN21SERIAL91KEY92CRYPTO. AI 91 и AI 92 имеют фиксированную
    # структуру, поэтому такой текст можно восстановить. Это совместимость
    # с выводом сканера; при генерации программа всё равно добавляет GS.
    compact_match = re.fullmatch(
        r"01(\d{14})21(.+?)91(.{4})92(.+)",
        value,
        re.DOTALL,
    )
    if compact_match:
        gtin, serial, key91, crypto92 = compact_match.groups()
        return (
            f"[01]{gtin}"
            f"[21]{serial}"
            f"[91]{key91}"
            f"[92]{crypto92}"
        )

    # Поддержка строки без видимых GS, но с явными AI в квадратных скобках.
    bracketed = value
    bracketed = bracketed.replace("{GS}", "").replace("<GS>", "").replace(GS, "")
    if re.fullmatch(
        r"\[01\]\d{14}\[21\].*?\[91\].{4}\[92\].+",
        bracketed,
        re.DOTALL | re.IGNORECASE,
    ):
        return bracketed

    return value


def parse_code(code: str) -> CodeParts:
    match = re.fullmatch(
        r"\[01\](\d{14})"
        r"\[21\](.*?)"
        r"\[91\](.{4})"
        r"\[92\](.+)",
        code,
        re.DOTALL,
    )
    if not match:
        raise ValueError(
            "Ожидаемый формат: "
            "[01]14 цифр GTIN[21]серийный номер[91]4 символа[92]криптохвост"
        )
    return CodeParts(*match.groups())


def gtin_check_digit(gtin_without_check: str) -> int:
    total = 0
    for index, char in enumerate(reversed(gtin_without_check)):
        total += int(char) * (3 if index % 2 == 0 else 1)
    return (10 - total % 10) % 10


def validate_gtin(gtin: str) -> Tuple[bool, str]:
    if len(gtin) != 14 or not gtin.isdigit():
        return False, "GTIN должен содержать 14 цифр."
    expected = gtin_check_digit(gtin[:-1])
    actual = int(gtin[-1])
    if expected != actual:
        return False, (
            f"Неверная контрольная цифра GTIN: указана {actual}, "
            f"ожидается {expected}."
        )
    return True, "GTIN корректен."


def validate_base64_loose(value: str) -> Tuple[bool, str]:
    if not value:
        return False, "Поле AI 92 не может быть пустым."
    if not re.fullmatch(r"[A-Za-z0-9+/=_-]+", value):
        return False, "AI 92 содержит символы вне Base64/Base64URL."
    try:
        padded = value + "=" * ((4 - len(value) % 4) % 4)
        base64.b64decode(padded.replace("-", "+").replace("_", "/"), validate=False)
    except Exception:
        return False, "AI 92 не удалось интерпретировать как Base64."
    return True, "AI 92 выглядит корректно."


def validate_code(code: str, strict_gtin: bool = True) -> Tuple[bool, List[str], Optional[CodeParts]]:
    messages: List[str] = []
    try:
        parts = parse_code(code)
    except ValueError as exc:
        return False, [str(exc)], None

    if not parts.serial:
        messages.append("AI 21 не может быть пустым.")
    elif len(parts.serial) > 20:
        messages.append(f"AI 21 длиннее 20 символов: {len(parts.serial)}.")
    elif not ALLOWED_SERIAL_RE.fullmatch(parts.serial):
        messages.append("AI 21 содержит управляющие или не-ASCII символы.")

    if len(parts.key91) != 4:
        messages.append("AI 91 должен содержать ровно 4 символа.")
    elif not ALLOWED_SERIAL_RE.fullmatch(parts.key91):
        messages.append("AI 91 содержит недопустимые символы.")

    ok92, msg92 = validate_base64_loose(parts.crypto92)
    if not ok92:
        messages.append(msg92)

    gtin_ok, gtin_msg = validate_gtin(parts.gtin)
    if strict_gtin and not gtin_ok:
        messages.append(gtin_msg)

    if any(ord(ch) < 32 or ord(ch) > 126 for ch in parts.serial + parts.key91 + parts.crypto92):
        messages.append("В переменных полях есть управляющие или не-ASCII символы.")

    return not messages, messages or ["Структура и поля корректны."], parts


def build_expected_payload(code: str) -> str:
    parts = parse_code(code)
    return (
        f"01{parts.gtin}"
        f"21{parts.serial}{GS}"
        f"91{parts.key91}{GS}"
        f"92{parts.crypto92}"
    )


def visible_payload(payload: str) -> str:
    return "{FNC1}" + payload.replace(GS, "{GS}")


def bwipp_escape(value: str) -> str:
    # BWIPP parse-режим: управляющие конструкции начинаются с ^.
    # Скобки и ^ внутри пользовательских значений кодируем по ASCII.
    result = []
    for char in value:
        if char in "()^":
            result.append(f"^{ord(char):03d}")
        else:
            result.append(char)
    return "".join(result)


def build_bwipp_data(code: str) -> str:
    """
    СТРУКТУРА ЗАФИКСИРОВАНА:
      специальный FNC1 в начале;
      ASCII GS (29) после AI 21;
      ASCII GS (29) после AI 91.

    Не заменять GS на FNC1.
    """
    parts = parse_code(code)
    return (
        "^FNC1"
        f"01{parts.gtin}"
        f"21{bwipp_escape(parts.serial)}"
        "^029"
        f"91{bwipp_escape(parts.key91)}"
        "^029"
        f"92{bwipp_escape(parts.crypto92)}"
    )


def generate_symbol(code: str, scale: int = 4) -> Image.Image:
    if not find_ghostscript():
        raise RuntimeError(
            "Ghostscript не найден. Запустите install_and_run.bat "
            "или установите Ghostscript."
        )

    image = treepoem.generate_barcode(
        barcode_type="datamatrix",
        data=build_bwipp_data(code),
        options={
            "parse": True,
            "parsefnc": True,
            "format": "square",
        },
        scale=max(2, scale),
    ).convert("L")

    # Только чёрный и белый, без антиалиасинга.
    image = image.point(lambda p: 0 if p < 128 else 255, mode="1").convert("L")
    border = max(4, scale * 2)
    return ImageOps.expand(image, border=border, fill=255)


def decode_and_verify(symbol: Image.Image, expected_payload: str):
    array = np.asarray(symbol.convert("L"))
    decode_options = {
        "formats": zxingcpp.BarcodeFormat.DataMatrix,
        "text_mode": zxingcpp.TextMode.Plain,
        "try_rotate": True,
        "try_downscale": True,
    }
    try:
        decoded = zxingcpp.read_barcode(array, try_invert=True, **decode_options)
    except TypeError:
        # zxing-cpp 2.2 is the last practical Python 3.8/Windows 7 wheel and
        # predates the try_invert keyword used by the current release.
        decoded = zxingcpp.read_barcode(array, **decode_options)
    if not decoded.valid:
        raise RuntimeError("Контрольное считывание DataMatrix не выполнено.")

    raw = getattr(decoded, "bytes", None)
    actual = (
        bytes(raw).decode("latin-1")
        if isinstance(raw, (bytes, bytearray))
        else decoded.text
    )

    if decoded.symbology_identifier != "]d2":
        raise RuntimeError(
            "Символ прочитан не как GS1 DataMatrix. "
            f"AIM: {decoded.symbology_identifier}"
        )
    if actual != expected_payload:
        raise RuntimeError(
            "Считанные данные не совпали.\n"
            f"Ожидалось: {visible_payload(expected_payload)}\n"
            f"Прочитано: {visible_payload(actual)}"
        )
    return decoded


def load_font(size: int, bold: bool = False):
    names = (
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibrib.ttf" if bold else r"C:\Windows\Fonts\calibri.ttf",
        r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf",
    )
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            pass
    return ImageFont.load_default()


def compose_label(
    symbol: Image.Image,
    width_mm: float,
    height_mm: float,
    dpi: int,
    caption: str,
    show_caption: bool,
) -> Image.Image:
    """
    Размещает DataMatrix и подпись в отдельных областях.

    Подпись никогда не накладывается на код:
    для неё заранее резервируется высота с дополнительным зазором.
    """
    width_px = max(1, round(width_mm / 25.4 * dpi))
    height_px = max(1, round(height_mm / 25.4 * dpi))
    label = Image.new("L", (width_px, height_px), 255)

    margin = max(6, round(1.5 / 25.4 * dpi))
    gap = max(6, round(1.2 / 25.4 * dpi))

    font = None
    text_bbox = (0, 0, 0, 0)
    caption_block_h = 0

    if show_caption and caption:
        # Подбираем размер шрифта, чтобы длинная подпись помещалась по ширине.
        font_size = max(10, round(2.5 / 25.4 * dpi))
        probe = ImageDraw.Draw(label)

        while font_size >= 8:
            font = load_font(font_size)
            text_bbox = probe.textbbox((0, 0), caption, font=font)
            text_w = text_bbox[2] - text_bbox[0]
            if text_w <= width_px - 2 * margin:
                break
            font_size -= 1

        text_h = text_bbox[3] - text_bbox[1]
        caption_block_h = text_h + gap

    # Область DataMatrix полностью отделена от области подписи.
    symbol_area_left = margin
    symbol_area_top = margin
    symbol_area_right = width_px - margin
    symbol_area_bottom = height_px - margin - caption_block_h

    available_w = symbol_area_right - symbol_area_left
    available_h = symbol_area_bottom - symbol_area_top

    if available_w <= 0 or available_h <= 0:
        raise ValueError("Размер этикетки слишком мал для DataMatrix и подписи.")

    # Масштабируем только целым коэффициентом, сохраняя чёткие модули.
    factor = max(1, min(available_w // symbol.width, available_h // symbol.height))
    rendered = symbol.resize(
        (symbol.width * factor, symbol.height * factor),
        Image.Resampling.NEAREST,
    )

    x = symbol_area_left + (available_w - rendered.width) // 2
    y = symbol_area_top + (available_h - rendered.height) // 2
    label.paste(rendered, (x, y))

    if show_caption and caption and font:
        draw = ImageDraw.Draw(label)
        text_bbox = draw.textbbox((0, 0), caption, font=font)
        text_w = text_bbox[2] - text_bbox[0]
        text_h = text_bbox[3] - text_bbox[1]

        text_x = max(margin, (width_px - text_w) // 2)
        # Верх подписи начинается строго ниже области DataMatrix.
        text_y = height_px - margin - text_h - text_bbox[1]
        draw.text((text_x, text_y), caption, font=font, fill=0)

    return label.convert("RGB")


def image_to_zpl(label: Image.Image, width_mm: float, height_mm: float, dpi: int) -> str:
    mono = label.convert("1")
    width, height = mono.size
    bytes_per_row = math.ceil(width / 8)
    rows = []
    pixels = mono.load()

    for y in range(height):
        row = bytearray()
        for byte_index in range(bytes_per_row):
            value = 0
            for bit in range(8):
                x = byte_index * 8 + bit
                black = x < width and pixels[x, y] == 0
                if black:
                    value |= 1 << (7 - bit)
            row.append(value)
        rows.append(row.hex().upper())

    total_bytes = bytes_per_row * height
    data = "".join(rows)
    return (
        "^XA\n"
        "^PW{width}\n"
        "^LL{height}\n"
        "^FO0,0\n"
        f"^GFA,{total_bytes},{total_bytes},{bytes_per_row},{data}\n"
        "^FS\n"
        "^XZ\n"
    ).format(width=width, height=height)


def safe_filename(value: str, fallback: str = "datamatrix") -> str:
    cleaned = SAFE_FILENAME_RE.sub("_", value).strip(" ._")
    return (cleaned or fallback)[:120]


def read_batch_file(path: str) -> List[BatchItem]:
    source = Path(path)
    rows: List[Dict] = []

    if source.suffix.lower() in (".xlsx", ".xlsm"):
        wb = load_workbook(source, read_only=True, data_only=True)
        ws = wb.active
        values = list(ws.iter_rows(values_only=True))
        if not values:
            return []
        headers = [str(v or "").strip().lower() for v in values[0]]
        for values_row in values[1:]:
            rows.append({
                headers[i] if i < len(headers) else f"col{i}": values_row[i]
                for i in range(len(values_row))
            })
    elif source.suffix.lower() == ".csv":
        with source.open("r", encoding="utf-8-sig", newline="") as fh:
            sample = fh.read(4096)
            fh.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
            except csv.Error:
                dialect = csv.excel
            rows = list(csv.DictReader(fh, dialect=dialect))
    else:
        raise ValueError("Поддерживаются XLSX и CSV.")

    aliases = {
        "code": ("код маркировки", "код", "datamatrix", "data matrix", "code"),
        "caption": ("подпись", "наименование", "caption", "name"),
        "filename": ("имя файла", "файл", "filename", "file"),
    }

    def pick(row: dict, names: Tuple[str, ...]):
        normalized = {str(k).strip().lower(): v for k, v in row.items()}
        for name in names:
            if name in normalized and normalized[name] is not None:
                return str(normalized[name]).strip()
        return ""

    items = []
    for index, row in enumerate(rows, start=2):
        code = pick(row, aliases["code"])

        # Упрощённый режим: если в файле только одна колонка или
        # заголовок отличается, берём первое непустое значение как код.
        if not code:
            for value in row.values():
                candidate = str(value or "").strip()
                if candidate:
                    code = candidate
                    break

        if not code and not any(str(v or "").strip() for v in row.values()):
            continue

        normalized = normalize_code(code)
        caption = pick(row, aliases["caption"])
        filename = pick(row, aliases["filename"])

        # Автозаполнение из самого кода:
        # подпись = AI 21;
        # имя файла = GTIN + AI 21.
        try:
            parts = parse_code(normalized)
            if not caption:
                caption = parts.serial
            if not filename:
                filename = f"{parts.gtin}_{parts.serial}"
        except ValueError:
            pass

        items.append(BatchItem(
            row_no=index,
            source_code=normalized,
            caption=caption,
            filename=filename,
        ))
    return items


def create_template_xlsx(path: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Коды"
    headers = ["Код маркировки"]
    ws.append(headers)
    ws.append([DEFAULT_CODE])

    fill = PatternFill("solid", fgColor="5B21B6")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 115
    thin = Side(style="thin", color="D9D9D9")
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=1):
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1180x820")
        self.minsize(980, 700)
        self._restore_window_settings()
        self.protocol("WM_DELETE_WINDOW", self._close_window)

        self.current_code = ""
        self.current_payload = ""
        self.current_symbol: Optional[Image.Image] = None
        self.current_label: Optional[Image.Image] = None
        self.preview_photo = None
        self._preview_resize_job = None
        self.batch_items: List[BatchItem] = []
        self.batch_source = ""

        self._setup_style()
        self._build_ui()

    def _setup_style(self):
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 17, "bold"))
        style.configure("Section.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Success.TLabel", foreground="#12833B")
        style.configure("Error.TLabel", foreground="#B42318")

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="PROJECTZ GS1 DataMatrix", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            root,
            text="Структура Честного Знака зафиксирована: {FNC1}01…21…{GS}91…{GS}92…",
        ).pack(anchor="w", pady=(0, 10))

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)

        self.single_tab = ttk.Frame(notebook, padding=10)
        self.batch_tab = ttk.Frame(notebook, padding=10)
        self.help_tab = ttk.Frame(notebook, padding=10)
        notebook.add(self.single_tab, text="Один код")
        notebook.add(self.batch_tab, text="Пакетная генерация")
        notebook.add(self.help_tab, text="Проверка и справка")

        self._build_single_tab()
        self._build_batch_tab()
        self._build_help_tab()

    def _window_settings_path(self) -> Path:
        settings_root = Path(os.environ.get("APPDATA", Path.home()))
        return settings_root / "PROJECTZ GS1 DataMatrix" / "window.json"

    def _restore_window_settings(self):
        path = self._window_settings_path()
        try:
            settings = json.loads(path.read_text(encoding="utf-8"))
            geometry = settings.get("geometry")
            match = re.fullmatch(r"(\d+)x(\d+)([+-]\d+)([+-]\d+)", str(geometry or ""))
            if match:
                width, height, x, y = (int(value) for value in match.groups())
                width = max(width, 980)
                height = max(height, 700)
                screen_width = self.winfo_screenwidth()
                screen_height = self.winfo_screenheight()
                if x + 120 < 0 or y + 120 < 0 or x > screen_width - 40 or y > screen_height - 40:
                    x, y = 20, 20
                self.geometry(f"{width}x{height}{x:+d}{y:+d}")
            if settings.get("maximized") is True:
                self.after_idle(lambda: self.state("zoomed"))
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            pass

    def _save_window_settings(self):
        path = self._window_settings_path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            payload = {
                "geometry": self.geometry(),
                "maximized": self.state() == "zoomed",
            }
            temporary_path = path.with_suffix(".tmp")
            temporary_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            temporary_path.replace(path)
        except OSError:
            pass

    def _close_window(self):
        self._save_window_settings()
        self.destroy()

    def _settings_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Параметры этикетки", padding=8)

        ttk.Label(frame, text="Размер:").grid(row=0, column=0, sticky="w")
        preset = ttk.Combobox(
            frame,
            values=list(LABEL_PRESETS),
            state="readonly",
            width=19,
        )
        if not hasattr(self, "preset_var"):
            self.preset_var = tk.StringVar(value="58 × 30 мм")
            self.width_var = tk.DoubleVar(value=58.0)
            self.height_var = tk.DoubleVar(value=30.0)
            self.dpi_var = tk.IntVar(value=300)
            self.show_caption_var = tk.BooleanVar(value=True)
        preset.configure(textvariable=self.preset_var)
        preset.grid(row=0, column=1, sticky="w", padx=(5, 14))
        preset.bind("<<ComboboxSelected>>", self._on_preset)
        self.preset_box = preset

        ttk.Label(frame, text="Ширина, мм:").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(frame, from_=10, to=200, increment=0.5, textvariable=self.width_var, width=8).grid(
            row=0, column=3, padx=(5, 14)
        )

        ttk.Label(frame, text="Высота, мм:").grid(row=0, column=4, sticky="w")
        ttk.Spinbox(frame, from_=10, to=200, increment=0.5, textvariable=self.height_var, width=8).grid(
            row=0, column=5, padx=(5, 14)
        )

        ttk.Label(frame, text="DPI:").grid(row=0, column=6, sticky="w")
        ttk.Combobox(
            frame, values=(203, 300, 600), state="readonly",
            textvariable=self.dpi_var, width=7
        ).grid(row=0, column=7, padx=(5, 14))

        ttk.Checkbutton(
            frame, text="Печатать подпись", variable=self.show_caption_var
        ).grid(row=0, column=8, sticky="w")

        return frame

    def _on_preset(self, _event=None):
        value = LABEL_PRESETS.get(self.preset_var.get())
        if value:
            self.width_var.set(value[0])
            self.height_var.set(value[1])
            if value == (100.0, 73.0) and hasattr(self, "pdf_mode_var"):
                self.pdf_mode_var.set("Одна этикетка — страница по размеру")

    def _build_single_tab(self):
        settings = self._settings_frame(self.single_tab)
        settings.pack(fill="x", pady=(0, 8))

        body = ttk.Panedwindow(self.single_tab, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body, padding=(0, 0, 8, 0))
        right = ttk.LabelFrame(body, text="Предпросмотр", padding=10)
        body.add(left, weight=3)
        body.add(right, weight=2)
        right.configure(padding=6)

        ttk.Label(left, text="Код маркировки", style="Section.TLabel").pack(anchor="w")
        self.code_text = tk.Text(left, height=8, wrap="word", font=("Consolas", 10))
        self.code_text.pack(fill="x", pady=(4, 8))
        self.code_text.insert("1.0", DEFAULT_CODE)

        ttk.Label(left, text="Подпись на этикетке").pack(anchor="w")
        self.caption_var = tk.StringVar()
        ttk.Entry(left, textvariable=self.caption_var).pack(fill="x", pady=(4, 8))

        actions = ttk.Frame(left)
        actions.pack(fill="x")
        ttk.Button(actions, text="Создать и проверить", command=self.generate_single).pack(side="left")
        ttk.Button(actions, text="Вставить", command=self.paste_code).pack(side="left", padx=6)
        ttk.Button(actions, text="Очистить", command=self.clear_single).pack(side="left")

        export = ttk.LabelFrame(left, text="Экспорт", padding=8)
        export.pack(fill="x", pady=10)
        self.save_png_btn = ttk.Button(export, text="PNG", command=self.save_png, state="disabled")
        self.save_pdf_btn = ttk.Button(export, text="PDF 1 этикетка", command=self.save_single_pdf, state="disabled")
        self.save_zpl_btn = ttk.Button(export, text="ZPL", command=self.save_zpl, state="disabled")
        self.copy_structure_btn = ttk.Button(
            export,
            text="Копировать для просмотра",
            command=self.copy_structure,
            state="disabled",
        )
        self.copy_raw_btn = ttk.Button(
            export,
            text="Копировать настоящий код",
            command=self.copy_raw_code,
            state="disabled",
        )
        for widget in (
            self.save_png_btn,
            self.save_pdf_btn,
            self.save_zpl_btn,
            self.copy_structure_btn,
            self.copy_raw_btn,
        ):
            widget.pack(side="left", padx=(0, 7))

        validation = ttk.LabelFrame(left, text="Контроль", padding=8)
        validation.pack(fill="both", expand=True)
        self.validation_text = tk.Text(
            validation, height=11, wrap="word", state="disabled",
            font=("Consolas", 9)
        )
        self.validation_text.pack(fill="both", expand=True)

        self.preview_label = ttk.Label(right, anchor="center")
        self.preview_label.pack(fill="both", expand=True)
        self.preview_label.bind("<Configure>", self._schedule_preview_render)
        self.single_status = ttk.Label(right, text="Готово к работе.", wraplength=380)
        self.single_status.pack(fill="x", pady=(8, 0))

    def _build_batch_tab(self):
        self._settings_frame(self.batch_tab).pack(fill="x", pady=(0, 8))
        pdf_mode = ttk.Frame(self.batch_tab)
        pdf_mode.pack(fill="x", pady=(0, 8))
        ttk.Label(pdf_mode, text="Страницы PDF:").pack(side="left")
        self.pdf_mode_var = tk.StringVar(value="Сетка на A4")
        ttk.Combobox(
            pdf_mode, textvariable=self.pdf_mode_var, state="readonly", width=43,
            values=("Сетка на A4", "Одна этикетка на A4",
                    "Одна этикетка — страница по размеру"),
        ).pack(side="left", padx=8)
        settings = ttk.LabelFrame(self.batch_tab, text="Пакетный экспорт", padding=8)
        settings.pack(fill="x", pady=(0, 8))

        ttk.Button(settings, text="Открыть XLSX/CSV", command=self.load_batch).pack(side="left")
        ttk.Button(settings, text="Скачать шаблон XLSX", command=self.save_template).pack(side="left", padx=6)
        ttk.Button(settings, text="Проверить все", command=self.validate_batch).pack(side="left")

        ttk.Separator(settings, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Label(settings, text="Выравнивание сетки:").pack(side="left")
        self.pdf_position_var = tk.StringVar(value="По центру")
        self.pdf_position_box = ttk.Combobox(
            settings,
            textvariable=self.pdf_position_var,
            values=("По центру", "Сверху слева", "Сверху справа"),
            state="readonly",
            width=17,
        )
        self.pdf_position_box.pack(side="left", padx=(5, 10))

        ttk.Button(settings, text="PNG в папку", command=self.export_batch_png).pack(side="left", padx=(0, 6))
        ttk.Button(settings, text="Сохранить PDF", command=self.export_batch_pdf).pack(side="left")
        ttk.Button(settings, text="ZPL в папку", command=self.export_batch_zpl).pack(side="left", padx=6)

        pdf_layout = ttk.LabelFrame(self.batch_tab, text="Раскладка PDF A4", padding=8)
        pdf_layout.pack(fill="x", pady=(0, 8))

        ttk.Label(pdf_layout, text="Поля листа, мм:").pack(side="left")
        self.pdf_margin_var = tk.DoubleVar(value=12.0)
        ttk.Spinbox(
            pdf_layout,
            from_=0,
            to=50,
            increment=0.5,
            textvariable=self.pdf_margin_var,
            width=7,
        ).pack(side="left", padx=(5, 14))

        ttk.Label(pdf_layout, text="Отступ между кодами, мм:").pack(side="left")
        self.pdf_gap_var = tk.DoubleVar(value=5.0)
        ttk.Spinbox(
            pdf_layout,
            from_=0,
            to=50,
            increment=0.5,
            textvariable=self.pdf_gap_var,
            width=7,
        ).pack(side="left", padx=(5, 14))

        ttk.Label(
            pdf_layout,
            text="Заполнение: слева направо, сверху вниз",
        ).pack(side="left")

        table_frame = ttk.Frame(self.batch_tab)
        table_frame.pack(fill="both", expand=True)
        columns = ("row", "gtin", "serial", "caption", "filename", "status")
        self.batch_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        headings = {
            "row": "Строка", "gtin": "GTIN", "serial": "Серийный номер",
            "caption": "Подпись", "filename": "Имя файла", "status": "Статус"
        }
        widths = {"row": 60, "gtin": 125, "serial": 180, "caption": 200, "filename": 150, "status": 240}
        for col in columns:
            self.batch_tree.heading(col, text=headings[col])
            self.batch_tree.column(col, width=widths[col], anchor="w")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.batch_tree.yview)
        self.batch_tree.configure(yscrollcommand=scrollbar.set)
        self.batch_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.batch_status = ttk.Label(
            self.batch_tab,
            text=(
                "Достаточно одной колонки с кодами. "
                "Подпись и имя файла заполняются автоматически."
            ),
        )
        self.batch_status.pack(fill="x", pady=(8, 0))

    def _build_help_tab(self):
        text = tk.Text(self.help_tab, wrap="word", font=("Segoe UI", 10))
        text.pack(fill="both", expand=True)
        text.insert(
            "1.0",
            f"""PROJECTZ GS1 DataMatrix {APP_VERSION}

Зафиксированная структура:
{{FNC1}}01GTIN21SERIAL{{GS}}91KEY{{GS}}92CRYPTO

Приложение:
• проверяет формат AI 01, 21, 91 и 92;
• проверяет контрольную цифру GTIN;
• проверяет допустимые ASCII-символы;
• проверяет Base64-подобную структуру AI 92;
• генерирует DataMatrix через BWIPP/treepoem;
• независимо считывает результат через ZXing-C++;
• требует AIM-идентификатор ]d2;
• сравнивает считанные данные посимвольно;
• копирует человекочитаемую структуру с {{FNC1}}/{{GS}};
• копирует настоящий код с двумя ASCII GS (29) для сайта проверки;
• экспортирует PNG с именем GTIN_СЕРИЙНЫЙ_НОМЕР.png;
• позволяет выбрать положение DataMatrix в PDF: центр, сверху слева или сверху справа;
• выполняет пакетную обработку XLSX/CSV.

Структура намеренно НЕ строится через автоматический gs1datamatrix:
первый знак — FNC1, а разделители после AI 21 и AI 91 —
буквальный ASCII GS с кодом 29.

Формат XLSX/CSV:
Код маркировки | Подпись | Имя файла
"""
        )
        text.configure(state="disabled")

    def paste_code(self):
        try:
            value = self.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("Буфер", "В буфере нет текста.")
            return
        self.code_text.delete("1.0", "end")
        self.code_text.insert("1.0", value)

    def clear_single(self):
        self.code_text.delete("1.0", "end")
        self.caption_var.set("")
        self.current_code = ""
        self.current_payload = ""
        self.current_symbol = None
        self.current_label = None
        self.preview_label.configure(image="")
        self.preview_photo = None
        self._set_validation("")
        self._set_export_state(False)
        self.single_status.configure(text="Поле очищено.")

    def _set_validation(self, text: str):
        self.validation_text.configure(state="normal")
        self.validation_text.delete("1.0", "end")
        self.validation_text.insert("1.0", text)
        self.validation_text.configure(state="disabled")

    def _set_export_state(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for widget in (
            self.save_png_btn,
            self.save_pdf_btn,
            self.save_zpl_btn,
            self.copy_structure_btn,
            self.copy_raw_btn,
        ):
            widget.configure(state=state)

    def _schedule_preview_render(self, _event=None):
        if self.current_label is None:
            return
        if self._preview_resize_job is not None:
            self.after_cancel(self._preview_resize_job)
        self._preview_resize_job = self.after(40, self._render_preview)

    def _render_preview(self):
        self._preview_resize_job = None
        if self.current_label is None:
            return
        available_width = self.preview_label.winfo_width() - 12
        available_height = self.preview_label.winfo_height() - 12
        if available_width <= 1 or available_height <= 1:
            self._preview_resize_job = self.after_idle(self._render_preview)
            return

        preview = ImageOps.contain(
            self.current_label,
            (available_width, available_height),
            method=Image.Resampling.NEAREST,
        )
        self.preview_photo = ImageTk.PhotoImage(preview)
        self.preview_label.configure(image=self.preview_photo)

    def get_label_settings(self):
        width = float(self.width_var.get())
        height = float(self.height_var.get())
        dpi = int(self.dpi_var.get())
        if width <= 0 or height <= 0:
            raise ValueError("Размер этикетки должен быть больше нуля.")
        return width, height, dpi

    def generate_single(self):
        code = normalize_code(self.code_text.get("1.0", "end"))
        ok, messages, parts = validate_code(code, strict_gtin=True)
        report = ["ПРЕДВАРИТЕЛЬНАЯ ПРОВЕРКА"]
        report.extend(("✓ " if ok else "• ") + item for item in messages)

        if not ok or parts is None:
            self._set_validation("\n".join(report))
            self._set_export_state(False)
            self.single_status.configure(text="Код не прошёл проверку.")
            messagebox.showerror("Ошибка структуры", "\n".join(messages))
            return

        try:
            width, height, dpi = self.get_label_settings()
            payload = build_expected_payload(code)
            symbol = generate_symbol(code, scale=4)
            decoded = decode_and_verify(symbol, payload)

            caption = self.caption_var.get().strip() or parts.serial
            label = compose_label(
                symbol, width, height, dpi, caption, self.show_caption_var.get()
            )

            self.current_code = code
            self.current_payload = payload
            self.current_symbol = symbol
            self.current_label = label

            report.extend([
                "",
                "НЕЗАВИСИМАЯ ПРОВЕРКА ИЗОБРАЖЕНИЯ",
                f"✓ Формат: {decoded.format}",
                f"✓ AIM: {decoded.symbology_identifier}",
                f"✓ Тип: {decoded.content_type}",
                "✓ Первый служебный символ: FNC1",
                "✓ После AI 21: ASCII GS (29)",
                "✓ После AI 91: ASCII GS (29)",
                "✓ Считанные данные совпадают посимвольно",
                "",
                "ФАКТИЧЕСКАЯ СТРУКТУРА",
                visible_payload(payload),
                "",
                f"Этикетка: {width:g} × {height:g} мм, {dpi} DPI",
            ])
            self._set_validation("\n".join(report))

            self._schedule_preview_render()
            self._set_export_state(True)
            self.single_status.configure(text="Создано и проверено успешно.")
        except Exception as exc:
            self._set_export_state(False)
            self._set_validation("\n".join(report + ["", f"ОШИБКА: {exc}"]))
            self.single_status.configure(text=f"Ошибка: {exc}")
            messagebox.showerror("Ошибка генерации", str(exc))

    def save_png(self):
        if self.current_label is None:
            return
        parts = parse_code(self.current_code)
        path = filedialog.asksaveasfilename(
            title="Сохранить PNG",
            defaultextension=".png",
            initialfile=f"{parts.gtin}_{safe_filename(parts.serial)}.png",
            filetypes=[("PNG", "*.png")],
        )
        if path:
            dpi = int(self.dpi_var.get())
            self.current_label.save(path, "PNG", dpi=(dpi, dpi), optimize=False)
            self.single_status.configure(text=f"PNG сохранён: {path}")

    def save_single_pdf(self):
        if self.current_label is None:
            return
        parts = parse_code(self.current_code)
        path = filedialog.asksaveasfilename(
            title="Сохранить PDF",
            defaultextension=".pdf",
            initialfile=f"GS1_{parts.gtin}_{safe_filename(parts.serial)}.pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return

        width, height, _dpi = self.get_label_settings()
        buffer = io.BytesIO()
        self.current_label.save(buffer, "PNG")
        buffer.seek(0)

        from reportlab.lib.utils import ImageReader
        c = canvas.Canvas(path, pagesize=(width * mm, height * mm))
        c.drawImage(ImageReader(buffer), 0, 0, width=width * mm, height=height * mm)
        c.showPage()
        c.save()
        self.single_status.configure(text=f"PDF сохранён: {path}")

    def save_zpl(self):
        if self.current_label is None:
            return
        parts = parse_code(self.current_code)
        path = filedialog.asksaveasfilename(
            title="Сохранить ZPL",
            defaultextension=".zpl",
            initialfile=f"GS1_{parts.gtin}_{safe_filename(parts.serial)}.zpl",
            filetypes=[("ZPL", "*.zpl"), ("TXT", "*.txt")],
        )
        if path:
            width, height, dpi = self.get_label_settings()
            Path(path).write_text(
                image_to_zpl(self.current_label, width, height, dpi),
                encoding="ascii"
            )
            self.single_status.configure(text=f"ZPL сохранён: {path}")

    def copy_structure(self):
        """
        Копирует человекочитаемое представление.
        В нём служебные символы показаны как {FNC1} и {GS}.
        Такой вариант удобен для проверки глазами, но не для загрузки на сайт.
        """
        if not self.current_payload:
            return
        visible = visible_payload(self.current_payload)
        self.clipboard_clear()
        self.clipboard_append(visible)
        self.update_idletasks()
        self.single_status.configure(
            text="Скопировано для просмотра: {FNC1} и {GS} показаны текстом."
        )

    def copy_raw_code(self):
        """
        Копирует фактическую строку GS1:
        - без текстового {FNC1} в начале;
        - с настоящими ASCII GS (29) после AI 21 и AI 91.

        После записи в буфер выполняется обратное чтение и проверяется,
        что в строке присутствуют ровно два разделителя ASCII 29.
        """
        if not self.current_payload:
            return

        raw_code = self.current_payload
        expected_gs_count = 2

        try:
            self.clipboard_clear()
            self.clipboard_append(raw_code)
            self.update_idletasks()

            copied = self.clipboard_get()
            actual_gs_count = copied.count(GS)

            if copied != raw_code:
                raise RuntimeError(
                    "Буфер обмена изменил содержимое строки. "
                    "Попробуйте вставить код напрямую из приложения."
                )

            if actual_gs_count != expected_gs_count:
                raise RuntimeError(
                    f"В буфере найдено разделителей GS: {actual_gs_count}; "
                    f"ожидалось: {expected_gs_count}."
                )

            self.single_status.configure(
                text=(
                    "Настоящий код скопирован. "
                    "В буфере подтверждено 2 разделителя ASCII GS (29)."
                )
            )
            messagebox.showinfo(
                "Код скопирован",
                "В буфер помещён настоящий код маркировки.\n\n"
                "Проверка выполнена успешно:\n"
                "• GS после AI 21 — найден;\n"
                "• GS после AI 91 — найден;\n"
                "• всего разделителей ASCII 29: 2.\n\n"
                "Этот вариант используйте для вставки на сайт проверки.",
            )
        except Exception as exc:
            self.single_status.configure(text=f"Ошибка копирования: {exc}")
            messagebox.showerror("Ошибка буфера обмена", str(exc))


    def save_template(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить шаблон",
            defaultextension=".xlsx",
            initialfile="Шаблон_массовой_генерации_GS1.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            create_template_xlsx(path)
            self.batch_status.configure(text=f"Шаблон сохранён: {path}")

    def load_batch(self):
        path = filedialog.askopenfilename(
            title="Открыть файл с кодами",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv")],
        )
        if not path:
            return
        try:
            self.batch_items = read_batch_file(path)
            self.batch_source = path
            self._refresh_batch_tree()
            self.batch_status.configure(
                text=f"Загружено строк: {len(self.batch_items)}. Выполните проверку."
            )
        except Exception as exc:
            messagebox.showerror("Ошибка файла", str(exc))

    def validate_batch(self):
        if not self.batch_items:
            messagebox.showwarning("Пакет", "Сначала загрузите XLSX или CSV.")
            return

        ok_count = 0
        for item in self.batch_items:
            code = normalize_code(item.source_code)
            ok, messages, parts = validate_code(code, strict_gtin=True)
            item.parts = parts
            if ok and parts:
                item.source_code = code
                item.parts = parts
                if not item.caption.strip():
                    item.caption = parts.serial
                if not item.filename.strip():
                    item.filename = f"{parts.gtin}_{parts.serial}"
                item.status = "Готов"
                item.error = ""
                ok_count += 1
            else:
                item.status = "Ошибка"
                item.error = "; ".join(messages)
        self._refresh_batch_tree()
        self.batch_status.configure(
            text=f"Проверено: {len(self.batch_items)}; готово: {ok_count}; "
                 f"ошибок: {len(self.batch_items) - ok_count}."
        )

    def _refresh_batch_tree(self):
        self.batch_tree.delete(*self.batch_tree.get_children())
        for item in self.batch_items:
            parts = item.parts
            status = item.status
            if item.error:
                status += f": {item.error}"
            self.batch_tree.insert("", "end", values=(
                item.row_no,
                parts.gtin if parts else "",
                parts.serial if parts else "",
                item.caption,
                item.filename,
                status,
            ))

    def _show_batch_progress(self, operation: str, current: int, total: int, details: str = ""):
        percent = round(current / total * 100) if total else 0
        text = f"{operation}: {current} из {total} ({percent}%)"
        if details:
            text += f"; {details}"
        self.batch_status.configure(text=text)
        self.update_idletasks()

    def _ensure_batch_validated(self) -> List[BatchItem]:
        self.validate_batch()
        valid = [item for item in self.batch_items if item.status == "Готов" and item.parts]
        if not valid:
            raise ValueError("Нет строк, готовых к генерации.")
        return valid

    def _make_batch_label(self, item: BatchItem) -> Image.Image:
        width, height, dpi = self.get_label_settings()
        payload = build_expected_payload(item.source_code)
        symbol = generate_symbol(item.source_code, scale=4)
        decode_and_verify(symbol, payload)
        caption = item.caption.strip() or item.parts.serial
        return compose_label(
            symbol, width, height, dpi, caption, self.show_caption_var.get()
        )

    def export_batch_png(self):
        try:
            valid = self._ensure_batch_validated()
            folder = filedialog.askdirectory(title="Папка для PNG")
            if not folder:
                return
            dpi = int(self.dpi_var.get())
            for index, item in enumerate(valid, start=1):
                self._show_batch_progress("Создание PNG", index, len(valid))
                label = self._make_batch_label(item)
                name = item.filename or f"{item.parts.gtin}_{item.parts.serial}"
                label.save(
                    Path(folder) / f"{safe_filename(name)}.png",
                    "PNG", dpi=(dpi, dpi), optimize=False
                )
            self.batch_status.configure(text=f"PNG сохранены: {len(valid)} шт. в {folder}")
        except Exception as exc:
            messagebox.showerror("Пакетный PNG", str(exc))

    def export_batch_zpl(self):
        try:
            valid = self._ensure_batch_validated()
            folder = filedialog.askdirectory(title="Папка для ZPL")
            if not folder:
                return
            width, height, dpi = self.get_label_settings()
            for index, item in enumerate(valid, start=1):
                self._show_batch_progress("Создание ZPL", index, len(valid))
                label = self._make_batch_label(item)
                name = item.filename or f"{item.parts.gtin}_{item.parts.serial}"
                Path(folder, f"{safe_filename(name)}.zpl").write_text(
                    image_to_zpl(label, width, height, dpi),
                    encoding="ascii"
                )
            self.batch_status.configure(text=f"ZPL сохранены: {len(valid)} шт. в {folder}")
        except Exception as exc:
            messagebox.showerror("Пакетный ZPL", str(exc))

    def export_batch_pdf(self):
        """
        Создаёт PDF сеткой на A4 или по одной этикетке на странице.
        """
        try:
            valid = self._ensure_batch_validated()

            mode = self.pdf_mode_var.get()
            label_page = mode == "Одна этикетка — страница по размеру"
            position = self.pdf_position_var.get().strip()
            allowed_positions = ("По центру", "Сверху слева", "Сверху справа")
            if position not in allowed_positions:
                raise ValueError("Выберите выравнивание сетки в PDF.")

            margin_mm = 0.0 if label_page else float(self.pdf_margin_var.get())
            gap_mm = 0.0 if mode != "Сетка на A4" else float(self.pdf_gap_var.get())
            if margin_mm < 0 or gap_mm < 0:
                raise ValueError("Поля листа и отступ между кодами не могут быть отрицательными.")

            label_width_mm, label_height_mm, _ = self.get_label_settings()

            path = filedialog.asksaveasfilename(
                title=f"Сохранить PDF — {mode}",
                defaultextension=".pdf",
                initialfile=(
                    f"GS1_DataMatrix_{label_width_mm:g}x{label_height_mm:g}_"
                    + safe_filename(position.replace(" ", "_"))
                    + ".pdf"
                ),
                filetypes=[("PDF", "*.pdf")],
            )
            if not path:
                return

            from reportlab.lib.utils import ImageReader

            page_w, page_h = A4
            pdf_dpi = 300
            page_margin = margin_mm * mm
            gap = gap_mm * mm
            label_w = label_width_mm * mm
            label_h = label_height_mm * mm
            if label_page:
                page_w, page_h = label_w, label_h

            available_w = page_w - 2 * page_margin
            available_h = page_h - 2 * page_margin
            if available_w <= 0 or available_h <= 0:
                raise ValueError("Поля листа не оставляют места для этикеток на A4.")

            columns = math.floor((available_w + gap) / (label_w + gap))
            rows = math.floor((available_h + gap) / (label_h + gap))
            if columns < 1 or rows < 1:
                raise ValueError(
                    f"Этикетка {label_width_mm:g} × {label_height_mm:g} мм не помещается "
                    f"на A4 с полями {margin_mm:g} мм без изменения размера."
                )

            if mode != "Сетка на A4":
                columns = rows = 1
            labels_per_page = columns * rows
            page_count = math.ceil(len(valid) / labels_per_page)
            grid_w = columns * label_w + (columns - 1) * gap
            grid_h = rows * label_h + (rows - 1) * gap

            if position == "Сверху слева":
                grid_x = page_margin
                grid_top = page_h - page_margin
            elif position == "Сверху справа":
                grid_x = page_w - page_margin - grid_w
                grid_top = page_h - page_margin
            else:
                grid_x = (page_w - grid_w) / 2
                grid_top = (page_h + grid_h) / 2

            c = canvas.Canvas(path, pagesize=(page_w, page_h))
            c.setTitle(
                f"PROJECTZ GS1 DataMatrix — сетка {columns} × {rows}"
            )

            for index, item in enumerate(valid):
                if index and index % labels_per_page == 0:
                    c.showPage()

                current_page = index // labels_per_page + 1
                self._show_batch_progress(
                    "Создание PDF",
                    index + 1,
                    len(valid),
                    f"страница {current_page} из {page_count}",
                )

                # Генерация и обязательное обратное считывание.
                original_dpi = self.dpi_var.get()
                try:
                    self.dpi_var.set(pdf_dpi)
                    label = self._make_batch_label(item)
                finally:
                    self.dpi_var.set(original_dpi)

                buffer = io.BytesIO()
                label.save(
                    buffer,
                    "PNG",
                    dpi=(pdf_dpi, pdf_dpi),
                    optimize=False,
                )
                buffer.seek(0)

                slot = index % labels_per_page
                row = slot // columns
                column = slot % columns
                x = grid_x + column * (label_w + gap)
                y = grid_top - label_h - row * (label_h + gap)

                c.drawImage(
                    ImageReader(buffer),
                    x,
                    y,
                    width=label_w,
                    height=label_h,
                    preserveAspectRatio=True,
                    mask="auto",
                )

            c.save()
            self.batch_status.configure(
                text=(
                    f"PDF сохранён: {path}; кодов: {len(valid)}; страниц: {page_count}; "
                    f"сетка: {columns} × {rows}; отступ: {gap_mm:g} мм; 300 DPI."
                )
            )
        except Exception as exc:
            messagebox.showerror("Пакетный PDF", str(exc))


def run_self_test(output_path: str) -> None:
    """Exercise the packaged encoder/decoder without opening the GUI."""
    code = normalize_code(DEFAULT_CODE)
    ok, messages, _ = validate_code(code, strict_gtin=False)
    if not ok:
        raise RuntimeError("; ".join(messages))
    symbol = generate_symbol(code, scale=2)
    decoded = decode_and_verify(symbol, build_expected_payload(code))
    Path(output_path).write_text(
        f"OK\nAIM={decoded.symbology_identifier}\nSIZE={symbol.width}x{symbol.height}\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        run_self_test(sys.argv[2])
    else:
        App().mainloop()
