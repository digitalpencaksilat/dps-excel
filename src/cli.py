import argparse
import shutil
import subprocess
import sys
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import questionary
from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text
from rich.filesize import decimal

from cleaner import clean_columns, detect_columns, save_new

console = Console()
PROMPT = {"qmark": "", "style": questionary.Style([
    ("qmark", "fg:#d32f2f"), ("question", "bold fg:#d32f2f"),
    ("answer", "fg:#eeeeee"), ("pointer", "fg:#d32f2f"),
    ("highlighted", "fg:#d32f2f"),
])}


def choose_file():
    command = None
    if sys.stdin.isatty():
        if sys.platform == "darwin" and (picker := shutil.which("osascript")):
            command = [picker, "-e", 'POSIX path of (choose file with prompt '
                       '"Pilih file Excel (.xlsx)")']
        elif sys.platform.startswith("linux"):
            if picker := shutil.which("zenity"):
                command = [picker, "--file-selection", "--title=Pilih file Excel",
                           "--file-filter=Excel | *.xlsx *.XLSX"]
            elif picker := shutil.which("kdialog"):
                command = [picker, "--getopenfilename", ".", "Excel (*.xlsx *.XLSX)"]

        if command:
            try:
                result = subprocess.run(command, capture_output=True, text=True, check=False)
            except OSError:
                pass
            else:
                if result.returncode == 0:
                    path = result.stdout.removesuffix("\n")
                    return Path(path) if path else None
                if "(-128)" in result.stderr or (
                    sys.platform.startswith("linux") and result.returncode == 1
                    and not result.stderr.strip()
                ):
                    return None
        console.print("Pemilih file tidak tersedia. Masukkan path secara manual.", style="yellow")
        path = questionary.path("Path file Excel (.xlsx):", **PROMPT).ask()
    else:
        path = input("Path file Excel (.xlsx): ")

    path = (path or "").strip()
    if len(path) >= 2 and path[0] == path[-1] and path[0] in {"'", '"'}:
        path = path[1:-1]
    return Path(path) if path else None


def choose_extra_columns(workbook, targets):
    automatic = {sheet: (row, {cell.column for cell in headers})
                 for sheet, row, headers in targets}
    selected = []
    for sheet in workbook.worksheets:
        include = questionary.confirm(f"Tambah kolom di sheet {sheet.title!r}?",
                                       default=False, **PROMPT).ask()
        if include is None:
            return None
        if not include:
            continue
        if sheet in automatic:
            row, excluded = automatic[sheet]
        else:
            answer = questionary.text(
                f"Baris header (1-{sheet.max_row}):", default="1", **PROMPT,
                validate=lambda value: value.isascii() and value.isdecimal()
                and len(value) <= len(str(sheet.max_row)) and 1 <= int(value) <= sheet.max_row,
            ).ask()
            if answer is None:
                return None
            row, excluded = int(answer), set()
        candidates = [cell for cell in sheet[row]
                      if cell.value is not None and cell.column not in excluded]
        if not candidates:
            console.print("Tidak ada header tambahan pada baris tersebut.", style="yellow")
            continue
        columns = questionary.checkbox(
            "Pilih kolom (Spasi untuk memilih, Enter untuk lanjut):",
            choices=[questionary.Choice(f"{cell.column_letter}: {cell.value!r}", value=cell)
                     for cell in candidates], **PROMPT,
        ).ask()
        if columns is None:
            return None
        if columns:
            selected.append((sheet, row, columns))
    return selected


def run(source, output=None, interactive=False):
    source = source.expanduser()
    output = (output or source.with_name(f"{source.stem}_bersih.xlsx")).expanduser()
    if source.suffix.lower() != ".xlsx" or output.suffix.lower() != ".xlsx":
        print("Input dan output harus berekstensi .xlsx.", file=sys.stderr)
        return 1
    if output.exists() or output.is_symlink():
        raise FileExistsError(f"File tujuan sudah ada: {str(output)!r}")

    if interactive:
        table = Table(show_header=False, box=box.ASCII, border_style="red")
        table.add_row("File terpilih", Text(source.name, style="bold red"))
        table.add_row("Lokasi", Text(str(source.absolute())))
        table.add_row("Ukuran", decimal(source.stat().st_size))
        table.add_row("Output", Text(str(output.absolute())))
        console.print(table)

    workbook = load_workbook(source, rich_text=True)
    try:
        targets = list(detect_columns(workbook))
        extra = []
        if interactive and sys.stdin.isatty():
            add = questionary.confirm("Tambah kolom lain?", default=False, **PROMPT).ask()
            if add:
                extra = choose_extra_columns(workbook, targets)
            if add is None or extra is None:
                print("Dibatalkan.")
                return 0
        if not targets and not extra:
            print("Header NIK/KK tidak ditemukan. Tidak ada file hasil.")
            return 1

        if interactive:
            table = Table("Sheet", "Baris header", "Kolom target", box=box.ASCII,
                          border_style="red")
            console.print("Karakter yang akan dihapus: U+200C.")
            for sheet, row, headers in targets + extra:
                columns = ", ".join(f"{c.column_letter}: {c.value!r}" for c in headers)
                table.add_row(Text(sheet.title), str(row), Text(columns))
            console.print(table)
            if sys.stdin.isatty():
                confirmed = questionary.confirm("Lanjutkan?", default=False, **PROMPT).ask()
            else:
                confirmed = input("Lanjutkan? [y/N] ").strip().casefold() in {"y", "ya"}
            if not confirmed:
                print("Dibatalkan.")
                return 0

        count = clean_columns(targets) + clean_columns(extra, identifier_format=False)
        save_new(workbook, output)
        if not count:
            print("Header ditemukan, tetapi tidak ada sel yang perlu dibersihkan.")
        print(f"{count} sel dibersihkan. Hasil: {str(output)!r}")
        return 0
    finally:
        workbook.close()


def main():
    parser = argparse.ArgumentParser(prog="dps-excel", description="Pembersih NIK/KK di Excel.")
    version = (Path(__file__).resolve().parents[1] / "VERSION").read_text(encoding="ascii").strip()
    parser.add_argument("--version", action="version", version=f"%(prog)s {version}")
    commands = parser.add_subparsers(dest="command")
    clean = commands.add_parser("clean", help="Bersihkan file tanpa wizard.")
    clean.add_argument("input", type=Path)
    clean.add_argument("-o", "--output", type=Path)
    args = parser.parse_args()
    try:
        if args.command == "clean":
            return run(args.input, args.output)
        console.print(Panel.fit("[bold red]DPS EXCEL[/bold red]\nPembersih kolom NIK dan nomor KK",
                                border_style="red", box=box.ASCII))
        path = choose_file()
        if path is None:
            print("Dibatalkan.")
            return 0
        return run(path, interactive=True)
    except (EOFError, KeyboardInterrupt):
        print("\nDibatalkan.", file=sys.stderr)
        return 130
    except (OSError, InvalidFileException, BadZipFile, ValueError) as error:
        print(f"Gagal memproses workbook ({type(error).__name__}). "
              "Periksa format file, lokasi, dan pastikan file hasil belum ada.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
