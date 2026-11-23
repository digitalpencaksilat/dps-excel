import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

LAUNCHER = Path(__file__).resolve().parents[1] / "dps-excel"
ZWNJ = "\u200c"
sys.path.insert(0, str(LAUNCHER.parent / "src"))
from cli import choose_file, run


class PickerTest(unittest.TestCase):
    def test_macos_picker(self):
        with patch("cli.sys.platform", "darwin"), patch("cli.sys.stdin.isatty", return_value=True), \
                patch("cli.shutil.which", return_value="/usr/bin/osascript"), \
                patch("cli.subprocess.run") as run, patch("cli.questionary.path") as manual:
            run.return_value = subprocess.CompletedProcess([], 0, "/tmp/data warga.xlsx\n", "")
            self.assertEqual(choose_file(), Path("/tmp/data warga.xlsx"))
            self.assertIn("Pilih file Excel (.xlsx)", run.call_args.args[0][-1])
            self.assertNotIn("of type", run.call_args.args[0][-1])
            run.return_value = subprocess.CompletedProcess([], 1, "", "User canceled. (-128)")
            self.assertIsNone(choose_file())
            manual.assert_not_called()
            run.return_value = subprocess.CompletedProcess([], 1, "", "No GUI")
            manual.return_value.ask.return_value = "/tmp/fallback.xlsx"
            self.assertEqual(choose_file(), Path("/tmp/fallback.xlsx"))


class CleanerTest(unittest.TestCase):
    def test_version_matches_version_file(self):
        version = LAUNCHER.with_name("VERSION").read_text(encoding="ascii").strip()
        result = self.cli("--version")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), f"dps-excel {version}")

    def test_extra_column_without_automatic_header(self):
        book = Workbook()
        sheet = book.active
        sheet["A2"] = "Nama"
        sheet["B2"] = "Alamat"
        sheet["A3"] = f"O'{ZWNJ}Neil"
        sheet["A3"].number_format = "0.00"
        sheet["B3"] = f"A/{ZWNJ}B"
        book.save(self.source)
        with patch("cli.sys.stdin.isatty", return_value=True), \
                patch("cli.questionary.confirm") as confirm, \
                patch("cli.questionary.text") as header, \
                patch("cli.questionary.checkbox") as checkbox:
            confirm.return_value.ask.side_effect = [True, True, True]
            header.return_value.ask.return_value = "2"
            checkbox.return_value.ask.side_effect = lambda: [
                checkbox.call_args.kwargs["choices"][0].value
            ]
            self.assertEqual(run(self.source, interactive=True), 0)
        result = load_workbook(self.output)
        try:
            self.assertEqual(result.active["A3"].value, "O'Neil")
            self.assertEqual(result.active["A3"].number_format, "0.00")
            self.assertFalse(result.active["A3"].quotePrefix)
            self.assertEqual(result.active["B3"].value, f"A/{ZWNJ}B")
        finally:
            result.close()

    def test_wizard_rejects_other_extensions(self):
        result = self.cli(text="data warga.xls\n")
        self.assertEqual(result.returncode, 1)
        self.assertIn("harus berekstensi .xlsx", result.stderr)
        self.assertFalse(self.output.exists())

    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)
        self.output = self.folder / "data warga_bersih.xlsx"
        self.source = self.folder / "data warga.xlsx"

    def cli(self, *args, text=None):
        return subprocess.run(
            [LAUNCHER, *args], cwd=self.folder, input=text,
            capture_output=True, text=True, check=False,
        )

    def test_wizard_and_direct_mode(self):
        book = Workbook()
        sheet = book.active
        sheet["C1"] = f"atas{ZWNJ}"
        sheet["C5"] = "  no.   NIK  "
        sheet["F5"] = "Nomor KK"
        sheet["C6"] = f"001{ZWNJ}234"
        sheet["F6"] = f"{ZWNJ}=1+1"
        sheet["C7"] = 1234
        sheet["F7"] = f'="{ZWNJ}"'
        sheet["D6"] = f"lain{ZWNJ}"
        book.create_sheet("Lain")["A1"] = f"lain{ZWNJ}"
        book.create_sheet("Kedua")["A2"] = "No KK"
        book["Kedua"]["A3"] = f"09{ZWNJ}"
        book.save(self.source)
        before = self.source.read_bytes()

        cancelled = self.cli(text="data warga.xlsx\nn\n")
        self.assertEqual(cancelled.returncode, 0, cancelled.stderr)
        self.assertIn("Baris header", cancelled.stdout)
        self.assertIn("C:", cancelled.stdout)
        self.assertFalse(self.output.exists())

        result = self.cli("clean", self.source.name)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("3 sel", result.stdout)
        self.assertEqual(self.source.read_bytes(), before)
        with self.output.open("rb") as stream:
            output = load_workbook(stream)
            sheet = output.active
            self.assertEqual(sheet["C6"].value, "001234")
            self.assertEqual(sheet["C6"].number_format, "@")
            self.assertEqual(sheet["F6"].data_type, "s")
            self.assertEqual(sheet["F6"].value, "=1+1")
            self.assertEqual(sheet["F7"].value, f'="{ZWNJ}"')
            self.assertEqual(sheet["C7"].value, 1234)
            self.assertEqual(sheet["D6"].value, f"lain{ZWNJ}")
            self.assertEqual(sheet["C1"].value, f"atas{ZWNJ}")
            self.assertEqual(output["Lain"]["A1"].value, f"lain{ZWNJ}")
            self.assertEqual(output["Kedua"]["A3"].value, "09")
            output.close()
        previous = self.output.read_bytes()
        self.assertNotEqual(self.cli("clean", self.source.name).returncode, 0)
        self.assertEqual(self.output.read_bytes(), previous)

    def test_missing_header_and_zero_changes(self):
        book = Workbook()
        book.save(self.source)
        self.assertIn("tidak ditemukan", self.cli("clean", self.source.name).stdout)
        self.assertFalse(self.output.exists())
        book.active["A1"] = "NIK"
        book.active["A2"] = "00123"
        book.save(self.source)
        result = self.cli(text="data warga.xlsx\nya\n")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("0 sel", result.stdout)
        self.assertTrue(self.output.exists())
