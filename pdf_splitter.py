from __future__ import annotations

import os
import shutil
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string
from pypdf import PdfReader, PdfWriter

# Windowsのファイル名に使えない文字と、予約済みの名前をまとめます。
INVALID_FILENAME_CHARS = set('<>:"/\\|?*')
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


# Excelで指定したセルを読み、保存ファイル名に使える文字列か確認します。
def read_base_name(excel_path: Path, sheet_name: str, cell_address: str) -> str:
    # A1 や B3 のようなセル番地に整えます。
    address = cell_address.strip().upper()
    try:
        coordinate_from_string(address)
    except Exception as exc:
        raise ValueError("セル番地を正しく入力してください（例: A1、B3）。") from exc

    formula_workbook = None
    value_workbook = None
    try:
        # 数式そのものと、Excelに保存された計算結果を別々に読み込みます。
        # openpyxl reads formula text and Excel's saved result separately; it does not calculate formulas.
        formula_workbook = load_workbook(excel_path, read_only=True, data_only=False)
        value_workbook = load_workbook(excel_path, read_only=True, data_only=True)
        if sheet_name not in formula_workbook.sheetnames:
            raise ValueError(f'指定したシート「{sheet_name}」がExcelにありません。')
        formula_cell = formula_workbook[sheet_name][address]
        value = value_workbook[sheet_name][address].value
        if formula_cell.data_type == "f" and (value is None or not str(value).strip()):
            raise ValueError(
                f"指定したセル（{sheet_name}!{address}）は数式セルですが、保存済みの計算結果がありません。\n"
                "Excelでブックを開き、数式を再計算してから保存し、もう一度お試しください。"
            )
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f"Excelファイルを読み込めませんでした。\n{exc}") from exc
    finally:
        if formula_workbook is not None:
            formula_workbook.close()
        if value_workbook is not None:
            value_workbook.close()

    if value is None or not str(value).strip():
        raise ValueError(f"指定したセル（{sheet_name}!{address}）が空です。ベースファイル名を入力してください。")

    # 前後の空白を取り除いて、連番を付ける前の基本名にします。
    name = str(value).strip()
    if any(ch in INVALID_FILENAME_CHARS for ch in name) or any(ord(ch) < 32 for ch in name):
        raise ValueError('セルの値にファイル名として使えない文字（\\ / : * ? " < > | など）が含まれています。')
    if name.endswith((" ", ".")):
        raise ValueError("セルの値の末尾に空白またはピリオドは使えません。")
    if name.split(".", 1)[0].upper() in WINDOWS_RESERVED_NAMES:
        raise ValueError(f'「{name}」はWindowsで使用できない予約名です。')
    return name

# PDFを読み込み、セルの名前にページ番号を付けて1ページずつ保存します。
def split_pdf(pdf_path: Path, excel_path: Path, output_dir: Path, sheet_name: str, cell_address: str) -> list[Path]:
    # 入力ファイルと出力先が正しい場所にあるか、処理前に確認します。
    if not pdf_path.is_file():
        raise ValueError("指定したPDFファイルが見つかりません。")
    if not excel_path.is_file():
        raise ValueError("指定したExcelファイルが見つかりません。")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("PDFファイルを選択してください。")
    if excel_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Excelファイル（.xlsx または .xlsm）を選択してください。")
    if not output_dir.is_dir():
        raise ValueError("出力先フォルダーが見つかりません。")

    base_name = read_base_name(excel_path, sheet_name, cell_address)
    try:
        reader = PdfReader(str(pdf_path), strict=True)
        if reader.is_encrypted:
            try:
                unlocked = reader.decrypt("")
            except Exception:
                unlocked = 0
            if not unlocked:
                raise ValueError("このPDFはパスワードで保護されています。解除したPDFを選択してください。")
        page_count = len(reader.pages)
        if page_count == 0:
            raise ValueError("PDFにページがありません。")
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f"PDFを読み込めませんでした。破損または未対応の可能性があります。\n{exc}") from exc

    # 先に全ページ分の保存名を作り、既存ファイルを上書きしないよう確認します。
    destinations = [output_dir / f"{base_name}_{page_no:03d}.pdf" for page_no in range(1, page_count + 1)]
    existing = [path.name for path in destinations if path.exists()]
    if existing:
        preview = "\n".join(existing[:8])
        extra = f"\nほか {len(existing) - 8} 件" if len(existing) > 8 else ""
        raise FileExistsError("同名の出力ファイルが既にあります。上書きは行いません。\n" + preview + extra)

    # 途中で保存に失敗しても不完全なファイルを残さないよう、一時フォルダーを使います。
    created: list[Path] = []
    staging_dir = Path(tempfile.mkdtemp(prefix=".pdf_split_", dir=output_dir))
    try:
        for index, page in enumerate(reader.pages, start=1):
            writer = PdfWriter()
            writer.add_page(page)
            staged = staging_dir / f"{base_name}_{index:03d}.pdf"
            with staged.open("xb") as stream:
                writer.write(stream)

        for staged in sorted(staging_dir.glob("*.pdf")):
            destination = output_dir / staged.name
            os.rename(staged, destination)
            created.append(destination)
    except Exception as exc:
        for path in created:
            try:
                path.unlink()
            except OSError:
                pass
        raise ValueError(f"分割ファイルを保存できませんでした。\n{exc}") from exc
    finally:
        shutil.rmtree(staging_dir, ignore_errors=True)

    return destinations


# tkinterを使った画面です。ファイル選択や実行ボタンの処理をまとめています。
class PdfSplitApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("PDF 1ページ分割・セル名＋連番")
        root.resizable(False, False)
        self.pdf_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.cell_var = tk.StringVar(value="A1")

        # 入力欄とボタンを上から順に配置する枠を作ります。
        frame = tk.Frame(root, padx=16, pady=16)
        frame.grid(sticky="nsew")
        self._file_row(frame, 0, "分割するPDF", self.pdf_var, self.choose_pdf)
        self._file_row(frame, 1, "Excelファイル", self.excel_var, self.choose_excel)
        tk.Label(frame, text="シート", width=20, anchor="w").grid(row=2, column=0, sticky="w", pady=5)
        self.sheet_box = ttk.Combobox(frame, textvariable=self.sheet_var, state="readonly", width=53)
        self.sheet_box.grid(row=2, column=1, columnspan=2, sticky="w", padx=6, pady=5)
        tk.Label(frame, text="セル番地", width=20, anchor="w").grid(row=3, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=self.cell_var, width=20).grid(row=3, column=1, sticky="w", padx=6, pady=5)
        self._file_row(frame, 4, "出力先フォルダー", self.output_var, self.choose_output)
        tk.Label(frame, text="選択したセルの値をファイル名に使用します（例: ABC123_001.pdf）。", anchor="w").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(8, 12)
        )
        tk.Button(frame, text="分割を実行", width=18, command=self.run).grid(row=6, column=0, columnspan=3)

    @staticmethod
    def _file_row(frame: tk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        tk.Label(frame, text=label, width=20, anchor="w").grid(row=row, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=variable, width=56).grid(row=row, column=1, padx=6, pady=5)
        tk.Button(frame, text="参照...", command=command).grid(row=row, column=2, pady=5)

    def choose_pdf(self) -> None:
        path = filedialog.askopenfilename(title="分割するPDFを選択", filetypes=[("PDF", "*.pdf")])
        if path:
            self.pdf_var.set(path)
            if not self.output_var.get():
                self.output_var.set(str(Path(path).parent))

    # Excelを選んだらシート名を読み込み、アクティブなシートを初期選択します。
    def choose_excel(self) -> None:
        path = filedialog.askopenfilename(title="Excelファイルを選択", filetypes=[("Excelブック", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                sheet_names = workbook.sheetnames
                active_sheet = workbook.active.title
            finally:
                workbook.close()
        except Exception as exc:
            messagebox.showerror("Excelを読み込めません", f"Excelファイルを開けませんでした。\n{exc}")
            return
        if not sheet_names:
            messagebox.showerror("シートがありません", "選択したExcelファイルにシートがありません。")
            return
        self.excel_var.set(path)
        self.sheet_box["values"] = sheet_names
        self.sheet_var.set(active_sheet if active_sheet in sheet_names else sheet_names[0])

    def choose_output(self) -> None:
        path = filedialog.askdirectory(title="出力先フォルダーを選択")
        if path:
            self.output_var.set(path)

    # 実行ボタンが押されたとき、入力を確認して分割処理を呼び出します。
    def run(self) -> None:
        if not all((self.pdf_var.get(), self.excel_var.get(), self.output_var.get(), self.sheet_var.get(), self.cell_var.get().strip())):
            messagebox.showwarning("入力が必要です", "PDF、Excel、シート、セル番地、出力先フォルダーをすべて指定してください。")
            return
        try:
            results = split_pdf(
                Path(self.pdf_var.get()), Path(self.excel_var.get()), Path(self.output_var.get()),
                self.sheet_var.get(), self.cell_var.get(),
            )
        except Exception as exc:
            messagebox.showerror("分割できませんでした", str(exc))
            return
        messagebox.showinfo("完了", f"{len(results)}ページを分割しました。\n保存先: {results[0].parent}")


# アプリの画面を作り、閉じられるまで待ち受けます。
def main() -> None:
    root = tk.Tk()
    PdfSplitApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
