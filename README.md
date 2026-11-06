# PDF 1ページ分割・セル名＋連番ツール

PDFを1ページずつに分け、Excelで指定したセルの値とページ番号を組み合わせた名前で保存するWindows向けPythonツールです。

## この版のファイル名

選択したセルに `ABC123` と入力され、PDFが5ページの場合、次の5ファイルを作成します。

```text
ABC123_001.pdf
ABC123_002.pdf
ABC123_003.pdf
ABC123_004.pdf
ABC123_005.pdf
```

セルの値をそのままファイル名にする版は [PDF-Split-Rename-By-Cell-Public](https://github.com/K5NISSHY/PDF-Split-Rename-By-Cell-Public) をご覧ください。

## フォルダー内のファイル

- `pdf_splitter.py`: 分割処理と画面
- `起動.bat`: ダブルクリックで起動
- `README_使い方.txt`: 日本語の簡易説明
- `サンプル.xlsx`: ファイル名の例を入れたExcelブック
- `サンプル.pdf`: 動作確認用の5ページPDF

## 動作環境

- Windows 11
- Python 3.10以降
- Excel形式のブック（`.xlsx` または `.xlsm`）

`起動.bat` はPythonと必要なライブラリを確認し、不足時は `pypdf` と `openpyxl` をインストールします。初回のインストール時にはインターネット接続が必要です。Pythonのインストール時は「Add python.exe to PATH」を有効にしてください。

### 使用ライブラリ

- [pypdf](https://pypi.org/project/pypdf/): PDFを読み込み、ページ単位に書き出します。
- [openpyxl](https://openpyxl.readthedocs.io/): Excelブックからシートとセルの値を読み込みます。
- `tkinter`: 画面を作るPython標準ライブラリです。

## 使い方

1. `起動.bat` をダブルクリックします。
2. 「分割するPDF」でPDFを選びます。
3. 「Excelファイル」でExcelブックを選びます。
4. シートを選び、ファイル名を読むセル番地（例: `A1`、`B3`）を入力します。
5. 出力先フォルダーを選びます。PDF選択後は、元PDFと同じフォルダーが初期値になります。
6. 「分割を実行」を押します。

## 安全上の動作

- 元PDFは読み取りのみで、変更しません。
- 出力ファイルが既にある場合は、上書きせず処理を中止します。
- セルが空欄、ファイル名に使えない文字、Windowsの予約名などを検出するとエラーを表示します。
- 数式セルでは数式を計算せず、Excelに保存された計算結果を使います。結果が空または未保存の場合はExcelで再計算して保存してください。
- 空のパスワードで開けない暗号化PDFには対応しません。

## 関連記事

Excelのセル値をそのままファイル名にする別版は [PDFを1ページずつ分割して、Excelのセル値でファイル名を付けるPythonツール](https://speeddemon.jp/?p=2061) をご覧ください。

## ライセンスと著作権

MIT Licenseです。著作権表示は `K5NISSHY / speeddemon.jp` です。正式な英語本文は [LICENSE](LICENSE)、日本語参考訳は [LICENSE.ja.md](LICENSE.ja.md) をご確認ください。
