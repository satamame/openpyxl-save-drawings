# openpyxl で保存した Workbook の画像を復元する

[openpyxl](https://openpyxl.readthedocs.io/) でワークブックを保存すると、図形や画像が保存されないという制限があります。
つまり、既存の Excel ブックを [Workbook](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.workbook.workbook.html) として開いて、セルの値をちょっと変えて上書き保存、という使い方をしたい場合に、図形や画像が失われてしまう、ということです。

## 解決策の概要

保存する前と保存後の Excel ブックを、それぞれ一時フォルダに zip 解凍します。  
保存前のブックからデータを復元し、最後に保存後のフォルダを zip 圧縮します。

この解決策では、以下のものを復元します。

1. 図形や画像
2. コメントの書式
3. データの入力規則

※復元されないものもあるかもしれません。

## デモ

Windows でデモを見る手順です。

### openpyxl でそのまま保存する

1. このリポジトリをクローンして、requirements.txt を使って仮想環境を作る。
    ```
    > python -m venv .venv --upgrade-deps
    > .venv/Scrips/activate
    (.venv)> pip install -r requirements.txt
    ```
1. openpyxl で a.xlsx を開いて b.xlsx として保存するには:
    ```
    (.venv)> python app.py a.xlsx b.xlsx --just-save
    ```
    a.xlsx (保存前)  
    ![](img/a-xlsx.png)

    b.xlsx (保存後)  
    ![](img/just-save.png)

### 解決策を使って保存する

1. openpyxl で a.xlsx を開いて変更を加えて保存し、画像を復元するには:
    ```
    (.venv)> python app.py a.xlsx b.xlsx
    ```
    b.xlsx (B1 セルを変更して保存し、画像を復元)  
    ![](img/restore-drawings.png)

## 使い方

openpyxl を使っているプロジェクトに、save_with_drawings.py というモジュールをコピーして使えます。

```python
from pathlib import Path
from openpyxl import load_workbook
from save_with_drawings import save_with_drawings

src = Path('a.xlsx')
wb = load_workbook(src)

# ここで何らかの変更を加える。

dest = Path('b.xlsx')
temp_dir_args = {'prefix': 'temp_', 'dir': '.'}

save_with_drawings(wb, src, dest, temp_dir_args)
```

※`save_with_drawings()` の引数はソースコードの docstring を見てください。
