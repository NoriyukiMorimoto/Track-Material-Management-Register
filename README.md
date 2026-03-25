# Track-Material-Management-Register
軌道材料管理簿の改良

軌道材料管理簿（Track Material Management Register）と帳票7（軌道材料受払簿）を整合させて生成するPythonシステムです。

---

## 概要

| 帳票 | 内容 |
|------|------|
| **軌道材料管理簿** | 工事全体の材料品目ごとに、前月繰越・当月受入・当月払出・翌月繰越数量を一覧管理する台帳 |
| **帳票7（軌道材料受払簿）** | 材料品目ごとの日次受払明細。取引日・摘要・受入/払出数量・残高の時系列記録 |

2種類の帳票は同一データソース（`TrackMaterialRegister`）から生成されるため、集計値（当月受入合計・当月払出合計・翌月繰越）が常に一致します。

---

## ディレクトリ構成

```
Track-Material-Management-Register/
├── src/
│   ├── models.py       # データモデル（ProjectInfo, MaterialItem, MaterialTransaction など）
│   ├── register.py     # 軌道材料管理簿 Excel 生成
│   ├── form7.py        # 帳票7（軌道材料受払簿）Excel 生成
│   └── main.py         # CLIエントリポイント
├── tests/
│   └── test_forms.py   # ユニットテスト（21件）
├── requirements.txt
└── README.md
```

---

## セットアップ

```bash
pip install -r requirements.txt
```

---

## 使い方

### サンプル帳票を生成する

```bash
python -m src.main generate --output ./output
```

`./output/` に以下の2ファイルが生成されます。

- `軌道材料管理簿_YYYY年MM月.xlsx`
- `帳票7_軌道材料受払簿_YYYY年MM月.xlsx`

### コードから利用する

```python
from datetime import date
from src.models import ProjectInfo, MaterialItem, MaterialTransaction, TrackMaterialRegister
from src.register import generate_register
from src.form7 import generate_form7

# 1. プロジェクト情報を設定
project = ProjectInfo(
    project_name="〇〇線軌道改良工事",
    project_number="R06-軌道-001",
    line_name="〇〇線",
    section="〇〇駅～〇〇駅間",
    contractor="施工業者名",
    year=2024,
    month=4,
)

# 2. 材料データを追加
register = TrackMaterialRegister(project=project)
rail = MaterialItem(material_name="50Nレール", specification="50N 25m",
                    unit="本", previous_balance=10)
rail.transactions = [
    MaterialTransaction(date(2024, 4, 5), "資材センターより受入", received_qty=20, issued_qty=0),
    MaterialTransaction(date(2024, 4, 10), "軌道補修", received_qty=0, issued_qty=15),
]
register.add_material(rail)

# 3. 整合性を検証してから帳票を生成
errors = register.validate_consistency()
if not errors:
    generate_register(register, "軌道材料管理簿.xlsx")
    generate_form7(register, "帳票7_軌道材料受払簿.xlsx")
```

---

## テスト実行

```bash
python -m pytest tests/ -v
```

---

## 帳票の整合性

軌道材料管理簿と帳票7は以下の関係を保ちます。

```
帳票7（材料Xの受払明細）の合計行
  受入数量合計 ── = ──▶ 軌道材料管理簿 材料X行 「当月受入」
  払出数量合計 ── = ──▶ 軌道材料管理簿 材料X行 「当月払出」
  翌月繰越残高 ── = ──▶ 軌道材料管理簿 材料X行 「翌月繰越」
```

データの不整合（残高マイナス等）は `validate_consistency()` で事前に検出できます。
