"""
軌道材料管理システム メインモジュール
Track Material Management System - Main Entry Point

Usage:
    python -m src.main --help
    python -m src.main generate --year 2025 --month 4 --output ./output
"""
import argparse
import os
import sys
from datetime import date

from .form7 import generate_form7
from .models import MaterialItem, MaterialTransaction, ProjectInfo, TrackMaterialRegister
from .register import generate_register


def _sample_register() -> TrackMaterialRegister:
    """サンプルデータを含む軌道材料管理簿を返す (for demonstration)"""
    project = ProjectInfo(
        project_name="〇〇線軌道改良工事",
        project_number="R06-軌道-001",
        line_name="〇〇線",
        section="〇〇駅～〇〇駅間",
        contractor="大鉄工業株式会社",
        year=2024,
        month=4,
    )
    register = TrackMaterialRegister(project=project)

    # 材料1: 50Nレール
    rail = MaterialItem(
        material_name="50Nレール",
        specification="50N 25m",
        unit="本",
        previous_balance=10,
    )
    rail.transactions = [
        MaterialTransaction(
            transaction_date=date(2024, 4, 5),
            description="資材センターより受入",
            received_qty=20,
            issued_qty=0,
        ),
        MaterialTransaction(
            transaction_date=date(2024, 4, 10),
            description="軌道補修 第1工区",
            received_qty=0,
            issued_qty=8,
        ),
        MaterialTransaction(
            transaction_date=date(2024, 4, 18),
            description="軌道補修 第2工区",
            received_qty=0,
            issued_qty=12,
        ),
        MaterialTransaction(
            transaction_date=date(2024, 4, 25),
            description="返納（余剰分）",
            received_qty=0,
            issued_qty=5,
            remarks="資材センター返却",
        ),
    ]
    register.add_material(rail)

    # 材料2: 犬釘
    spike = MaterialItem(
        material_name="犬釘",
        specification="25×180mm",
        unit="本",
        previous_balance=500,
    )
    spike.transactions = [
        MaterialTransaction(
            transaction_date=date(2024, 4, 5),
            description="資材センターより受入",
            received_qty=1000,
            issued_qty=0,
        ),
        MaterialTransaction(
            transaction_date=date(2024, 4, 10),
            description="軌道補修 第1工区",
            received_qty=0,
            issued_qty=320,
        ),
        MaterialTransaction(
            transaction_date=date(2024, 4, 18),
            description="軌道補修 第2工区",
            received_qty=0,
            issued_qty=480,
        ),
    ]
    register.add_material(spike)

    # 材料3: タイプレート
    tie_plate = MaterialItem(
        material_name="タイプレート",
        specification="50N用",
        unit="枚",
        previous_balance=200,
    )
    tie_plate.transactions = [
        MaterialTransaction(
            transaction_date=date(2024, 4, 3),
            description="資材センターより受入",
            received_qty=400,
            issued_qty=0,
        ),
        MaterialTransaction(
            transaction_date=date(2024, 4, 10),
            description="軌道補修 第1工区",
            received_qty=0,
            issued_qty=160,
        ),
        MaterialTransaction(
            transaction_date=date(2024, 4, 18),
            description="軌道補修 第2工区",
            received_qty=0,
            issued_qty=240,
        ),
    ]
    register.add_material(tie_plate)

    return register


def cmd_generate(args) -> int:
    """帳票を生成するコマンド"""
    os.makedirs(args.output, exist_ok=True)
    register = _sample_register()

    # 出力ファイルパス
    register_path = os.path.join(
        args.output,
        f"軌道材料管理簿_{register.project.year}年{register.project.month:02d}月.xlsx",
    )
    form7_path = os.path.join(
        args.output,
        f"帳票7_軌道材料受払簿_{register.project.year}年{register.project.month:02d}月.xlsx",
    )

    # 整合性検証
    errors = register.validate_consistency()
    if errors:
        print("【警告】データの整合性エラーが検出されました：", file=sys.stderr)
        for e in errors:
            print(f"  - {e}", file=sys.stderr)
        return 1

    # 帳票生成
    generate_register(register, register_path)
    print(f"軌道材料管理簿を生成しました: {register_path}")

    generate_form7(register, form7_path)
    print(f"帳票7（軌道材料受払簿）を生成しました: {form7_path}")

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="軌道材料管理システム - Track Material Management System"
    )
    subparsers = parser.add_subparsers(dest="command")

    gen_parser = subparsers.add_parser(
        "generate", help="帳票を生成する (Generate forms)"
    )
    gen_parser.add_argument(
        "--output", default="./output",
        help="出力先ディレクトリ (Output directory, default: ./output)"
    )

    args = parser.parse_args()
    if args.command == "generate":
        return cmd_generate(args)

    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
