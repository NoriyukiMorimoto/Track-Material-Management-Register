"""
軌道材料管理システム データモデル
Track Material Management System - Data Models
"""
from dataclasses import dataclass, field
from datetime import date
from typing import List, Optional


@dataclass
class ProjectInfo:
    """工事情報 (Project/Work Information)"""
    project_name: str        # 工事名
    project_number: str      # 工事番号
    line_name: str           # 路線名
    section: str             # 工区・区間
    contractor: str          # 施工業者名
    year: int                # 年度
    month: int               # 月


@dataclass
class MaterialTransaction:
    """材料受払記録 (Material Receipt/Payment Transaction)"""
    transaction_date: date       # 月日
    description: str             # 摘要（使用目的・受入元など）
    received_qty: float          # 受入数量
    issued_qty: float            # 払出数量
    remarks: str = ""            # 備考


@dataclass
class MaterialItem:
    """材料品目 (Material Item)"""
    material_name: str           # 材料品名
    specification: str           # 規格・寸法
    unit: str                    # 単位
    previous_balance: float      # 前月繰越数量
    transactions: List[MaterialTransaction] = field(default_factory=list)

    @property
    def total_received(self) -> float:
        """当月受入合計 (Total received in current month)"""
        return sum(t.received_qty for t in self.transactions)

    @property
    def total_issued(self) -> float:
        """当月払出合計 (Total issued in current month)"""
        return sum(t.issued_qty for t in self.transactions)

    @property
    def current_balance(self) -> float:
        """翌月繰越数量 (Carry-over balance to next month)"""
        return self.previous_balance + self.total_received - self.total_issued

    def get_running_balances(self) -> List[float]:
        """各取引後の残高リストを返す (Returns running balance after each transaction)"""
        balance = self.previous_balance
        balances = []
        for t in self.transactions:
            balance += t.received_qty - t.issued_qty
            balances.append(balance)
        return balances


@dataclass
class TrackMaterialRegister:
    """軌道材料管理簿 (Track Material Management Register)"""
    project: ProjectInfo
    materials: List[MaterialItem] = field(default_factory=list)

    def add_material(self, material: MaterialItem) -> None:
        """材料品目を追加する"""
        self.materials.append(material)

    def validate_consistency(self) -> List[str]:
        """
        データの整合性を検証する (Validate data consistency)
        Returns list of error messages; empty list means all consistent.
        """
        errors = []
        for item in self.materials:
            balance = item.previous_balance
            for i, t in enumerate(item.transactions):
                balance += t.received_qty - t.issued_qty
                if balance < 0:
                    errors.append(
                        f"{item.material_name}: 取引{i + 1}後に残高がマイナスになります "
                        f"(残高: {balance:.3f} {item.unit})"
                    )
        return errors
