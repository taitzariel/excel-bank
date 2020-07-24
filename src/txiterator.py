import datetime
from abc import ABC, abstractmethod
from typing import Iterator
from openpyxl import load_workbook
from tx import Transaction


class TransactionIteratable(ABC):

    def __init__(self, filename: str) -> None:
        workbook = load_workbook(filename=filename, read_only=True)
        self._row_gen = workbook.active.rows
        while self._is_header_row(next(self._row_gen)):
            pass

    @abstractmethod
    def _is_header_row(self, row) -> bool:
        """this method should return whether we have reached the row before the first data row"""

    def __iter__(self) -> Iterator[Transaction]:
        for row in self._row_gen:
            if not row[0].value:
                return
            yield self._convert(row)

    @abstractmethod
    def _convert(self, row) -> Transaction:
        """this method should convert a row to a Transaction"""


class BankTransactions(TransactionIteratable):
    def _is_header_row(self, row) -> bool:
        return row[0].value != "תאריך"

    def _convert(self, row) -> Transaction:
        return Transaction(
            amount=-row[3].value,
            business=row[2].value,
            transaction_date=row[0].value,
            charge_date=row[1].value,
            details="",
            card="",
            notes="",
            transaction_sum=None,
        )


class CreditTransactions(TransactionIteratable):
    def _is_header_row(self, row) -> bool:
        return row[0].value != "כרטיס"

    def _convert(self, row) -> Transaction:
        return Transaction(
            amount=row[8].value,
            business=row[1].value,
            charge_date=datetime.datetime.strptime(row[7].value, "%d/%m/%Y"),
            transaction_date=datetime.datetime.strptime(row[2].value, "%d/%m/%Y"),
            details=row[6].value,
            card=row[0].value,
            notes="",
            transaction_sum=row[3].value,
        )
