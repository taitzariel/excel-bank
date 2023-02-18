import datetime
from abc import ABC, abstractmethod
from typing import Iterator, Optional
from openpyxl import load_workbook
from tx import Transaction


class FormatError(Exception):
    def __init__(self, message: str, filename: Optional[str] = None, row: Optional[int] = None) -> None:
        if filename:
            message += f", filename {filename}"
        if row:
            message += f", row {row}"
        super().__init__(message)


class TransactionIteratable(ABC):

    def __init__(self, filename: str) -> None:
        workbook = load_workbook(filename=filename, read_only=True)
        self._row_gen = workbook.active.rows
        try:
            while self._is_header_row(next(self._row_gen)):
                pass
        except StopIteration:
            raise FormatError(f"failed to find header row", filename=filename)
        self._filename = filename

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
            tid=str(row[5].value).strip(),
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
        business = row[1].value
        transaction_date = row[2].value
        charge_date = row[7].value
        if not charge_date:
            print(
                f"warning: charge date empty, using transaction date instead, {self._str_from_row(row)}"
            )
            charge_date = transaction_date
        amount = row[8].value
        if not isinstance(amount, (float, int)):
            print(f"non-numeral value found for charge sum: {amount}, assuming 0, {self._str_from_row(row)}")
            amount = 0
        elif amount == 0:
            print(f"warning: charge amount empty, {self._str_from_row(row)}")
        transaction_sum = row[3].value
        return Transaction(
            amount=amount,
            business=business,
            charge_date=datetime.datetime.strptime(charge_date, "%d/%m/%Y"),
            transaction_date=datetime.datetime.strptime(transaction_date, "%d/%m/%Y"),
            details=row[6].value,
            card=row[0].value,
            notes="",
            transaction_sum=transaction_sum,
        )

    def _str_from_row(self, row) -> str:
        return str({"filename": self._filename, "business": row[1].value, "transaction_sum": row[3].value})
