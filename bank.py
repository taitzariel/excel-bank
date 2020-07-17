from dataclasses import dataclass
from typing import Any, Collection, Generator, Iterator
from openpyxl import Workbook, load_workbook
from abc import ABC, abstractmethod
import contextlib


@dataclass
class Transaction:
    amount: Any
    business: str


class Transactions(ABC):

    def transaction_generator(self):
        first_data_row = get_first_data_row()
        #find row where data starts
        #while row is non-empty:
        #  if row irrelevant:
        #     skip
        #  convert row to Transaction and yield
        pass

    # @abstractmethod
    # def get_first_data_row(self) -> int:
    #     return -3


class TransactionWorkbookWriter:

    def __init__(self, outfile: str) -> None:
        self._wb = Workbook()
        self._outfile = outfile

    def __enter__(self) -> Any:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._wb.save(self._outfile)

    def process(self, transactions: Iterator[Transaction]) -> None:
        sheet = self._wb.active
        for transaction in transactions:
            sheet.append(transaction)  # todo convert to row


class TransactionsMerger:

    def __init__(self, reports: Collection[Transactions]) -> None:
        self._reports = reports

    def merge(self, transactions_processor: Any) -> None:  # todo: Any
        with transactions_processor as processor:
            for report in self._reports:
                processor.process(report.transaction_generator())


class BankTransactions(Transactions):

    def __init__(self, filename) -> None:
        self._workbook = load_workbook(filename=filename, read_only=True)
        self._sheet = self._workbook.active  # todo: don't keep unnecessary fields
        self._row_gen = self._sheet.rows
        while next(self._row_gen)[0].value != "תאריך":
            pass

    def transaction_generator(self):
        #find row where data starts - done
        for row in self._row_gen:
            if not row[0].value:
                return
            if self._irrelevant_row(row):
                continue
            yield self._convert(row)
        #while row is non-empty:
        #  if row irrelevant:
        #     skip
        #  convert row to Transaction and yield

    def _irrelevant_row(self, row) -> bool:
        return False

    def _convert(self, row):
        return Transaction(amount="23", business="petrol")


class CreditTransactions(Transactions):
    def __init__(self, filename) -> None:
        self._filename = filename

    def transaction_generator(self):
        pass


def main() -> None:
    bankfile = BankTransactions("/tmp/excel/ca.xlsx")
    creditfile = CreditTransactions("/tmp/excel/ashrai.xlsx")
    outfile = "/tmp/excel/merged.xlsx"
    merger = TransactionsMerger([bankfile, creditfile])
    merger.merge(TransactionWorkbookWriter(outfile=outfile))


main()
