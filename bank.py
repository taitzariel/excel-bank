from dataclasses import dataclass
from typing import Any
from openpyxl import Workbook


@dataclass
class Transaction:
    amount: Any
    business: str


class TransactionsMerger:

    def __init__(self, bankfile: str, creditfile: str) -> None:
        self._creditfile = creditfile
        self._bankfile = bankfile

    def generate_wb(self, outfile: str) -> None:
        pass


def main() -> None:
    bankfile = "~/tmp/excel/ca.xlsx"
    creditfile = "~/tmp/excel/ashrai.xlsx"
    outfile = "~/tmp/excel/merged.xlsx"
    merger = TransactionsMerger(bankfile=bankfile, creditfile=creditfile)
    merger.generate_wb(outfile=outfile)
