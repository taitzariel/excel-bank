from abc import ABC
from dataclasses import dataclass
from enum import Enum
from typing import Any, Collection, Iterator, Tuple, Dict
from openpyxl import Workbook, load_workbook


class Category(Enum):
    tax = "מס"
    education = "חינוך"
    food = "אוכל"
    atm = "כספומט"
    running_expenses = "שוטף"
    fuel = "דלק"
    insurance = "ביטוח"
    transport = "תחבורה"
    savings = "חסכון"
    mortgage = "משכנתא"
    donation = "תרומה"
    mentoring = "הדרכה"
    other = "אחר"


category_from_description: Dict[str, Category] = {
    "משכנתא": Category.mortgage,
    "מס" : Category.tax,

}


@dataclass
class Transaction:
    amount: Any
    business: str
    date: Any

    @property
    def category(self) -> Category:  # todo: better to perform once only
        for keyword, category in category_from_description.items():
            if keyword in self.business:
                return category
            else:
                return Category.other


class Transactions(ABC):

    def transaction_generator(self):
        # first_data_row = get_first_data_row()
        # find row where data starts
        # while row is non-empty:
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
        self._sheet = self._wb.active
        header_row = "סכום החיוב", "בית העסק", "תאריך עסקה", "טיב"
        self._sheet.append(header_row)  # todo: ensure header row matches data rows
        self._sheet.sheet_view.rightToLeft = True
        column_widths = {'a': 13, 'b': 33, 'c': 20}  # todo
        for column, width in column_widths.items():
            self._sheet.column_dimensions[column].width = width
        self._outfile = outfile

    def __enter__(self) -> Any:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._wb.save(self._outfile)

    def process(self, transactions: Iterator[Transaction]) -> None:
        for transaction in transactions:
            self._sheet.append(self._convert(transaction))  # todo convert to row

    def _convert(self, transaction: Transaction) -> Tuple:
        return transaction.amount, transaction.business, transaction.date, transaction.category.value


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
        # find row where data starts - done
        for row in self._row_gen:
            if not row[0].value:
                return
            if self._irrelevant_row(row):
                continue
            yield self._convert(row)
        # while row is non-empty:
        #  if row irrelevant:
        #     skip
        #  convert row to Transaction and yield

    def _irrelevant_row(self, row) -> bool:  # todo: abstract
        return False

    def _convert(self, row):
        return Transaction(amount=-row[3].value, business=row[2].value, date=row[0].value)


class CreditTransactions(Transactions):
    def __init__(self, filename) -> None:
        self._filename = filename

    def transaction_generator(self):
        pass


def main() -> None:
    bankfile = BankTransactions("/tmp/excel/ca.xlsx")
    creditfile = CreditTransactions("/tmp/excel/ashrai.xlsx")
    outfile = "/tmp/excel/merged.xlsx"
    merger = TransactionsMerger([bankfile])  # todo , creditfile])
    merger.merge(TransactionWorkbookWriter(outfile=outfile))


if __name__ == "__main__":
    main()
