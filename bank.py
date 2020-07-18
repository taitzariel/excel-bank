import datetime
from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum
from typing import Any, Collection, Iterator, Tuple, Dict, Set, Optional
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


descriptions_by_category: Dict[Category, Set[str]] = {
    Category.mortgage: {
        "משכנתא",
    },
    Category.tax: {
        "מסים",
    },
    Category.running_expenses: {
        "ועד",
        "אינטרנט",
        "חברת החשמל",
        "019",
        "פלאפון",
        "בזק",
    },
    Category.savings: {
        "חסכון",
    },
    Category.donation: {
        "פעמונים",
        "מוסדות חב\"ד",
        "מכון מאיר",
        "עטרת",
        "מה יפו פעמי",
        "התורה והארץ",
        "גרעין יפו",
        "בית דוד בית שמש",
        "המרכז העולמי לחסד",
    },
    Category.insurance: {
        "מכבי",
        "שירותי ברי",
        "ביטוח",
        "פניקס",
        "מגדל",
    },
    Category.education: {
        "אמונה",
    },
    Category.atm: {
        "כספומט"
    },
    Category.mentoring: {
        "שר שלום",
    },
    Category.fuel: {
        "פנגו",
        "פז",
        "כלל חובה",
    },
    Category.food:{
        "מכולת",
        "יינות ביתן",
        "שופרסל",
        "רמי לוי",
    }
}

category_from_description: Dict[str, Category] = {}

# todo: one-liner?
# map(lambda cat, keywords: ((yield cat, keyword) for keyword in keywords), descriptions_by_category.items())
for cat, keywords in descriptions_by_category.items():
    for keyword in keywords:
        category_from_description[keyword] = cat
# category_from_description = {keyword: cat for cat, keywords in descriptions_by_category.items() }


@dataclass
class Transaction:
    amount: Any
    business: str
    charge_date: datetime.datetime
    transaction_date: datetime.datetime

    @property
    def category(self) -> Category:  # todo: better to perform once only
        for kw, category in category_from_description.items():
            if kw in self.business:
                return category
        return Category.other

    details: str
    card: str
    notes: str
    transaction_sum: Any


class Transactions(ABC):

    def __init__(self, filename) -> None:
        self._workbook = load_workbook(filename=filename, read_only=True)
        self._sheet = self._workbook.active  # todo: don't keep unnecessary fields
        self._row_gen = self._sheet.rows
        while self.is_header_row(next(self._row_gen)):
            pass

    @abstractmethod
    def is_header_row(self, row) -> bool:
        return True

    def transaction_generator(self):
        for row in self._row_gen:
            if not row[0].value:
                return
            yield self._convert(row)

    @abstractmethod
    def _convert(self, row) -> Optional[Transaction]:  # todo
        return None


class TransactionWorkbookWriter:

    def __init__(self, outfile: str, filters: dict) -> None:
        self._wb = Workbook()
        self._sheet = self._wb.active
        header_row = "סכום החיוב", "בית עסק", "תאריך עסקה", "טיב", "פירוט", "כרטיס", "הערות", "סכום העסקה"
        self._sheet.append(header_row)  # todo: ensure header row matches data rows
        self._sheet.sheet_view.rightToLeft = True
        column_widths = {'a': 10, 'b': 30, 'c': 20, 'd': 13, 'e': 20, 'f': 10, 'g': 20, 'h': 15}  # todo
        for column, width in column_widths.items():
            self._sheet.column_dimensions[column].width = width
        self._outfile = outfile
        self._filters = filters

    def __enter__(self) -> Any:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._wb.save(self._outfile)

    def process(self, transactions: Iterator[Transaction]) -> None:
        for transaction in transactions:
            if self._relevant(transaction):
                self._sheet.append(self._convert(transaction))

    @staticmethod
    def _convert(transaction: Transaction) -> Tuple:
        return (
            transaction.amount,
            transaction.business,
            transaction.transaction_date,
            transaction.category.value,
            transaction.details,
            transaction.card,
            transaction.notes,
            transaction.transaction_sum,
        )

    def _relevant(self, transaction: Transaction) -> True:
        return (
                self._filters and "month" in self._filters and transaction.charge_date.month == self._filters["month"]
                and
                "כרטיס ויזה" not in transaction.business
        )


class TransactionsMerger:

    def __init__(self, reports: Collection[Transactions]) -> None:
        self._reports = reports

    def merge(self, transactions_processor: Any) -> None:  # todo: Any
        with transactions_processor as processor:
            for report in self._reports:
                processor.process(report.transaction_generator())


class BankTransactions(Transactions):
    def is_header_row(self, row) -> bool:
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


class CreditTransactions(Transactions):
    def is_header_row(self, row) -> bool:
        return row[0].value != "כרטיס"

    def _convert(self, row) -> Transaction:
        return Transaction(
            amount=-row[8].value,
            business=row[1].value,
            charge_date=datetime.datetime.strptime(row[7].value, "%d/%m/%Y"),
            transaction_date=datetime.datetime.strptime(row[2].value, "%d/%m/%Y"),
            details=row[6].value,
            card=row[0].value,
            notes="",
            transaction_sum=row[3].value,
        )


def main() -> None:
    bankfile = BankTransactions("/tmp/excel/ca.xlsx")
    creditfile = CreditTransactions("/tmp/excel/ashrai.xlsx")
    outfile = "/tmp/excel/merged.xlsx"
    filters = {"month": 6}  # todo
    merger = TransactionsMerger(reports=[bankfile, creditfile])
    merger.merge(TransactionWorkbookWriter(outfile=outfile, filters=filters))


if __name__ == "__main__":
    main()
