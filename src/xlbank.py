import itertools
from txconsumer import TransactionWorkbookWriter
from txiterator import BankTransactions, CreditTransactions


def main() -> None:
    bankfile = BankTransactions("/tmp/excel/ca.xlsx")
    creditfile = CreditTransactions("/tmp/excel/ashrai.xlsx")
    outfile = "/tmp/excel/merged.xlsx"
    filters = {"month": 6}
    with TransactionWorkbookWriter(outfile=outfile, filters=filters) as processor:
        processor.accept(itertools.chain(bankfile, creditfile))


if __name__ == "__main__":
    main()
