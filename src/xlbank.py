import itertools
from txconsumer import TransactionWorkbookWriter
from txiterator import BankTransactions, CreditTransactions
import argparse


def main() -> None:
    args = parse_args()
    bankfile = BankTransactions(args.bank)
    creditfile = CreditTransactions(args.credit)
    filters = {"month": args.month}
    with TransactionWorkbookWriter(outfile=args.out, filters=filters) as processor:
        processor.accept(itertools.chain(bankfile, creditfile))


def parse_args():
    parser = argparse.ArgumentParser(description='Merge current account and credit card transactions and summarize')
    parser.add_argument('month', type=int, help='Month for which to produce summary')
    parser.add_argument('bank', type=str, help='Excel file containing current account transactions')
    parser.add_argument('credit', type=str, help='Excel file containing credit card transactions')
    parser.add_argument('out', type=str, help='Output file')
    args = parser.parse_args()
    return args


if __name__ == "__main__":
    main()
