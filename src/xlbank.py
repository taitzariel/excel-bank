#!/usr/bin/python3

import itertools
from txconsumer import TransactionWorkbookWriter
from txiterator import BankTransactions, CreditTransactions
import argparse


def main() -> None:
    args = parse_args()
    bankfile = BankTransactions(args.bank)
    creditfiles = (CreditTransactions(credit) for credit in args.credit)
    txfilter = TransactionWorkbookWriter.Filter(month=args.month, excludebusiness=("כרטיס ויזה",))
    with TransactionWorkbookWriter(outfile=args.out, txfilter=txfilter) as processor:
        processor.accept(itertools.chain(bankfile, *creditfiles))


def parse_args():
    parser = argparse.ArgumentParser(description='Merge current account and credit card transactions and summarize')
    parser.add_argument('month', type=int, help='Month for which to produce summary')
    parser.add_argument('bank', type=str, help='Excel file containing current account transactions')
    parser.add_argument('credit', type=str, help='Excel file(s) containing credit card transactions', nargs='+')
    parser.add_argument('out', type=str, help='Output file')
    args = parser.parse_args()
    return args


if __name__ == "__main__":
    main()
