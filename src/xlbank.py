#!/usr/bin/python3
from datetime import datetime
from dateutil.relativedelta import relativedelta
import itertools
from enum import Enum

from txconsumer import TransactionWorkbookWriter
from txiterator import BankTransactions, CreditTransactions
import argparse


def main() -> None:
    args = parse_args()
    bankfile = BankTransactions(args.bank)
    creditfiles = (CreditTransactions(credit) for credit in args.credit)
    begin = _datetime_from_str(args.begin)
    print(f"beginning {begin}")
    end = (_datetime_from_str(args.end) if args.end else begin) + relativedelta(months=1)
    print(f"ending {end}")
    txfilter = TransactionWorkbookWriter.Filter(begin=begin, end=end, excludebusiness=("כרטיס ויזה",))
    with TransactionWorkbookWriter(outfile=args.out, txfilter=txfilter) as processor:
        processor.accept(itertools.chain(bankfile, *creditfiles))


class _Month(Enum):
    jan = 1
    feb = 2
    mar = 3
    apr = 4
    may = 5
    jun = 6
    jul = 7
    aug = 8
    sep = 9
    oct = 10
    nov = 11
    dec = 12


def _datetime_from_str(val: str) -> datetime:
    error = ValueError("date format is [mmmyy] e.g. feb19")
    if len(val) != 5:
        raise error
    try:
        month = _Month[val[:3]]
    except KeyError:
        raise error
    try:
        year = int(val[3:])
    except ValueError:
        raise error
    return datetime(year=year, month=month.value, day=1)


def parse_args():
    parser = argparse.ArgumentParser(description='Merge current account and credit card transactions and summarize')
    parser.add_argument('--out', type=str, required=True, help='Output file')
    parser.add_argument(
        '--begin', type=str, required=True, help='Starting month (inclusive) from which to produce summary e.g. feb19'
    )
    parser.add_argument(
        '--end',
        type=str,
        required=False,
        help='Ending month (inclusive) for which to produce summary e.g. apr19, defaults to FROM month',
    )
    parser.add_argument('--bank', type=str, required=True, help='Excel file containing current account transactions')
    parser.add_argument('credit', type=str, help='Excel file(s) containing credit card transactions', nargs='+')
    args = parser.parse_args()
    return args


if __name__ == "__main__":
    main()
