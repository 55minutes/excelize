#!/usr/bin/env python

from .._utils import get_production_version
from ..excel_writer import Book, Sheet, quick_columns
from clt_utils import env
from clt_utils.argparse import is_file
from clt_utils.logging import debug
from codecs import open
import argparse
import json
import unicodecsv as csv

__VERSION__ = get_production_version()


def setup_parser():
    """
    Setup the command line utility.
    """
    desc = 'Generate an XLS based on a JSON descriptor and a CSV data source'
    parser = argparse.ArgumentParser(description=desc)
    parser.add_argument(
        'json', metavar='JSON', type=is_file, help='the JSON descriptor')
    parser.add_argument(
        'xls', metavar='XLS', help='the XLS output')
    parser.add_argument(
        '-v', '--version', action='version', version='%(prog)s ' + __VERSION__)
    parser.add_argument(
        '--debug', action='store_true',
        help='show debug messages')
    return parser


def main():
    parser = setup_parser()
    args = parser.parse_args()

    env.DEBUG = args.debug

    book = Book(args.xls, optimized_write=True)

    for sh in json.load(open(args.json, encoding='utf-8')):
        columns = None
        if sh.get('columns', None):
            columns = quick_columns(
                *[(c['name'], c.get('is_date?', False)) for c in sh['columns']]
            )
        debug([(c.name, c.is_date) for c in columns])
        # TODO: Naive rows! We need to take into consideration what the
        # incoming date format is and translate that to a proper datetime
        # object.
        rows = csv.reader(open(sh['data_path']), encoding='utf-8')
        sheet = Sheet(sh['name'], rows, title=sh.get('title', None),
                      columns=columns)
        book.add_sheet(sheet)

    book.save()


if __name__ == '__main__':
    main()
