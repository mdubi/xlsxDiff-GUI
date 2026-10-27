# -*- coding: utf-8 -*-
"""CLI compatibility wrapper for xlsxDiff GUI's enhanced engine.

2026 modifications maintained by Michał Dubrawski; AGPL-3.0-or-later.
Original xlsxDiff by Rafał Czeczótka (2020–2023).
"""
import argparse
from xlsx_diff_engine import DiffOptions, compare_workbooks, parse_index_spec


def main(argv=None):
    p = argparse.ArgumentParser(description='Porównanie dwóch XLSX/XLSM (stara → nowa → raport DIFF.xlsx)')
    p.add_argument('input1', help='starszy plik XLSX lub XLSM')
    p.add_argument('input2', help='nowszy plik XLSX lub XLSM')
    p.add_argument('output', help='raport XLSX')
    p.add_argument('-f', '--formula', action='store_true', help='porównuj treści formuł')
    p.add_argument('-x', '--highlight', action='store_true', help='podświetl nagłówki zmian')
    p.add_argument('-c', '--icolumn', action='append', default=[], help='np. Arkusz!A,B (identyfikator wierszy)')
    p.add_argument('-r', '--irow', action='append', default=[], help='np. Arkusz!1 (identyfikator kolumn)')
    p.add_argument('-X', '--no_highlight_added_removed', action='store_true',
                   help='wyłącz tła wstawionych/usuniętych wierszy/kolumn')
    p.add_argument('-a', '--autofilter', action='store_true', help='autofiltr (wymaga -x)')
    p.add_argument('-e', '--noempty', action='store_true', help='ignoruj komórki puste w obu plikach')
    p.add_argument('-N', '--whole-numbers', action='store_true',
                   help='przekreślaj i podkreślaj CAŁE zmienione liczby (domyślnie wył.)')
    p.add_argument('-v', '--verbose', action='store_true', help='pokazuj postęp')
    p.add_argument('-q', '--quiet', action='store_true', help='ukryj komunikaty')
    p.add_argument('--version', action='version', version='xlsxDiff GUI engine 1.2.1 (based on xlsxDiff by Rafał Czeczótka)')
    args = p.parse_args(argv)
    options = DiffOptions(
        compare_formulas=args.formula,
        highlight_headers=args.highlight,
        index_columns=parse_index_spec(';'.join(args.icolumn), 'column'),
        index_rows=parse_index_spec(';'.join(args.irow), 'row'),
        highlight_added_removed=not args.no_highlight_added_removed,
        autofilter=args.autofilter,
        skip_empty=args.noempty,
        whole_number_diff=args.whole_numbers,
    )
    last_message = None
    def progress(frac, message):
        nonlocal last_message
        if args.quiet:
            return
        if args.verbose or message != last_message and ('arkusz' in message.lower() or message == 'Gotowe'):
            print(f'{100 * frac:.0f}%  {message}', flush=True)
        last_message = message
    try:
        stats = compare_workbooks(args.input1, args.input2, args.output, options,
                                  progress=progress,
                                  warn=lambda msg: print('UWAGA:', msg, flush=True))
    except Exception as exc:
        p.exit(1, f'BŁĄD: {exc}\n')
    if not args.quiet:
        print(f'Gotowe. Komórki zmienione: {stats.changed_cells}, '
              f'arkusze zmienione: {stats.changed_sheets}, '
              f'dodane: {stats.new_sheets}, usunięte: {stats.removed_sheets}')


if __name__ == '__main__':
    main()
