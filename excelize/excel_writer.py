import xlwt

BOLD_STYLE = xlwt.easyxf('font: bold on')
DATE_STYLE = xlwt.easyxf(num_format_str='yyyy-mm-dd')


class Book(object):
    def __init__(self, outfile):
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.outfile = outfile
        self.sheets = []

    def add_sheet(self, sheet):
        sheet.book = self
        self.sheets.append(sheet)

    def save(self):
        for sheet in self.sheets:
            sheet.write()
        self.workbook.save(self.outfile)


class Sheet(object):
    def __init__(self, name, rows, title=None, columns=None):
        """
        name: The name of the worksheet.
        rows: And iterable of iterables which constitute the data rows.
        title: A string which is written to cell A1.
        columns: A list of Column objects. These will become column headings.
        """
        for p in ('name', 'rows', 'title', 'columns'):
            setattr(self, p, locals()[p])

    def get_worksheet(self):
        if not hasattr(self, '_worksheet'):
            self._worksheet = self.book.workbook.add_sheet(self.name)

        return self._worksheet
    worksheet = property(get_worksheet)

    def get_next_blank_row(self):
        return max(self.worksheet.rows.keys() or [-1]) + 1
    next_blank_row = property(get_next_blank_row)

    def write_title(self):
        if self.title:
            self.worksheet.write(0, 0, self.title)

    def write_column_headers(self):
        x = 0
        if self.title:
            x = self.next_blank_row + 1
        for y, column in enumerate(self.columns or []):
            self.worksheet.write(x, y, column.name, BOLD_STYLE)

    def write_rows(self):
        for x, row in enumerate(self.rows, self.next_blank_row):
            for y, v in enumerate(row):
                if self.columns[y].is_date:
                    self.worksheet.write(x, y, v, DATE_STYLE)
                else:
                    self.worksheet.write(x, y, v)

    def write(self):
        self.write_title()
        self.write_column_headers()
        if not self.rows:
            return
        self.write_rows()


class Column(object):
    # TODO: It would be better if we input a custom format, rather than the
    # rigid id_date.
    def __init__(self, name, is_date=False):
        self.name = name
        self.is_date = is_date


def quick_columns(*args):
    cols = []
    for a in args:
        if isinstance(a, list) or isinstance(a, tuple):
            cols.append(Column(a[0], a[1]))
        else:
            cols.append(Column(a))

    return cols
