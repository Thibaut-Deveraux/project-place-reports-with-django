import xlsxwriter

def make_test_xlsx(response):
    book = xlsxwriter.Workbook(response, {'in_memory': True})
    sheet = book.add_worksheet('test')
    sheet.write(0, 0, 'Hello, world!')
    book.close()
    return book

