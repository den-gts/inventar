# -*-coding: utf-8 -*
import logging
from reportlab.lib.units import mm
from reportlab.lib import colors, pagesizes
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle, Flowable, Paragraph, BaseDocTemplate, PageTemplate, Frame
from reportlab.lib.styles import getSampleStyleSheet
import xlrd, datetime, os.path
from reportlab.lib.textsplit import getCharWidths
import pyphen, argparse, locale, sys

# Настройка логгирования.
log = logging.getLogger('main')
log.setLevel(logging.DEBUG)
logHandlerFile = logging.FileHandler('log.txt', mode='w')
logHandlerFile.setLevel(logging.DEBUG)
logHandlerStream = logging.StreamHandler()
logHandlerStream.setLevel(logging.DEBUG)
logFormater = logging.Formatter('%(asctime)s  [%(levelname)s]: %(message)s')
logHandlerStream.setFormatter(logFormater)
logHandlerFile.setFormatter(logFormater)
log.addHandler(logHandlerFile)
log.addHandler(logHandlerStream)



# регистрация шрифтов
pdfmetrics.registerFont(TTFont('gost', 'arial.TTF'))
pdfmetrics.registerFont(TTFont('slim', 'ARIALN.TTF'))
# а так же стиля
styleSheet = getSampleStyleSheet()['Normal']
styleSheet.fontName = 'slim' # имя шрифта
styleSheet.fontSize = 12  # размер шрифта
styleSheet.leading = 8  # межстрочный интервал
rowHeight = 8*mm  #  высота строки таблицы
rowCount = 33  # количество строк на странице
Hoffset = 4*mm  # отступ
columnWidths = (17*mm, 15*mm, 41*mm, 7*mm, 9*mm, 53*mm, 10*mm, 15*mm, 19*mm) # ширина колонок в файле PDF
tabStyle = TableStyle([  # стил таблицы в PDF
                           ('GRID', (0, 0), (-1, -1), 0.4*mm, colors.black), 
                           ('FONT', (0, 0), (-1, -1), 'slim', 12),
                           #('FONT', (4, 1), (4, -1), 'slim', 11),
                           ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                           ('ALIGN', (2, 0), (2, -1), 'LEFT'),
                           ('LEFTPADDING', (2, 0), (2, -1), 1),
                           ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                           ])


def tmplPage(canvas, doc):  # шаблон страницы
    canvas.saveState()

    class rotateText(Flowable): # определение класса поворачиваемого текста на 90 градусов
        def __init__(self, text):
            self.text = text
            self.parag = Paragraph(self.text, styleSheet)

        def wrap(self, availWidth, availHeight):
            return -availWidth, availHeight

        def draw(self):
            self.canv.rotate(90)
            self.parag.wrapOn(self.canv, self._fixedHeight, self.parag._fixedWidth)
            self.parag.drawOn(self.canv, 0, 0)

    dataTable = [[u'Инвен-\nтарный\nномер', u'Дата', u'Обозначение', rotateText(u'Кол. листов'), # заголовок таблицы
               rotateText(u'Формат'), u'Наименование', u'Кем\nвыпу-\nщен',
               Paragraph(u'Подпись\nо\nприем-\nке до-\nкумента', styleSheet), u'При-\nмечание']]

    t = Table(dataTable, columnWidths, 18*mm) # формирование таблицы
    t.setStyle(tabStyle) # присвоение таблице стиля
    t.wrapOn(canvas, *pagesizes.A4)
    t.drawOn(canvas, 20*mm, pagesizes.A4[1] - (pagesizes.A4[1] - rowHeight*rowCount - t._height)/2 - t._height + Hoffset) # отрисовка таблицы
    canvas.drawString(pagesizes.A4[0] - 10*mm, 5*mm, str(doc.page)) # отрисовка номера страницы
    canvas.restoreState()


def parseWorkSheet(sheet):  # разбор екселевского листа
    data = []
    for rownum in xrange(1, sheet.nrows):  # цикл строк
        row = sheet.row_values(rownum)  # значение ячеек в строке
        if not row[0]:  # если дата пуста то выходим из функции
            break
        if not row[2]:  # если инвентарный номер пуст пропускаем строку, пробелы обрезаюся слева и справа при проверке
            continue
        log.info("Обработка строчки %s %s" % (row[3].encode('utf-8'), row[7].encode('utf-8')))
        try:
            row[0] = datetime.date(*xlrd.xldate_as_tuple(row[0], 0)[:3]).strftime('%d.%m.%y')  # форматирование даты
        except ValueError as er:
            log.error('Ошибка в колонке дата(%s) в строке номер %d' % (er, rownum + 1))
            sys.exit()
        formatIndex = row[5].find('(', 0)  # обработка форматов документа.
        if formatIndex > 0:				  # если документ выполнен в разных форматах
            row[5] = row[5][:formatIndex]  # убираем скобки
        row[4] = int(row[4])
        numberCol = (2, 0, 3, 4, 5, 7, 17) # порядковый номер колонок в екселевском файле
        dataRow = [row[x] for x in numberCol] # строка с данными. оставляем только нужные нам колонки.

        if sheet.name.lower() == u'аннулированные':  # если лист называется "аннулированные"
            dataRow.extend(("", u"Аннулир."))  # r к строке с данными добавляем в примичание "Аннулир."
        else:
            dataRow.extend(("", ""))
        data.append(dataRow)
    return data


def parseXLS(xlsFile):  # разбор екселевского файла
    workbook = xlrd.open_workbook(xlsFile)  # открытие екселевского файла
    data1 = parseWorkSheet(workbook.sheet_by_index(0))
    # Парсинг листа "анулированные"
    data2 = parseWorkSheet(workbook.sheet_by_index(1))
    data1.extend(data2)
    try:
        data1 = sorted(data1, key=lambda col: col[0])  # сортировка данных по инвентарному номеру (колонка 0)
    except ValueError:
        log.error(u"ошибка в поле инв.номер")
        sys.exit(1)
    return data1


def softWarpString(text, width): # функция переноса строки
    dic = pyphen.Pyphen(lang='ru-RU')
    charWidths = getCharWidths(text, 'gost', 12)
    hypPositions = dic.positions(text)
    rowPosition = []
    rowWidth = 0
    for charIndex in xrange(0, len(charWidths)):
        rowWidth += charWidths[charIndex]
        if rowWidth >= width:
            rowPosition.append(charIndex)
            rowWidth = 0
    prv = 0
    result = []
    for r in rowPosition:
        sufix = ""
        lessH = filter(lambda h: h < r, hypPositions)
        r = lessH[-1]
        if " " not in (text[r-1], text[r+1]):
            sufix = "-"
        result.append(text[prv:r] + sufix)
        prv = r
    result.append(text[prv:])
    return result


def calcWarps(data, descrWidth): # вычислить перенос строки
    result = []
    for row in data:
        rowDesc = softWarpString(row[5], descrWidth)
        if len(rowDesc) > 1:
            row[5] = rowDesc[0]
            result.append(row)
            for r in rowDesc[1:]:
                result.append(("", "", "", "", "", r, "", "", ""))
        else:
            result.append(row)
    return result

# выполнение программы начинается отсюда
# обработка аргументов коммандной строки.
parser = argparse.ArgumentParser()
parser.add_argument('input', help='входной файл XLS')
parser.add_argument('output', help='выходной файл PDF', nargs='?')
opt = parser.parse_args(sys.argv[1:])
xlsFile = opt.input.decode(locale.getpreferredencoding())
xlsFile = os.path.normpath(xlsFile)
if not os.path.exists(xlsFile):
    log.error(u'Файл %s не найден' % xlsFile)
    sys.exit(1)
if not opt.output:
    outputFile = os.path.splitext(xlsFile)[0].encode(locale.getpreferredencoding()) + ".pdf"
else:
    outputFile = os.path.normpath(opt.output).encode(locale.getpreferredencoding())
try:
    data = calcWarps(parseXLS(xlsFile), columnWidths[5] + 2*mm)  # вычисление таблицы данных
except IOError:
    log.error(u'ошибка при открытии файла %s' % xlsFile)
    sys.exit(1)
except xlrd.biffh.XLRDError:
    log.error(u'не правильный формат файла %s' % xlsFile)
    sys.exit(1)
except ValueError as err:
    log.error(u'ошибка при разборе XLS файла(%s)' % err)
    sys.exit(1)

while len(data) % rowCount:  # заполняем пустыми строками до конца страницы
    data.append(9*[''])
table = Table(data, columnWidths, 8*mm)
table.setStyle(tabStyle)
content = []
frame = Frame(20*mm, (pagesizes.A4[1] - rowHeight*rowCount - 18*mm)/2 + Hoffset,
              reduce(lambda x, y: x + y, columnWidths), rowCount*rowHeight,
              leftPadding=0, rightPadding=0, bottomPadding=0, topPadding=0, showBoundary=1)
doc = BaseDocTemplate(outputFile.decode(locale.getpreferredencoding()),
                      pagesize=pagesizes.A4,
                      leftMargin=25*mm,
                      rightMargin=5*mm,
                      topMargin=5*mm,
                      bottomMargin=5*mm,
                      pageTemplates=[PageTemplate(onPage=tmplPage, frames=[frame])])
content.append(table)
try:
    doc.build(content)
except Exception as err:
    log.error(u'Ошибка построения PDF файла %s' % (err))
    sys.exit(1)