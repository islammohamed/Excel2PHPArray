from openpyxl import load_workbook
import glob
import sys
import codecs
reload(sys);
sys.setdefaultencoding("utf8")

def loadSheetFile(path):
        sheet = load_workbook(path)
        return sheet.active

def openWriteableFile(name):
        file = codecs.open(name, 'w+', "utf8")
        file.write(u'\ufeff')
        return file


def removeBaseUrl(haystack, base_url = []):
        if not haystack:
                return ''

        for url in base_url:
                haystack = haystack.replace(url, '')
        return haystack

def exportSheet(sheet, name, output_file):
        rows = sheet.rows
        rest_of_rows = rows[1:]
        for row in rest_of_rows:
                old_url, new_url = removeBaseUrl(row[0].value), removeBaseUrl(row[1].value)
                if new_url in ['#N/A', 'N/A', None]:
                        new_url = ''
                value  = u"\n    \"%s\" => \"%s\", " % (old_url, new_url)
                value = value.encode('utf8')
                print value
                output_file.write(value)

def getExcelSheets():
        return glob.glob('data/*.xlsx')

if __name__ == '__main__':
        output_file = openWriteableFile('array.php')
        output_file.write("<?php\n\n$urlList = [");
        for sheet in getExcelSheets():
                sheetFile = loadSheetFile(sheet)
                print sheetFile.title
                exportSheet(sheetFile, sheetFile.title, output_file)

        output_file.write('\n];');
