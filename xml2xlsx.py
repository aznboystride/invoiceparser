import os
import xml.etree.ElementTree as et
import openpyxl


def xml2xlsx(xmlfile, savepath):
	tree = et.parse(xmlfile)
	root = tree.getroot()
	rows = root[0][0]
	wb = openpyxl.load_workbook(savepath)
	sheet = wb['Sheet 1']
	sheet.delete_rows(1, 300)
	r = 1
	for row in rows:
		c = 1
		for cell in row:
			sheet.cell(row=r, column=c).value = cell[0].text
			c += 1
		r += 1
	wb.save(savepath)

def main():
    xml2xlsx("C:\\Users\\peynu\\Documents\\redline\\recent_invoice.xls", "C:\\Users\\peynu\\Documents\\redline\\test.xlsx")

if __name__ == '__main__':
	main()
