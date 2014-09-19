from xlwt import *

fnt = Font()
fnt.name = 'Arial'
fnt.colour_index = 4
fnt.bold = True

borders = Borders()
borders.left = 6
borders.right = 6
borders.top = 6
borders.bottom = 6

al = Alignment()
al.horz = Alignment.HORZ_CENTER
al.vert = Alignment.VERT_CENTER

style = XFStyle()
style.font = fnt
style.borders = borders
style.alignment = al


wb = Workbook()
ws0 = wb.add_sheet('sheet0')
ws1 = wb.add_sheet('sheet1')
ws2 = wb.add_sheet('sheet2')

for i in range(0, 0x200, 2):
    ws0.write_merge(i, i+1, 1, 5, 'test %d' % i, style)
    ws1.write_merge(i, i, 1, 7, 'test %d' % i, style)
    ws2.write_merge(i, i+1, 1, 7 + (i%10), 'test %d' % i, style)


wb.save('merged.xls')

from xlwt import *

wb = Workbook()
ws0 = wb.add_sheet('sheet0')


fnt = Font()
fnt.name = 'Arial'
fnt.colour_index = 4
fnt.bold = True

borders = Borders()
borders.left = 6
borders.right = 6
borders.top = 6
borders.bottom = 6

style = XFStyle()
style.font = fnt
style.borders = borders

ws0.write_merge(3, 3, 1, 5, 'test1', style)
ws0.write_merge(4, 10, 1, 5, 'test2', style)
ws0.col(1).width = 0x0d00

wb.save('merged0.xls')

from xlwt import *

wb = Workbook()
ws0 = wb.add_sheet('sheet0')

fnt1 = Font()
fnt1.name = 'Verdana'
fnt1.bold = True
fnt1.height = 18*0x14

pat1 = Pattern()
pat1.pattern = Pattern.SOLID_PATTERN
pat1.pattern_fore_colour = 0x16

brd1 = Borders()
brd1.left = 0x06
brd1.right = 0x06
brd1.top = 0x06
brd1.bottom = 0x06

fnt2 = Font()
fnt2.name = 'Verdana'
fnt2.bold = True
fnt2.height = 14*0x14

brd2 = Borders()
brd2.left = 0x01
brd2.right = 0x01
brd2.top = 0x01
brd2.bottom = 0x01

pat2 = Pattern()
pat2.pattern = Pattern.SOLID_PATTERN
pat2.pattern_fore_colour = 0x01F

fnt3 = Font()
fnt3.name = 'Verdana'
fnt3.bold = True
fnt3.italic = True
fnt3.height = 12*0x14

brd3 = Borders()
brd3.left = 0x07
brd3.right = 0x07
brd3.top = 0x07
brd3.bottom = 0x07

fnt4 = Font()

al1 = Alignment()
al1.horz = Alignment.HORZ_CENTER
al1.vert = Alignment.VERT_CENTER

al2 = Alignment()
al2.horz = Alignment.HORZ_RIGHT
al2.vert = Alignment.VERT_CENTER

al3 = Alignment()
al3.horz = Alignment.HORZ_LEFT
al3.vert = Alignment.VERT_CENTER

style1 = XFStyle()
style1.font = fnt1
style1.alignment = al1
style1.pattern = pat1
style1.borders = brd1

style2 = XFStyle()
style2.font = fnt2
style2.alignment = al1
style2.pattern = pat2
style2.borders = brd2

style3 = XFStyle()
style3.font = fnt3
style3.alignment = al1
style3.pattern = pat2
style3.borders = brd3

price_style = XFStyle()
price_style.font = fnt4
price_style.alignment = al2
price_style.borders = brd3
price_style.num_format_str = '_(#,##0.00_) "money"'

ware_style = XFStyle()
ware_style.font = fnt4
ware_style.alignment = al3
ware_style.borders = brd3


ws0.merge(3, 3, 1, 5, style1)
ws0.merge(4, 10, 1, 6, style2)
ws0.merge(14, 16, 1, 7, style3)
ws0.col(1).width = 0x0d00


wb.save('merged1.xls')
