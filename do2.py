#!/usr/bin/env python
# -*- coding: utf-8 -*-
# make.py
# Author: Wang-Yichun

#from tempfile import TemporaryFile
from xlwt import easyxf
from xlrd import open_workbook
from xlutils.copy import copy

rb = open_workbook('source.xls',formatting_info=True)
rs = rb.sheet_by_index(0)
wb = copy(rb)
ws = wb.get_sheet(0)

plain = easyxf('')
for i,cell in enumerate(rs.col(2)):
    if not i:
        continue
    ws.write(i,2,cell.value,plain)

for i,cell in enumerate(rs.col(4)):
	print (i, cell)
	if not i:
		continue
	ws.write(i,4,cell.value-1000)

wb.save('output.xls')
