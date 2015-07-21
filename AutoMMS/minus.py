#!/usr/bin/env python
# -*- coding: utf-8 -*-
# minus.py
# Auto Material Management System Ver1.2 (自动物料管理系统 v1.2)
# Author: Wang-Yichun

import xlrd
import xlwt
import xlutils.copy
import os
import time

title = 'data'

in_xls_name = title + '.xls'
out_xls_name = title + '.xls'
date_str = time.strftime('%Y-%m-%d_%H-%M-%S',time.localtime(time.time()))
out_sub_xls_name = 's' + date_str + '.xls'
out_backup_xls_name = 'b' + date_str + '.xls'

title_row_idx = 2
curr_number_tag = u'[CURR]'
ndiff_number_tag = u'[NDIFF]'
diff_number_tag = u'[DIFF]'

output_path1 = 'main_backup'
output_path2 = 'sub_list'

full_sub_xls_name = output_path2 + '/' + out_sub_xls_name

if not os.path.isdir(output_path1):
	os.makedirs(output_path1)
	
if not os.path.isdir(output_path2):
	os.makedirs(output_path2)

sourceFile = in_xls_name
targetFile = output_path1 + '/' + out_backup_xls_name
if not os.path.exists(targetFile):
	open(targetFile, "wb").write(open(sourceFile, "rb").read())

rbook = xlrd.open_workbook(in_xls_name,formatting_info=False)
wbook = xlutils.copy.copy(rbook)
wbook2 = xlwt.Workbook()


wsheet2 = wbook2.add_sheet('Sheet 1', cell_overwrite_ok=True)

wsheet2_cur_row_idx = 0

update_number = 0

for sheet_index in range(rbook.nsheets):

	rsheet = rbook.sheet_by_index(sheet_index)
	wsheet = wbook.get_sheet(sheet_index)

	wsheet2_cur_row_idx = wsheet2_cur_row_idx + 1
	wsheet2.write(wsheet2_cur_row_idx,0,rsheet.name)
	wsheet2_cur_row_idx = wsheet2_cur_row_idx + 1

	#找数量列 idx
	curr_col_idx = -1
	ndiff_col_idx = -1
	diff_col_idx = -1
	for i in range(rsheet.ncols):
		title_cell_value = rsheet.cell(title_row_idx, i).value
		title_cell_ctype = rsheet.cell(title_row_idx, i).ctype
		
		if title_cell_value != '' and title_cell_ctype == xlrd.XL_CELL_TEXT:
			
			if curr_number_tag in title_cell_value:
				curr_col_idx = i
			if ndiff_number_tag in title_cell_value:
				ndiff_col_idx = i
			if diff_number_tag in title_cell_value:
				diff_col_idx = i

		#复制修改的行标题到子表
		wsheet2.write(wsheet2_cur_row_idx, i, title_cell_value)

	wsheet2_cur_row_idx = wsheet2_cur_row_idx + 1

	if curr_col_idx == -1 or ndiff_col_idx == -1 or diff_col_idx == -1:
		continue;
		
	#处理修改的行
	
	for i in range(title_row_idx + 1, rsheet.nrows):
		curr_cell = rsheet.cell(i, curr_col_idx)
		ndiff_cell = rsheet.cell(i, ndiff_col_idx)
		diff_cell = rsheet.cell(i, diff_col_idx)
		
		need_write = False
		result_value = curr_cell.value
		if ndiff_cell.value != '' and ndiff_cell.ctype == xlrd.XL_CELL_NUMBER:
			result_value = result_value - ndiff_cell.value
			need_write = True
		if diff_cell.value != '' and diff_cell.ctype == xlrd.XL_CELL_NUMBER:
			result_value = result_value + diff_cell.value
			need_write = True
		
		if need_write == True:
			update_number = update_number + 1
			
			print i,curr_col_idx
			wsheet.write(i, curr_col_idx, result_value)
			wsheet.write(i, ndiff_col_idx, '')
			wsheet.write(i, diff_col_idx, '')
			
			#复制修改的行到子表
			for j in range(rsheet.ncols):
				wsheet2.write(wsheet2_cur_row_idx,j,rsheet.cell_value(i,j))
			wsheet2_cur_row_idx = wsheet2_cur_row_idx + 1

wbook.save(out_xls_name)
wbook2.save(full_sub_xls_name)

print 'Operate succeeded! (', update_number ,'updated!)'