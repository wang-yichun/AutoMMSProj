#!/usr/bin/env python
# -*- coding: utf-8 -*-
# do.py
# Author: Wang-Yichun

import time

file_name = 'v2015_{para0}.xls'

str = time.strftime('%Y-%m-%d_%H:%M:%S',time.localtime(time.time()))

file_name2 = file_name.replace('{para0}', str)

print file_name2


import os

output_path1 = 'main_backup'
if not os.path.isdir(output_path1):
	os.makedirs(output_path1)
	
output_path2 = 'sub_list'
if not os.path.isdir(output_path2):
	os.makedirs(output_path2)

sourceFile = 'v2015e.xls'
targetFile = 'v2015e_backup.xls'
if not os.path.exists(targetFile):
	open(targetFile, "wb").write(open(sourceFile, "rb").read())
