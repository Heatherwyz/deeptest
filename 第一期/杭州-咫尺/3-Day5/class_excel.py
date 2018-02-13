# -*- coding:utf-8 -*-
__author__ = "Heather"

import os
from openpyxl import Workbook,load_workbook

#excel自定义封装类
class LYMOpenXL:
    def __init__(self,path,read_only=False):
        self.wb = None
        if os.path.exists(path):