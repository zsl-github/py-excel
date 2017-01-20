# -*- coding: utf-8 -*-
import xlrd
import xlwt
import xlutils
import sys
import os
import re
import types
import Levenshtein
from xlwt import *    
from xlrd import open_workbook

from xlrd import open_workbook  
import sys  
from xlutils.copy import copy
import struct
class FileUtils:
    # 支持文件类型 
    # 用16进制字符串的目的是可以知道文件头是多少字节 
    # 各种文件头的长度不一样，少则2字符，长则8字符 
    def __typeList(self): 
        return { 
            "FFD8FF": "JPEG", 
            "47494638":"GIF",
            "49492A00":"TIFF",
            "424D":"Windows Bitmap",
            "41433130":"CAD",
            "38425053":"Adobe Photoshop",
            "7B5C727466":"Rich",
            "3C3F786D6C":"XML",
            "68746D6C3E":"HTML",
            "44656C69766572792D646174653A":"Email",
            "CFAD12FEC5FD746F":"Outlook Express",
            "2142444E":"Outlook",
            "D0CF11E0":"Word-Excel",
            "5374616E64617264204A":"MS Access",
            "89504E47": "PNG"} 

    # 字节码转16进制字符串 
    def __bytes2hex(self, byte): 
        num = len(byte) 
        hexstr = u"" 
        for i in range(num): 
            #转化为十六进制
            t = u"%x" % byte[i] 
            if len(t) % 2: 
                #如果是一位的话，进行补零操作
                hexstr += u"0"
            hexstr += t 
        #转化为大写
        return hexstr.upper() 

    # 获取文件类型 
    def filetype(self, filename): 
        # 必需二制字读取 
        binfile = open(filename, 'rb') 
        tl = self.__typeList() 
        ftype = 'unknown'
        numOfBytes = 100
        # 一个 "B"表示一个字节 
        hbytes = struct.unpack_from('B'*numOfBytes, binfile.read(200)) 
        f_hcode = self.__bytes2hex(hbytes) 
        binfile.close() 
        for hcode in tl.keys(): 
            # 需要读多少字节 
            if f_hcode.startswith(hcode):
                ftype = tl[hcode] 
                return tl[hcode]
        return None

