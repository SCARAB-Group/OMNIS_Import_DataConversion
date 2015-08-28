#! /opt/local/bin/python
 
import odict
import os
import xlrd
import xlwt
import math
import datetime
import copy
 
import functools
import operator
 
import unicodedata 
 
#import matplotlib
#matplotlib.use('TkAgg')
 
#import numpy as np
#import matplotlib.pyplot as plt
 
#paths
#data_path1 = '../../141021/'
#data_path_out1 = '../../141021/Converted_data/'

#data_path1 = '../../141112/Import_files_to_convert/'
#data_path_out1 = '../../141112/Import_files_to_convert/Converted_141118_2/'

#Files
#Orig_order_file = 'Solna-hist-prep-table-IMPORT_01.xls'
#datafile_out = 'Solna-hist-prep-table_IMPORT_01_converted.xls'

#File 1
#Orig_order_file = 'Solna-hist-prep-table_WORK_FILE.xls'
#datafile_out = 'Solna-hist-prep-table_WORK_FILE.xls_converted141203.xls'


##############################################
# Conversion 141203 
#data_path1 = '../../141201/'
#data_path_out1 = '../../141201/ConvertedFiles_141203/'
#File 
#Orig_order_file = 'Solna-hist-prep-table_WORK_FILE_DirektUttag.xls'
#datafile_out = 'Solna-hist-prep-table_WORK_FILE_DirektUttag_converted141203.xls'


##############################################
# Conversion 141219
#data_path1 = '../../141219/'
#data_path_out1 = '../../141219/'
#File 
#Orig_order_file = 'Solna-hist-1201-1209_WORK_FILE.xls'
#datafile_out = 'Solna-hist-1201-1209_WORK_FILE_Converted_141219.xls'


##############################################
# Conversion 150223
#data_path1 = '../../150223/'
#data_path_out1 = '../../150223/'
#File 
#Orig_order_file = 'Huddinge-hist-facs-prep-table_JOSTAP.xls'
#datafile_out = 'Huddinge-hist-facs-prep-table_JOSTAP_Converted_150223.xls'


################################################
### Conversion 150305
##data_path1 = '../../150305/'
##data_path_out1 = '../../150305/'
###File 
##Orig_order_file = 'Huddinge_OMNIS.xls'
##datafile_out = 'Huddinge_OMNIS_Converted_150305.xls'

##############################################
# Conversion 150331
#data_path1 = '../../150331/'
#data_path_out1 = '../../150331/'
#File 
#Orig_order_file = 'Solna-endokrin_150331.xls'
#datafile_out = 'Solna-endokrin_Converted_150331.xls'


##############################################
# Conversion 150414
#data_path1 = '../../150414/'
#data_path_out1 = '../../150414/'
#File 
#Orig_order_file = 'Huddinge-peri-njur-neuro-prep-table_2015-04-13.xls'
#datafile_out = 'Huddinge-peri-njur-neuro-prep-table_Converted_2015-04-13.xls'


##############################################
# Conversion 150423
#data_path1 = '../../150423/'
#data_path_out1 = '../../150423/'
#File 
#Orig_order_file = 'Huddinge_PERI_NJUR_NEURO_150422.xls'
#datafile_out = 'Huddinge_PERI_NJUR_NEURO_Converted_150422.xls'

##############################################
# Conversion 150520
data_path1 = '../../150520/'
data_path_out1 = '../../150520/'
#File 
Orig_order_file = 'SOLNA_ENDO_IMPORT_V2.xls'
datafile_out = 'SOLNA_ENDO_IMPORT_V2_Converted_150520.xls'


#Functions
from itertools import izip, imap
from copy import deepcopy
 
missing = object()
 
 
class odict(dict):
 
    def __init__(self, *args, **kwargs):
        dict.__init__(self)
        self._keys = []
        self.update(*args, **kwargs)
 
    def __delitem__(self, key):
        dict.__delitem__(self, key)
        self._keys.remove(key)
 
    def __setitem__(self, key, item):
        if key not in self:
            self._keys.append(key)
        dict.__setitem__(self, key, item)
 
    def __deepcopy__(self, memo=None):
        if memo is None:
            memo = {}
        d = memo.get(id(self), missing)
        if d is not missing:
            return d
        memo[id(self)] = d = self.__class__()
        dict.__init__(d, deepcopy(self.items(), memo))
        d._keys = self._keys[:]
        return d
 
    def __getstate__(self):
        return {'items': dict(self), 'keys': self._keys}
 
    def __setstate__(self, d):
        self._keys = d['keys']
        dict.update(d['items'])
 
    def __reversed__(self):
        return reversed(self._keys)
 
    def __eq__(self, other):
        if isinstance(other, odict):
            if not dict.__eq__(self, other):
                return False
            return self.items() == other.items()
        return dict.__eq__(self, other)
 
    def __ne__(self, other):
        return not self.__eq__(other)
 
    def __cmp__(self, other):
        if isinstance(other, odict):
            return cmp(self.items(), other.items())
        elif isinstance(other, dict):
            return dict.__cmp__(self, other)
        return NotImplemented
 
    @classmethod
    def fromkeys(cls, iterable, default=None):
        return cls((key, default) for key in iterable)
 
    def clear(self):
        del self._keys[:]
        dict.clear(self)
 
    def copy(self):
        return self.__class__(self)
 
    def items(self):
        return zip(self._keys, self.values())
 
    def iteritems(self):
        return izip(self._keys, self.itervalues())
 
    def keys(self):
        return self._keys[:]
 
    def iterkeys(self):
        return iter(self._keys)
 
    def pop(self, key, default=missing):
        if default is missing:
            return dict.pop(self, key)
        elif key not in self:
            return default
        self._keys.remove(key)
        return dict.pop(self, key, default)
 
    def popitem(self, key):
        self._keys.remove(key)
        return dict.popitem(key)
 
    def setdefault(self, key, default=None):
        if key not in self:
            self._keys.append(key)
        dict.setdefault(self, key, default)
 
    def update(self, *args, **kwargs):
        sources = []
        if len(args) == 1:
            if hasattr(args[0], 'iteritems'):
                sources.append(args[0].iteritems())
            else:
                sources.append(iter(args[0]))
        elif args:
            raise TypeError('expected at most one positional argument')
        if kwargs:
            sources.append(kwargs.iteritems())
        for iterable in sources:
            for key, val in iterable:
                self[key] = val
 
    def values(self):
        return map(self.get, self._keys)
 
    def itervalues(self):
        return imap(self.get, self._keys)
 
    def index(self, item):
        return self._keys.index(item)
 
    def byindex(self, item):
        key = self._keys[item]
        return (key, dict.__getitem__(self, key))
 
    def reverse(self):
        self._keys.reverse()
 
    def sort(self, *args, **kwargs):
        self._keys.sort(*args, **kwargs)
 
    def __repr__(self):
        return 'odict.odict(%r)' % self.items()
 
    __copy__ = copy
    __iter__ = iterkeys
 
 
if __name__ == '__main__':
   import doctest
   doctest.testmod()
 
 
def get_dict(sh,col1,col2):
    new_dict = odict([('dummy1','dummy2')])
    for rownum in range(1,sh.nrows):
        #new_dict.update({str(sh.cell(rowx=rownum,colx=col1).value).split('.')[0]: str(sh.cell(rowx=rownum,colx=col2).value).split('.')[0]})
        new_dict.update({strip_decimal(sh.cell(rowx=rownum,colx=col1).value): strip_decimal(sh.cell(rowx=rownum,colx=col2))})
    return new_dict
 
 
def check_zeros(val):
    if val < 10:
        return "0" + str(val)
    return str(val)
 
def conv_datetime(vec):
    return str(vec[0])+ "-" +check_zeros(vec[1]) + "-" + check_zeros(vec[2])+ " " + check_zeros(vec[3])+ ":" + check_zeros(vec[4])
 
 
def get_list(sh,wb):
    datalist = []
    for rownum in range(0,sh.nrows):
        curRow = sh.row_values(rownum)
        tempRow = []
        for curCellVal in curRow:
            tempRow.append(curCellVal)
        datalist.append(tempRow)
    return datalist
 
def strip_decimal(var):
    return str(var).split('.')[0]
 
def listfile(dirname):
    file_list = []
    for f in os.listdir(dirname):
        if os.path.isfile(os.path.join(dirname, f)) and 'xls' in os.path.join(dirname, f):
            file_list.append(f)
    return file_list
 
#def read_xls_file(filename):
 
 
def write_to_xls(filename,datalist):
    wbk = xlwt.Workbook(encoding='latin-1')
    sheet = wbk.add_sheet('sheet 1')
 
 
#   style1 = xlwt.easyxf('font: name Times New Roman, colour_index black; pattern: back_colour orange, pattern thick_forward_diag,fore-colour orange')
#   style2 = xlwt.easyxf('font: name Times New Roman, colour_index black; pattern: back_colour white, pattern thick_forward_diag,fore-colour white')
#   style = style1
    # indexing is zero based, row then column
 
    #print datalist[0]
    for i in range(len(datalist)):
        for j in range(len(datalist[0])):
            #print "i: " + str(i)
            #print "j: " + str(j)
            sheet.write(i,j,datalist[i][j])
    wbk.save(filename)
 
#***********************
#********Main
 
#Read file
 
# Read original order file
wb = xlrd.open_workbook(data_path1 + Orig_order_file)
sh = wb.sheet_by_index(0)
OrigSampleList = get_list(sh,wb)
 
 
curNrOfTubesIndex = 0
curNrOfSamplesIndex = 11
 
newArray = []
 
i = 0
for item in OrigSampleList:
    if i > 0:
 
        nrOfCols = len(item)      
        curNrOfTubes = int(item[curNrOfTubesIndex])
        curNrOfSamples = int(item[curNrOfSamplesIndex])
 
        if curNrOfTubes == 0:
            newArray.append(item)

        for newRow in range(curNrOfTubes):
            itemCopy = copy.deepcopy(item)
           
            if newRow > 0:
                itemCopy[curNrOfSamplesIndex] = 0
 
            newArray.append(itemCopy)
            i = i + 1
    else:
        newArray.append(item)
        i = i + 1
   
        
 
 
#for item in newArray:
#    print item
   
 
write_to_xls(data_path_out1 + datafile_out,newArray)
