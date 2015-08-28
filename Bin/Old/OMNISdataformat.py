#! /opt/local/bin/python

import odict
import os


import xlrd
import xlwt
import math
import datetime


import functools
import operator

import matplotlib
matplotlib.use('TkAgg')

import numpy as np
import matplotlib.pyplot as plt

import copy

#paths
data_path1 = '../Data/'
data_path_out1 = '../Result/'

#Files
Orig_order_file = 'OMNISdata.xls'
datafile_out = 'OMNISdataFormated.xls'

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
        datalist.append(sh.row_values(rownum))

        #Quick fix for date n time, to be shaped up
        #date_time = xlrd.xldate_as_tuple(sh.row_values(1)[0], wb.datemode)
        #datetime_str = conv_datetime(date_time) #N.B: fix with datetime instead
        #datalist[-1][0] = datetime_str
        #a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, book.datemode))
        #datalist.append(add_post(sh.row_values(rownum)))

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
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet('sheet 1')

 #   style1 = xlwt.easyxf('font: name Times New Roman, colour_index black; pattern: back_colour orange, pattern thick_forward_diag,fore-colour orange')
 #   style2 = xlwt.easyxf('font: name Times New Roman, colour_index black; pattern: back_colour white, pattern thick_forward_diag,fore-colour white')
 #   style = style1
    # indexing is zero based, row then column
    for i in range(len(datalist)):
        sheet.write(i,1,str(datalist[i]))
    wbk.save(filename)

#***********************
#********Main

#Read file

# Read original order file
wb = xlrd.open_workbook(data_path1 + Orig_order_file)
sh = wb.sheet_by_index(0)
OrigSampleList = get_list(sh,wb)


curNrOfTubesIndex = 2
curNrOfSamplesIndex = 5

newArray = []

for item in OrigSampleList:
    #print item
    nrOfCols = len(item)
    
    #print item[curNrOfTubesIndex]
    #print item[curNrOfSamplesIndex]
    curNrOfTubes = int(item[curNrOfTubesIndex])
    curNrOfSamples = int(item[curNrOfSamplesIndex])
    #print "Number of samples. " + str(curNrOfSamples)
    print "Number of tubes. " + str(curNrOfTubes)

    for newRow in range(curNrOfTubes):
        itemCopy = copy.deepcopy(item)
        for el in range(len(itemCopy)):
            itemCopy[el] = str(itemCopy[el]).strip("'")
        
        if newRow > 0:
            itemCopy[curNrOfSamplesIndex] = 0
        newArray.append(str(itemCopy).strip("[").strip("]"))
        


for item in newArray:
    print item
    

write_to_xls(data_path1 + datafile_out,newArray)

