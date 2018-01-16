import csv
import docx
from docx import Document
import json
import getopt
import re
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl import Workbook
import sys
import os
from shutil import copyfile
import time





#Next attempt to create a DocHandle for the current Word input file
temp = 'Book1.xlsm'
destination = 'Book1Out.xlsm'
copyfile(temp, destination)
try:
	coursePlan = openpyxl.load_workbook(destination, read_only=False, keep_vba=True)
	coursePlan.save(destination)   #The save here is confirmed to corrupt - why?
except IOError:
	print("++-->> File open for: ", inputfile, " FAILED!")
print('Complete')