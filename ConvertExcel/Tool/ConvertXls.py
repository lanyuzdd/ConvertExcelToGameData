import xlrd
import os
import sys
from optparse import OptionParser

reload(sys)
sys.setdefaultencoding("utf-8")

class ExcelLoad():
	def __init__(self,_path):
		self.path = _path
		self.data = ""
		self.sheets_name = []

	def covertXml(self,_filename):
		self.data = xlrd.open_workbook(self.path)
		self.sheets_name = self.data.sheet_names()
		# for el in self.sheets_name:
		# 	table = self.data.sheet_by_name(el)
		# 	self.loadData(el,table)
		table = self.data.sheet_by_name(self.sheets_name[0])
		self.loadData(_filename.split(".")[0],table)

	def loadData(self,_name,_data):
		if(param == "lua"):
			self.loadDataLua(_name,_data)
		elif(param == "json"):
			self.loadDataJson(_name,_data)
		elif(param == "js"):
			self.loadDataJs(_name,_data)

	def loadDataLua(self,_name,_data):
		file = open(outPath + _name + "Data.lua","w")
		file.writelines(_name + " = {\n")
		file.writelines("cc.exports." + _name + "Data = {\n")
		cols = _data.ncols
		rows = _data.nrows
		for i in range(rows):
			if(i >= 2):
				file.write("  [\"" + str(self.isTypeTitle(_data.cell(i,1))) + "\"] = {")
				for j in range(cols):
					if(j >= 2):
						file.write("[\"" + str(_data.cell(1,j).value) + "\"] = ")
						cellString = self.isType(_data.cell(i,j))
						file.write(cellString)
						if(j != cols - 1):
							file.write(",")
				file.write(" },\n")
		file.writelines("}")

	def loadDataJson(self,_name,_data):
		file = open(outPath + _name + "Data.json","w")
		file.writelines("{\n")
		cols = _data.ncols
		rows = _data.nrows
		for i in range(rows):
			if(i >= 2):
				file.write("    \"" + str(self.isTypeTitle(_data.cell(i,1))) + "\" : {")
				for j in range(cols):
					if( j>= 2):
						file.write("\"" + str(_data.cell(1,j).value) + "\" : ")
						cellString = self.isType(_data.cell(i,j))
						file.write(cellString)
						if(j != cols - 1):
							file.write(",")
				file.write(" },\n")
		file.writelines("}")

	def loadDataJs(self,_name,_data):
		file = open(outPath + _name + "Data.js","w")
		file.writelines("var "+_name+"Data = {\n")
		cols = _data.ncols
		rows = _data.nrows
		for i in range(rows):
			if(i >= 2):
				file.write("    " + str(self.isTypeTitle(_data.cell(i,1))) + " : {")
				for j in range(cols):
					if(j >= 1):
						file.write(str(_data.cell(1,j).value) + ": ")
						cellString = self.isType(_data.cell(i,j))
						file.write(cellString)
						if(j != cols - 1):
							file.write(",")
				file.write(" },\n")
		file.writelines("}")

	def isType(self,_st):
		if isinstance(_st.value,str):
			sr = "\"" + _st.value + "\""
		elif _st.ctype == xlrd.XL_CELL_NUMBER:
			if _st.value == int(_st.value):
				sr = str(int(_st.value))
			else:
				sr = str(_st.value)
		elif _st.ctype == xlrd.XL_CELL_BOOLEAN:
			if _st.value == 1:
				sr = "true"
			else:
				sr = "false"
		else:
			sr = "\"" + str(_st.value) + "\""
		return sr

	def isTypeTitle(self,_st):
		if _st.ctype == xlrd.XL_CELL_NUMBER:
			sr = str(int(_st.value))
		else:
			sr = _st.value
		return sr


def main():	
	for parent,dirnames,filenames in os.walk(inPath): 
		for filename in filenames: 
			if filename[filename.find("."):] ==".xls" or filename[filename.find("."):] ==".xlsx":                   
				obj = ExcelLoad(os.path.join(parent,filename))
				obj.covertXml(filename)

def _check_python_version():
    major_ver = sys.version_info[0]
    if major_ver > 2:
        print ("The python version is %d.%d. But python 2.x is required. (Version 2.7 is well tested)\n"
               "Download it here: https://www.python.org/" % (major_ver, sys.version_info[1]))
        return False
    return True

if __name__ == "__main__":
	if not _check_python_version():
		exit()

	parser = OptionParser()
	parser.add_option('-i', '--input', dest='opt_input', help='directory of input')
	parser.add_option('-o', '--output', dest='opt_output', help='directory of output')
	parser.add_option('-l', '--language', dest='opt_language', help='convert language type')
	opts, args = parser.parse_args()

	inPath = opts.opt_input
	outPath = opts.opt_output
	param = opts.opt_language

	main()