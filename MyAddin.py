import sys
import datetime
import traceback
import numpy as np
import pandas as pd

import itertools

import uno
import unohelper
from com.andlla.calc.MyAddin import MyAddinXI

ERR_CALC = '#CALC!'

ERR_NA = None
ERR_NUM = float("NaN") # !NUM

class MyAddinImpl(unohelper.Base, MyAddinXI):

	def __init__(self, ctx):
		self.ctx = ctx
		self.debug = False
		self.appendToArrayLen1 = True

	def myPython( self, *args ):
		return sys.version + " " + str(datetime.datetime.now())

	def myPythonModules(self, *args):
		try:
			bufMods = sys.modules
			retVal = tuple( (it, str(bufMods[it])) for it in bufMods.keys() )

			if self.debug: print("Return: " + str(retVal))

			return retVal
		except:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return myErr



	def myDebug( self, isDebug ):
		try:
			self.debug = bool(isDebug)
			return("Debug is: " + str(self.debug))
		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return myErr


	def myFilter(self, array, include):
		try:
			if self.debug:
				print("****************************")
				print("myFilter_Actual: " + str(datetime.datetime.now()))
				print("array: " + str(array))
				print("include: " + str(include))


			arrayX = np.array(array).flatten()
			arrayI = np.array(include).flatten()

			retAns = arrayX[arrayI == 1]
			if self.debug:
				print("Return: " + str(retAns))

			retAns = tuple( (it, ) for it in retAns )
			if len(retAns) == 0:
				retAns = ((ERR_NA,),)

			if len(retAns) == 1 and self.appendToArrayLen1:
				retAns = (*retAns, (ERR_NA,))

			if self.debug:
				print("Return: " + str(retAns))

			return retAns

		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return ((myErr,),)



	def myTextSplit (self, text, delimiter):
		try:
			if self.debug:
				print("****************************")
				print("myTextSplit: " + str(datetime.datetime.now()))
				print("text: " + str(text))
				print("delimiter: " + str(delimiter))


			retAns = text.split(delimiter)
			retAns = tuple( (it, ) for it in retAns )

			if len(retAns) == 0:
				retAns = ((ERR_NA,),)

			if (len(retAns) == 1) and self.appendToArrayLen1:
				retAns = (*retAns, (ERR_NA,))

			if self.debug:
				print("Return: " + str(retAns))

			return retAns

		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return myErr


	def myUnique (self, array):
		try:
			if self.debug:
				print("****************************")
				print("myUnique: " + str(datetime.datetime.now()))
				print("array: " + str(array))


			retAns = tuple(dict.fromkeys(array))

			if self.debug:
				print("Return: " + str(retAns))

			return retAns

		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return ((myErr,),)


	def mySort_Actual(self, sortArray, sortIndex=1, sortOrder=1):
		try:
			if self.debug:
				print("****************************")
				print("mySort: " + str(datetime.datetime.now()))
				print("sortOrder: " + str(sortOrder))
				print("sortIndex: " + str(sortIndex))
				print("sortArray: " + str(sortArray))

			sortIndex = int(sortIndex)
			sortOrder = int(sortOrder)

			if sortOrder == 1:
				reverse = False
			elif sortOrder == -1:
				reverse = True
			else:
				retStr = "Unrecognized Sort Order: " + str(sortOrder)
				if self.debug: print("Return: " + retStr)
				return retStr

			stringify = True
			if self.debug: print("stringify-pre: " + str(stringify))

			if all(isinstance(item[sortIndex - 1], float) for item in sortArray):
				stringify = False

			if self.debug: print("stringify-post: " + str(stringify))

			retVal = tuple(sorted(sortArray, key=lambda r: str(r[sortIndex - 1]) if stringify else r[sortIndex - 1], reverse=reverse))
			if self.debug: print("Return: " + str(retVal))
			return retVal

		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return ((myErr,),)


	def mySort(self, *args):
		try:
			return self.mySort_Actual(*args)
		except Exception as error:
			print(traceback.format_exc())



	def myInterp(self, arrayX, arrayY, interpX):
		try:
			if self.debug:
				print("****************************")
				print("myInterp: " + str(datetime.datetime.now()))
				print("arrayX: " + str(arrayX))
				print("arrayY: " + str(arrayY))
				print("interpX: " + str(interpX))

			arrayX = np.array(arrayX).flatten()
			arrayY = np.array(arrayY).flatten()

			# sort values
			arraySortIdx = np.argsort(arrayX)
			arrayX = arrayX[arraySortIdx]
			arrayY = arrayY[arraySortIdx]

			if self.debug:
				print("arrayX: " + str(arrayX))
				print("arrayY: " + str(arrayY))

			retVal = np.interp(interpX, arrayX, arrayY)

			if self.debug: print("Return: " + str(retVal))

			return retVal

		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return ((myErr,),)



	def myXLookup (self, lookupArray, lookupValue, returnArray):
		try:
			if self.debug:
				print("****************************")
				print("myXLookup: " + str(datetime.datetime.now()))
				print("lookupArray: " + str(lookupArray))
				print("lookupValue: " + str(lookupValue))
				print("returnArray: " + str(returnArray))

			try:
				findIdx = lookupArray.index((lookupValue,))
				retVal = (returnArray[findIdx],)
			except ValueError:
				retVal = (("#N/A",), )

			if self.debug: print("Return: " + str(retVal))

			return retVal
		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return ((myErr,),)


	myWebCsvCache = dict()

	def myWebCsv(self, url):
		try:
			if self.debug:
				print("****************************")
				print("myWebCsv: " + str(datetime.datetime.now()))
				print("url: " + str(url))

			retVal = self.myWebCsvCache.get(url)

			if retVal == None:
				if self.debug: print("No data in cache, request from web")
				buf = pd.read_csv(url)
				resA = tuple(ith for ith in buf)
				resB = list(tuple(it) for it in buf.itertuples(index=False))
				resB.insert(0, resA)
				retVal = tuple(resB)
				self.myWebCsvCache[url] = retVal
			else:
				if self.debug: print("Cache Hit")

			if self.debug:
				print(str(retVal))

			return retVal

		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return ((myErr,),)



	def myWebCsvClear(self, url):
		self.myWebCsvCache = dict()
		return "myWebCsv Cache Cleared: " + str(datetime.datetime.now())

	def myWebCsvList(self, url):
		try:
			retVal = tuple( (str(it),) for it in self.myWebCsvCache.keys() )
			if (len(retVal) == 1) and self.appendToArrayLen1:
				retVal = (*retVal, (ERR_NA,))

			if self.debug:
				print(str(retVal))

			return retVal
		except Exception as error:
			myErr = traceback.format_exc()
			print(myErr)
			if self.debug: return ((myErr,),)


def createInstance(ctx):
#	print("MyAddinImpl.createInstance: " + str(ctx))
	return MyAddinImpl(ctx)


#print("MyAddinImpl.g_ImplementationHelper")
g_ImplementationHelper = unohelper.ImplementationHelper()

#print("MyAddinImpl.g_ImplementationHelper.addImplementation")
g_ImplementationHelper.addImplementation(createInstance, 'com.andlla.calc.MyAddin.python.MyAddinImpl',('com.sun.star.sheet.AddIn',),)

