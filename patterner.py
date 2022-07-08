# -*- coding: utf-8 -*-
"""
Created on Mon Jun 27 15:53:06 2022

@author: 13619993    - Alex Hamilton Fretheim

patterner - a library for analyzing the pattern sheets from Headcount Observations.

Includes:
    
    +ModelSet - a class that represents a set of patterner observations for a particular model on a particular assembly
    line. Includes the Patterners, a constructor that runs off of one of the xlsm files that includes the entire patterner
    set, and basic metadata about the patterner, including the name of the model, any submodels and the line, and a field
    for time series data. Also includes a method for outputting arrays for Neural Networks and other ML type analysis, in
    the standard format that most Python-based libraries would expect, and a method for merging with other sets.

    +TimeSeriesUtility - a simple function designed primarily to handle exceptions specific to this package for the time
    series dictionary fields of ModelSet.

    +Exceptions:

        -> WhichException - a exception thrown when an unrecognizable value for a "which" parameter is given to any 
        method in ModelSet.
    
        -> TabException - an exception thrown, primarily by the TimeSeriesUtility function, when keys are given that do not 
        match the patterner names. The most common causes for this error are a standard header used when the header is not
        standardized, and attempting to input time series data that includes dates for which no observation exists. I'm 
        considering the addition of a mass-load method that would simply not write to observations that are not present as
        a workaround to the second form of this error, and the primary workaround to the first form is not to attempt 
        standard headers when your header is not standard.
    
        -> TimeSeriesPidgeonholeException - an exception thrown when the information in a time series cannot cover all the
        patterners in the set, indicating missing or mislabelled information.

    Testing Notes: As of 7-7-22 at 1:14 PM, all functioning of this library in its present form, which includes the ModelSet
    class, the TimeSeriesUtility function, is confirmed up to primary use through testing. This confirmation does NOT include
    exception handling.


"""

import pandas as pd

#7-1-22 testing status: we have tested the constructor, but have not yet tested the addTimeSeries method.
class ModelSet:
    def __init__(self, filename, name="", descrip = "", line=""):
        self.name = name
        self.description = descrip
        self.line = line
        self.patternlist = pd.read_excel(filename, "Pattern Analysis Generator", usecols = "A", header=0) #the patternlist is the complete list of patterners in the file.
        self.patterners = {}    #contains PANDAS dataframes
        self.timeseries1 = {}   #contains doubles
        self.timeseries2 = {}   #contains doubles
        self.timeseries3 = {}   #contains doubles
        for pat in self.patternlist['Current Patterners:']:
            self.patterners[pat] = pd.read_excel(filename, pat, header=None, na_filter = False)
            #Initializing the time series slots:
            self.timeseries1[pat] = -1.0
            self.timeseries2[pat] = -1.0
            self.timeseries3[pat] = -1.0
        
        #metadata:
        self.submodels = []
        self.modelname = name
        self.modeldescription = descrip
        self.assemblyline = line

        #names for your time series:
        self.tsname1 = ""
        self.tsname2 = ""
        self.tsname3 = ""
        
        #Available time series:
        self.tsavailable = [1, 2, 3]

    #End of constructor
    
    #This method adds time series data. The standard header (sh) parameter is optional, and defaults to an empty string.
    
    #Name is also optional. When given, it adds a time series name to the appropriate field in this class for the
    #reference of users.

    #A non-trivial sh should be used if the x-array is in pure dates and the linecount file consistently uses the same
    #standard heading.
    
    def addSubmodel(self, sub):
        self.submodels.append(sub)
    
    def addTimeSeries(self, x, y, which, sh="", name=""):
        if(len(y) < len(x) | len(y) < len(self.patterners)):
            raise TimeSeriesPidgeonholeException()
        else:
            self.tsavailable.remove(which) #removes the time series slot used from the list of available slots.
        
        if(which == 1):
            TimeSeriesUtility(x, y, self.timeseries1, sh)
            self.tsname1 = name
        elif(which == 2):
            TimeSeriesUtility(x, y, self.timeseries2, sh)
            self.tsname2 = name
        elif(which == 3):
            TimeSeriesUtility(x, y, self.timeseries3, sh)
            self.tsname3 = name
        else:
            raise WhichException()
        
    def patternMatch(self, key):
        return key in self.patterners.keys()   #True or false



#End of class

#Utility functions:
    
def TimeSeriesUtility(X, Y, dct, sh):
    for i in range(len(X)):
        if (sh + X[i]) in dct.keys():
            dct[sh + X[i]] = Y[i]
        else:
            raise TabException(sh + X(i))

        

#Exceptions:

class WhichException(Exception):
    
    def __init__(self):
        super().__init__("ERROR: Incorrect value for which parameter.")

class TabException(Exception):
    
    def __init__(self, X):
        super().__init__("ERROR: This x-value does not match a patterner name: " + X + " This error can be caused by incorrect use of the standard header parameter (sh), or by adding time series information for times you did not observe the line at.")

class TimeSeriesPidgeonholeException(Exception):
    
    def __init__(self):
        super().__init__("ERROR: No Time Series information for at least one patterner")