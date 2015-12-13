# -*- coding: utf-8 -*-
"""
Created on Thu Dec 10 13:27:59 2015
@author: Nuala_Timoney

Usage: Word_pair_extraction
-h help
-i Input file, xls file downloaded from patentscope, this file will scan for word pairings in the title field
-r This is a file containing excluded word pairings, which don't have a useful meaning (i.e. "of a")
-o This file, a ".json" file, is where the filtered dictionary is saved, where all of the entries have been counted
"""

# Function to extract the command line arguments, returns the filenames
def main(argv):
   import sys, getopt
   inputfile = ''
   removelistfile = ''
   outputfile = ''
   try:
      opts, args = getopt.getopt(argv,"hi:r:o:")
   except getopt.GetoptError:
      print 'test.py -i <inputfile> -r <removelistfile> -o <outputfile>'
      print '(ERROR)'
      sys.exit(2)        
   for opt, arg in opts:
      if opt == '-h':
         print 'test.py -i <inputfile> -o <outputfile>'
         sys.exit()
      elif opt in ("-i", "--ifile"):
	 
         inputfile = arg
      elif opt in ("-r", "--rfile"):
	 
         removelistfile = arg
      elif opt in ("-o", "--ofile"):
         outputfile = arg       
   return (inputfile, removelistfile, outputfile)

     
     
# Function to extract word pairs from the title field of the data,
# returns all of the word pairs in one title field.
def extractdata(book):
    import re
    sh = book.sheet_by_index(0)
    dataRows = sh.nrows
    print dataRows

    Pairs = []
    temp = []
  
    for x in range(2, dataRows):
        ab = str(sh.cell_value(rowx=x, colx=2).encode('ascii', 'ignore'))
        #gathers a unicode string of all paired words
        a = re.sub("[^\w]", " ",  ab).split()
        for i in range(len(a)):
            if i < len(a)-1:
                temp = [a[i].lower(), a[i+1].lower()]
                Pairs.append(" ".join(temp))
                temp = []
    return Pairs

# Function to filter the main list, based on a list of excluded terms returns filtered list
def filter_list(title_words, remove_list):
    text = title_words
    for item in remove_list:
        text = list(filter(lambda x: x!= item, text))
    return text
    
# Function to do a count of all the terms in text. Does not include double counting
# Much much faster than using dict to do the counting   
def quick_count(text):
    y = []
    key = []
    value = []
    inc = 0
    for x in range(0,len(text)):
        if text[x] not in y:
            key.append(text[x])
            value.append(text.count(text[x]))
            y.append(text[x])
            inc= inc+1
    single_dict = dict(zip(key, value))
    return single_dict
    



if __name__ == "__main__":
   import sys
   #sys.path.append("/Users/Nuala/anaconda/lib/python2.7/site-packages")
   
   import xlrd,json
   inputfile, removelist, outputfile = main(sys.argv[1:])
   book = xlrd.open_workbook(inputfile)
   title_words = extractdata(book)
   with(open(removelist)) as f:
	remove_list = f.read().splitlines()
   text = filter_list(title_words, remove_list)
   single_dict = quick_count(text)
   with open(outputfile, 'w') as fp:
    json.dump(single_dict, fp)
