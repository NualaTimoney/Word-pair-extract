# Word-pair-extract
Python script to take text from the field from a patentscope .xls file, and count up the frequency of pairs of neighbouring words excluding given exceptions.

Usage:

-i Input file, xls file downloaded from patentscope, this file will scan for word pairings in the title field
-r This is a file containing excluded word pairings, which don't have a useful meaning (i.e. "of a")
-o This file, a ".json" file, is where the filtered dictionary is saved, where all of the entries have been counted
