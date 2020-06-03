# CIMC_data_cleaning__Tool

Description: This is a data cleaning tool design to reduce the amount of time of manual ross-data checking of discrepancies of company names. It utilizes Apache Solr with Pysolr as its search engine and Fuzzy Search as its search algorithmn.

Instruction: Upload the desire Excel document and choose between Fuzzy Search (New document) and Fuzzy Search (Regular). Fuzzy Search (New document) will produce a new document with the results of the search append to the inputted column as its output while the Fuzzy Search (Regular) option would append a new column directly within the original document. Input the required column name for processing after the selection. 

Key tools: Apache Solr, Pysolr, Pandas, Numpy, tkinter
