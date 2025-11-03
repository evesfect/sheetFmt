# sheetFmt
XLSX bulk formatting program to recover non-standard and locally stored data of any format.
Created by @evesfect

Builds a mapping graph with scanned data gradually, supports ai suggestion for fast mapping.
Extracts the data from inputted sources and converts into ready-to-upload format as wished, works in well defined phases, can be chained to achieve desired output.

Scan->Map->Format->Convert, each phase can be run independently. Check the config for data targeting and extraction values.

Uses python scripts to interact with xlsx files and Golang for everything else. Works from terminal and has a terminal gui for mapping. Can generate auto suggestions with configurable trust weight if provided with an llm key.

I have personally used it to recover 2 years worth of non-standard data of a company, standardize it and upload it to Salesforce.
