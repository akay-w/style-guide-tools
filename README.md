# style-guide-tools
Style guide check tools for use on Excel files

Created in Python 3.7.1

Non-standard dependent libraries: openpyxl

These are 2 different translation style guide check tools that I created for specific types of projects.

Style_guide_tool_1.py:
I wrote this to work on a specific format of Excel file that we were using (sheet named "checkSheet", with source text in column 4 and target text in column 6, with the check results output in column 15). It checks for correction translations of specific terms, correct symbol usage, and capitalization rules for different file types according to a style guide that we had for specific projects.

Style_guide_tool_2.py:
This tool was designed to run on an .xlsm file with the source text in column 1 and the target text in column 2, with the checks starting in row 2 (leaving room for a header). It checks for correct symbol usage, correct translation of VDC/VAC, spaces between numbers and common units, forbidden terms, commas in numbers with 4 or more digits, matching numbers of brackets, contractions, etc. according to a style guide that we had for specific projects.
