# dreq_tools
Tools for updating the PRIMAVERA data request.

These files have not been tested since 2017 and will only work with
Python 2.7. The data request included in the `xls` directory is from version
01.00.13 of the CMIP6 data request.

The following Python libraries are required and can be installed with pip:

openpyxl

dreqPy

Running `upgrade_dreq.py` will load the existing PRIMAVERA 01.00.07 version of data request
and upgrade it to version 01.00.13 of the HighResMIP data request.

`find_missing_vars.py` can be run to identify variables that are in version 01.00.13 of the 
HighResMIP data request but aren't in the current 01.00.07 version of the PRIMAVERA
data request.
