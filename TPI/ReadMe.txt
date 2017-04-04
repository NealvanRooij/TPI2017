ListingACL Script
---------------------
Those scripts list the arborescence security of a given path and print them on a Excel sheet. 

It works with the added module Import-Excel, already added in the package. 
You can read more about it here: https://github.com/dfinke/ImportExcel
To install Import-Excel, start .\ImportExcel-master\InstallModule.ps1

There is no need to install anything, running the script GetACL.ps1 launches everything
Use the Task Scheduler to plan the automatization of it.
The path analysed can be change in .\Settings\source.txt

This project version: 1.0
Last update: 31.03.2017
Author: Neal van Rooij, Practeo SA
