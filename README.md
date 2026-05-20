## Door Schedule Automation

## Overview : 
this project is a pyhton script that's update a main excel file (door scedule) from multiple revised floor schedule files .

The script matches doors by `ROOM NUMBER`(key) and updates the master file based on matching column headers . 

## Problem
The revised floor files and the master file can have different column orders, which makes manual updating slow and error-prone.

## Solution
The script reads each revised floor file, stores door data in a dictionary using `ROOM NUMBER` as the key, then updates matching rows in the master Excel file.

## Features
- Reads multiple `.xlsx` files from a folder
- Matches doors using `ROOM NUMBER`
- Updates the master schedule based on header names
- Skips temporary Excel files
- Saves a new updated Excel file instead of overwriting the original

## Requirements
- Python 3
- openpyxl
