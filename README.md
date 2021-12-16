# Shapefile2Excel
Python and vba tool to import Ersi Shapefile as Shapes into Excel. 

## Files
1) README.md - This File
2) LICENSE - Apache License 
3) load_shapes_to_excel_GITHUB.xlsm - Excel Spreadsheet and Macro to complete Shape creation from ERSI Shapefile
4) main.py - Python Script to create input to Excel VBA, from ERSI Shapefile (controlled from VBA)
5) python_script.bat - Batch file, which runs Python and called from VBA

## Instructions (from Excel)
1) Click on **"RUN SHAPEFILE2EXCEL"** button
2) In File Dialog, select Shapefile to Open

## Output (in Excel)
1) 3 Sheets are created "MAPS", "MAPS META", "MAPS META GRP"
   - "MAPS" will contain the Shapes polygons
   - "MAPS META" will contain the Shape Ids, Group Ids (parent), Group Name (Parent Group Name - for multipolygons), and Name (Polygon Name)
   - "MAPS META GRP" contains the hierarchie of parents and children relationship
2) Each Polygon is grouped and named as per Group Name (Sub polygons are named as per Name)

## Requirement
1) MS Excel
2) Anaconda Python 3.8

## Initial Instructions:
### In Excel Spreadsheet (attached)
1) In Main, Cell B2, Enter location of where batch file (of this package) is saved
2) In Main, Cell B3, Enter location of where python script (part of this package) is saved

### In Command Batch file (python_script.bat, attached)
1) Replace "[ROOT FOLDER FOR ANACONDA]" with location where python Anaconda is installed


### License:
   Copyright 2021 Rahman Mohamud Faisal MOORABY                                
   Licensed under the Apache License, Version 2.0 (the "License");             
   you may not use this file except in compliance with the License.            
   You may obtain a copy of the License at                                     
       http://www.apache.org/licenses/LICENSE-2.0                              
   Unless required by applicable law or agreed to in writing, software         
   distributed under the License is distributed on an "AS IS" BASIS,           
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.    
   See the License for the specific language governing permissions and         
   limitations under the License.                                              
