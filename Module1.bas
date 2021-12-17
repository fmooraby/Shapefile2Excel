Attribute VB_Name = "Module1"
'''''''''''''''''''''''''''''''''''''
'___________________________________'
'''         SHAPEFILE2EXCEL       '''
'-----------------------------------'
'''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Copyright 2021 Rahman Mohamud Faisal MOORABY                                '
'   Licensed under the Apache License, Version 2.0 (the "License");             '
'   you may not use this file except in compliance with the License.            '
'   You may obtain a copy of the License at                                     '
'       http://www.apache.org/licenses/LICENSE-2.0                              '
'   Unless required by applicable law or agreed to in writing, software         '
'   distributed under the License is distributed on an "AS IS" BASIS,           '
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.    '
'   See the License for the specific language governing permissions and         '
'   limitations under the License.                                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public name_index_selection As Integer

Sub load_shapes()

    '''' VARIABLE DECLARATION
    Dim file_path As String
    Dim get_directory As String
    Dim python_exe As String
    Dim python_script As String
    Dim EXE_str As String
    Dim wsh As Object
    Dim waitCompleted As Boolean: waitCompleted = True
    Dim winStyle As Integer: winStyle = 1

    Dim w As String
    Dim h As String
    Dim x_off As String
    Dim y_off As String
    Dim simplify As String
    
    Dim line_string As String
    Dim line_string_arr() As String
    Dim file_2 As String
    Dim x_str() As String
    Dim y_str() As String
    
    Dim regex As Object
    
    
    
    '''' SET REGEX PARAMETERS TO NOT PARSE COMMAS BETWEEN QUOTES
    Set regex = CreateObject("vbscript.regexp")
    regex.IgnoreCase = True
    regex.Global = True
    regex.Pattern = ",(?=([^" & Chr(34) & "]*" & Chr(34) & "[^" & Chr(34) & "]*" & Chr(34) & ")*(?![^" & Chr(34) & "]*" & Chr(34) & "))"
    
    '''' OPEN FILE CREATED BY PYTHON SCRIPT
    With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "ShapeFile", "*.shp", 1
        'Show the dialog box
        .Show
        
        'Store in fullpath variable
        file_path = .SelectedItems.Item(1)
        get_directory = Left(file_path, InStrRev(file_path, Application.PathSeparator))
    End With
    
    '''' GET OTHER PARAMETERS (SCALE, OFFSET, ETC.)
    w = Chr(34) & ThisWorkbook.Sheets("Main").Range("C6").Value & Chr(34)           ' WIDTH
    h = Chr(34) & ThisWorkbook.Sheets("Main").Range("C7").Value & Chr(34)           ' HEIGHT
    x_off = Chr(34) & ThisWorkbook.Sheets("Main").Range("C8").Value & Chr(34)       ' HORIZONTAL OFFSET
    y_off = Chr(34) & ThisWorkbook.Sheets("Main").Range("C9").Value & Chr(34)       ' VERTICAL OFFSET
    simplify = Chr(34) & ThisWorkbook.Sheets("Main").Range("C10").Value & Chr(34)   ' TO SIMPLIFY OR NOT TO OF POLYGON (SIMPLIFY WOULD RUN QUICKER AND WITH A SMALLER BYTE SIZE
    
    python_exe = ThisWorkbook.Sheets("Main").Range("B2").Value      ' PYTHON LOCATION
    python_script = ThisWorkbook.Sheets("Main").Range("B3").Value   ' LOCATION OF PYTHON SCRIPT
    
    ' BUILD COMMAND
    EXE_str = python_exe & " " & python_script & " " & w & " " & h & " " & x_off & " " & y_off & " " & Chr(34) & "file=" & file_path & Chr(34) & " " & simplify
    ThisWorkbook.Sheets("Main").Range("A20").Value = EXE_str ' SAVE COMMAND IN CELL A20 (FOR DEBUGGING PURPOSE)
    
    '''' RUN PYTHON SCRIPT
        '''' WAIT FOR SCRIPT TO COMPLETE
    Set wsh = VBA.CreateObject("WScript.Shell")
    
    '''' OPEN CSV FILE CREATED BY SCRIPT FOR CHOSING COLUMNS:
    Open get_directory + "\out.csv" For Input As #1
    num_line = 0
    
    Dim name_arr() As String    ' ARRAY TO SAVE LIST OF COLUMN NAMES (HEADER OF FILE)
    Dim col_sample As String    ' STRING TO SAVE LIST OF FIRST ROW SAMPLE DATA
    
    ' LOAD ONLY 2 LINES OF FILE
    For ILine = 1 To 2
        Line Input #1, line_col_string              ' READ LINE OF FILE
        If ILine = 1 Then                           ' LINE 1 (HEADER)
         name_arr = Split(line_col_string, ",")     ' ARRAY TO SAVE LIST OF COLUMN NAMES (HEADER OF FILE)
        End If
        
        If ILine = 2 Then                                           ' LINE 2 (SAMPLES)
            col_sample = Replace(line_col_string, ",", vbNewLine)   ' ARRAY TO SAVE LIST OF FIRST ROW SAMPLE DATA (AND REPLACE "," BY NEW LINE)
        End If
    Next
    
    '''' LOAD USER FORM TO ALLOW USER TO CHOOSE WHICH FIELD CONTAINS NAME OF EACH POLYGON
    columns_form.ListBox1.Height = 10 * (UBound(name_arr) - LBound(name_arr) + 1)           ' RESIZE LISTBOX ACCORDING TO NUMBER OF ENTRIES (LISTBOX CONTAINS NAME OF FIELDS)
    columns_form.LabelSamples.Height = 10 * (UBound(name_arr) - LBound(name_arr) + 1)       ' RESIZE LABEL IN THE SAME WAY. LABEL HAS THE SAMPLES
    columns_form.ok_button.Top = 36 + (10 * (UBound(name_arr) - LBound(name_arr) + 1))      ' MOVE OK BUTTON ACCORDING TO SIZE OF LISTBOX
    columns_form.Height = 93.75 + (10 * (UBound(name_arr) - LBound(name_arr) + 1))          ' RESIZE WHOLE FORM ACCORDINGLY
    
    ' LOAD VALUES
    columns_form.ListBox1.List = name_arr           ' POPULATE LISTBOX WITH COLUMN FIELDS
    columns_form.LabelSamples.Caption = col_sample  ' POPULATE LABEL WITH SAMPLES
    
    ' SHOW FORM
    columns_form.Show   ' OPEN COLUMNS_FORM FOR FURTHER CODES
    
    ' CLOSE FILE
    Close #1
    
    '''' OPEN CSV FILE CREATED BY SCRIPT:
    Open get_directory + "\out.csv" For Input As #1
    

    '''' CREATE SHEETS FOR MAPS META DATA
    
    '''' SHEET MAPS WILL CONTAIN THE SHAPES
    
    '''' MAP META IS ORGANISED AS:
    ''''    id          : UNIQUE ID
    ''''    Grp id      : GROUP ID (PARENT POLYGON ID (FOR MULTIPOLYGONS LAYER)
    ''''    Sub id      : SUB ID (CHILD POLYGON ID)
    ''''    Grp Name    : GROUP NAME (PARENT POLYGON NAME (FOR MULTIPOLYGONS LAYER)
    ''''    Name        : NAME (CHILD POLYGON NAME - SUFFIXING GRP NAME WITH SUB ID)
    
    
    Dim is_meta As Integer
    Dim is_maps As Integer
    is_meta = 0
    is_maps = 0
    For Each sht In ActiveWorkbook.Sheets
        If sht.Name = "MAPS" Then           ' SHEET "MAPS" ALREADY EXISTS
            is_maps = 1
        End If
        If sht.Name = "MAPS META" Then      ' SHEET "MAPS META" ALREADY EXISTS
            is_meta = 1
        End If
    Next sht
    
    '''' IF SHEETS "MAPS" AND "MAPS META" DO NOT EXIST, THEN CREATE THEM
    Dim ws As Worksheet
    If is_maps = 0 Then
        With ThisWorkbook
            Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            ws.Name = "MAPS"
            ws.Tab.Color = RGB(10, 0, 0)
        End With
    End If
    
    If is_meta = 0 Then
        With ThisWorkbook
            Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            ws.Name = "MAPS META"
            ws.Tab.Color = 163433
        End With
    End If
    
    '''' POPULATE HEADERS OF MAPS META SHEET
    ThisWorkbook.Sheets("MAPS META").Cells(1, 1) = "id"
    ThisWorkbook.Sheets("MAPS META").Cells(1, 2) = "Grp id"
    ThisWorkbook.Sheets("MAPS META").Cells(1, 3) = "Sub id"
    ThisWorkbook.Sheets("MAPS META").Cells(1, 4) = "Grp Name"
    ThisWorkbook.Sheets("MAPS META").Cells(1, 5) = "Name"
    
    
    '''' INITIALISE PARAMETERS FOR FILE PROCESSING
    Dim id_row As Integer
    id_row = 1
    
    Dim r As Double
    r = 0
    
    Dim intList_X() As Double
    Dim intList_Y() As Double
    Dim xpos As Integer
    Dim ypos As Integer
    
    Dim ub As Integer
    Dim arr() As Single
    
    ' SET SCREEN UPDATING TO FALSE TO RUN THE CODE FASTER
    Application.ScreenUpdating = False
        
    '''' PROCESS FILE
    Do Until EOF(1)         ' UNTIL END OF FILE
        Line Input #1, line_string          ' READ EACH LINE
        line_string_arr = Split(regex.Replace(line_string, ";"), ";")   ' SPLIT LINE TO ARRAY
        
        '''' IDENTIFY X and Y FIELD position
        i_pos = 0
        If r = 0 Then
            For Each l In line_string_arr
                If l = "x" Then
                    xpos = i_pos + 1
                End If
                If l = "y" Then
                    ypos = i_pos + 1
                End If
                i_pos = i_pos + 1
            Next
        End If
        
        '''' IF NOT HEADER (I.E. DATA)
        If r > 0 Then
            '''' POPULATE MAPS META DATA
            ThisWorkbook.Sheets("MAPS META").Cells(r + 1, 1) = r            ' CREATE UNIQUE ID FOR EACH POLYGON (I.E. ROW)
            ThisWorkbook.Sheets("MAPS META").Cells(r + 1, 2) = line_string_arr(0)           ' GET GROUP ID FROM COL 1
            ThisWorkbook.Sheets("MAPS META").Cells(r + 1, 3) = line_string_arr(1)           ' GET SUB ID FROM COL 2
            ThisWorkbook.Sheets("MAPS META").Cells(r + 1, 4) = line_string_arr(name_index_selection)            'GET GROUP NAME (AS PER SELECTION FROM USER)
            ThisWorkbook.Sheets("MAPS META").Cells(r + 1, 5) = line_string_arr(name_index_selection) & "_" & CStr(line_string_arr(1))   ' GENERATE A UNIQUE NAME BASED ON GROUP NAME AND SUB ID
            
            ' REDIMENSIONING OF ARRAYS
            ReDim xstr(LBound(Split(line_string_arr(xpos), ",", -1, vbTextCompare)) To UBound(Split(line_string_arr(xpos), ",", -1, vbTextCompare)))
            ReDim ystr(LBound(Split(line_string_arr(ypos), ",", -1, vbTextCompare)) To UBound(Split(line_string_arr(ypos), ",", -1, vbTextCompare)))
            
            '''' GET COORDINATES OF POLYGONS
            x_str = Split(line_string_arr(xpos), ",", -1, vbTextCompare)    ' X COORDINATES ARRAY OF POLYGON
            y_str = Split(line_string_arr(ypos), ",", -1, vbTextCompare)    ' Y COORDINATES ARRAY OF POLYGON
            
            ' REDIMENSION OF ARRAY FOR SHAPES AND GROUP META
            ' FLOAT ARRAYS (AS OPPOSED TO STRING) TO CREATE SHAPES
            ReDim arr(LBound(x_str) To UBound(x_str), 1)
            ReDim intList_X(LBound(x_str) To UBound(x_str))
            ReDim intList_Y(LBound(y_str) To UBound(y_str))
        
            For d = LBound(x_str) To UBound(x_str)
                intList_X(d) = CDbl(Replace(x_str(d), """", ""))
                intList_Y(d) = CDbl(Replace(y_str(d), """", ""))
                
                arr(d, 0) = CDbl(Replace(x_str(d), """", ""))
                arr(d, 1) = CDbl(Replace(y_str(d), """", ""))
            Next d
            
            '''' CREATE POLYGONS AND NAME ACCORDINGLY
            is_shapes_tab = 0
            
            '''' CHECK IF SHEET MAPS EXISTS
            For Each sht In ActiveWorkbook.Sheets
                If sht.Name = "MAPS" Then
                    is_shapes_tab = 1
                End If
            Next sht
            
            '''' CREATE SHAPE TAB (MAPS) IF IT DOES NOT EXIST
            If is_shapes_tab = 0 Then
                With ThisWorkbook
                    Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
                    ws.Name = "MAPS"
                    ws.Tab.Color = RGB(10, 10, 0)
                End With
            End If
            ThisWorkbook.Sheets("MAPS").Activate    ' ACTIVATE SHEET MAPS
            Set ws = ThisWorkbook.ActiveSheet
            Set ws_shp = ws.Shapes.AddPolyline(arr) ' CREATE POLYGON BASED ON X Y COORDINATES AS SAVED IN ARR
            ws_shp.Name = line_string_arr(name_index_selection) & "_" & CStr(line_string_arr(1))    ' NAME SHAPE (THIS IS THE CHILD POLYGON
            
        End If
        
        r = r + 1 ' FOR EACH ROW
    Loop
    
    ' CLOSE CSV FILE
    Close #1
    
    
    '''' GROUPING PROCESS
    
    '''' CREATE MAPS META GROUPS
    '''' MAPS META GROUP CONTAINS THE RELATIONSHIP BETWEEN THE PARENT AND CHILD IDS AND NAMES
    '''' EACH ROW REPRESENT ONE GROUP, AND COLUMN C ONWARDS CONTAIN THE CHILDREN NAMES
    is_meta = 0
    '''' CHECK IF MAPS META GROUPS EXIST
    For Each sht In ActiveWorkbook.Sheets
        If sht.Name = "MAPS META GRP" Then
            is_meta = 1
        End If
    Next sht
    
    '''' IF MAPS META DOES NOT EXIST, CREATE IT
    If is_meta = 0 Then
        With ThisWorkbook
            Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            ws.Name = "MAPS META GRP"
            ws.Tab.Color = RGB(10, 10, 0)
        End With
    End If
    
    '''' POPULATE HEADERS
    ThisWorkbook.Sheets("MAPS META GRP").Cells(1, 1) = "GRP ID"
    ThisWorkbook.Sheets("MAPS META GRP").Cells(1, 2) = "GRP NAME"
    ThisWorkbook.Sheets("MAPS META GRP").Cells(1, 3) = "NAMES"
  
    ''' GROUP AND NAME GROUPS OF SHAPES
    
    ' INITIALISATION
    r = 2
    rr = 1
    Dim sp_arr()
    
    '''' FOR EACH ROW IN MAPS META (I.E.EACH CHILD POLYGON)
    While ThisWorkbook.Sheets("MAPS META").Cells(r, 1) <> ""
    
            '''' SUB ID = 0 IS FIRST CHILD OF ANY MULTIPOLYGON
            If ThisWorkbook.Sheets("MAPS META").Cells(r, 3) = 0 Then
                If (r > 2) Then     ' SKIP HEADERS
                    ThisWorkbook.Sheets("MAPs").Activate    ' ACTIVATE MAPS SHEET
                    
                    '''' GROUP SHAPES AND RENAME GROUPS
                    If (UBound(sp_arr) - LBound(sp_arr) + 1) > 1 Then   ' MULTIPOLYGONS
                        ActiveSheet.Shapes.Range(sp_arr).Select             ' SELECT CHILDREN OF SAME GROUPS
                        Selection.Group.Select                              ' GROUP
                        Selection.Name = ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 2)  ' RENAME
                    Else    ' SINGLE POLYGON
                        ActiveSheet.Shapes.Range(sp_arr).Select ' SELECT
                        Selection.Name = ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 2)  ' RENAME AS PARENT NAME
                    End If
                    
                End If
                                
                
                rr = rr + 1     ' EACH ROW
                ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 1) = ThisWorkbook.Sheets("MAPS META").Cells(r, 2)    ' GROUP ID
                ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 2) = ThisWorkbook.Sheets("MAPS META").Cells(r, 4)    ' GROUP NAME
                ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 3) = ThisWorkbook.Sheets("MAPS META").Cells(r, 5)    ' FIRST CHILD POLYGON NAMES
            
            '''' FOR SUBSEQUENT CHILDREN (IF EXIST)
            Else
                ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 3 + ThisWorkbook.Sheets("MAPS META").Cells(r, 3)) = ThisWorkbook.Sheets("MAPS META").Cells(r, 5) ' POPULATE NAME OF CHILD ON THE SAME ROW AS PARENT
            End If
            
            ReDim Preserve sp_arr(ThisWorkbook.Sheets("MAPS META").Cells(r, 3))
            sp_arr(ThisWorkbook.Sheets("MAPS META").Cells(r, 3)) = ThisWorkbook.Sheets("MAPS META").Cells(r, 5) ' POPULATE NAMES OF CHILDREN POLYGONS
            
            
        r = r + 1
        Wend
    
    '''' LAST GROUP
    If (UBound(sp_arr) - LBound(sp_arr) + 1) > 1 Then   ' IF MULTIPOLYGON
        ActiveSheet.Shapes.Range(sp_arr).Select         ' SELECT CHILDREN OF SAME GROUPS
        Selection.Group.Select                          ' GROUP
        Selection.Name = ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 2)  ' RENAME
    Else    ' FOR SINGLE POLYGON
        ActiveSheet.Shapes.Range(sp_arr).Select         ' SELECT
        Selection.Name = ThisWorkbook.Sheets("MAPS META GRP").Cells(rr, 2)  ' RENAME AS PARENT NAME
    End If
    

    '''' DELTE FILE TEMP FILE (NO LONGER REQUIRED)
    Kill get_directory + "\out.csv"
    
    '''' RE-ENABLED SCREEN UPDATING (END OF PROCESS)
    Application.ScreenUpdating = True
    
    End Sub
