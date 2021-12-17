VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} columns_form 
   Caption         =   "UserForm1"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   OleObjectBlob   =   "columns_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "columns_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub ok_button_Click()
     ' Record selection in Public variable name_index_selection
     name_index_selection = columns_form.ListBox1.ListIndex
     
     ' Hide form on OK click
     columns_form.Hide
End Sub
