::  ''''''''''''''''''''''''''''''''''''
::  ___________________________________'
::  ''         SHAPEFILE2EXCEL       '''
::  -----------------------------------'
::  ''''''''''''''''''''''''''''''''''''
::  
::  
::  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::     Copyright 2021 Rahman Mohamud Faisal MOORABY                                '
::     Licensed under the Apache License, Version 2.0 (the "License");             '
::     you may not use this file except in compliance with the License.            '
::     You may obtain a copy of the License at                                     '
::         http://www.apache.org/licenses/LICENSE-2.0                              '
::     Unless required by applicable law or agreed to in writing, software         '
::     distributed under the License is distributed on an "AS IS" BASIS,           '
::     WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.    '
::     See the License for the specific language governing permissions and         '
::     limitations under the License.                                              '
::  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

:: ::::::::::::::::::::::::::::::::::::::::::::::::::::: ::
::							 ::
::	NOTE: OVERWRITE " [ROOT FOLDER FOR ANACONDA] 	 ::
::		WITH ROOT PATH WHERE ANACONDA IS SAVED	 ::
::							 ::
:: ::::::::::::::::::::::::::::::::::::::::::::::::::::: ::

set arg1=%1
set arg2=%2
set arg3=%3
set arg4=%4
set arg5=%5
set arg6=%6
set arg7=%7

echo %1
echo %2
echo %3
echo %4
echo %5
echo %6
echo %7
@CALL "C:\[ROOT FOLDER FOR ANACONDA]\Anaconda3\Scripts\activate.bat"
python %1 %2 %3 %4 %5 %6 %7