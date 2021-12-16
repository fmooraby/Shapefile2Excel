#####################################
#___________________________________#
###         SHAPEFILE2EXCEL       ###
#-----------------------------------#
#####################################


#################################################################################
#   Copyright 2021 Rahman Mohamud Faisal MOORABY                                #
#   Licensed under the Apache License, Version 2.0 (the "License");             #
#   you may not use this file except in compliance with the License.            #
#   You may obtain a copy of the License at                                     #
#       http://www.apache.org/licenses/LICENSE-2.0                              #
#   Unless required by applicable law or agreed to in writing, software         #
#   distributed under the License is distributed on an "AS IS" BASIS,           #
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.    #
#   See the License for the specific language governing permissions and         #
#   limitations under the License.                                              #
#################################################################################

#### LIBRARIES
import geopandas as gdx
import pandas as pd
import numpy as np
import math as mt
import ntpath as p
import sys
import csv

##### FUNCTION TO TRANSLATE SHAPEFLE TO STRING OF X AND Y COORDINATES
#####     fname : file location of shafile
#####     W     : Overall width of Excel Shapefile in px
#####     H     : Overall height of Excel Shapefile in px
#####     X_off : Horizontal offset from top-left corner of Excel Sheet in px
#####     Y_off : Vertical offset from top-left corner of Excel Sheet in px
def translate_shapefile(fname, W, H, X_off, Y_off, simplify):

    #### LOAD SHAPEFILE
    fpath, base = p.split(fname) # GET FOLDER OF FILE
    gdf = gdx.read_file(fname)  # LOAD FILE INTO GEOPANDAS

    #### EXPLODE ANY MULTIPOLYGONS TO POLYGONS
    gdf = gdf.to_crs("EPSG:4326")   #SET TO WGS84 PROJECTION
    gdf3 = gdf.explode()            # EXPLODE MULTIPOLYGONS TO POLYGONS

    #### CREATE X, Y COLUMNS
    gdf3['y'] = ''#np.nan
    gdf3['x'] = ''#np.nan

    #### for diagnosis
    gdf4 = gdf3.copy()

    #### GET BOUNDING BOX PARAMETERS
    min_x = gdf3.total_bounds[0]
    min_y = gdf3.total_bounds[1]
    max_x = gdf3.total_bounds[2]
    max_y = gdf3.total_bounds[3]

    ### GET DIMENSIONS AND OFFSET FROM INPUT
    WIDTH = W
    HEIGHT = H
    X_OFFSET = X_off
    Y_OFFSET = Y_off

    ### CALCULATE RATIO BASED ON EXCEL DIMENSIONS AND OVERALL POLYGON DIMENSION
    X_SCALE = mt.floor(WIDTH / (max_x - min_x))
    Y_SCALE = mt.floor(HEIGHT / (max_y - min_y))

    #### Simplification if simplify is True
    if simplify == True:
        gdf = gdf.simplify(0.05, preserve_topology=False)

    ### FOR EACH ROW IN SHAPEFILE
    for i, r in gdf3.iterrows():
        r['x'] = (X_SCALE * (np.array(r['geometry'].exterior.coords.xy[0], dtype=np.float32) - min_x)) + X_OFFSET # Convert each x coordinates in polygon to excel px value
        r['y'] = (Y_SCALE * (max_y - np.array(r['geometry'].exterior.coords.xy[1], dtype=np.float32))) + Y_OFFSET # Convert each y coordinates in polygon to excel px value
        xstr = ', '.join(str(item) for item in r['x'].tolist())
        ystr = ', '.join(str(item) for item in r['y'].tolist())


        ##### CREATE STRING LIST OF X AND Y COORDINATES OF POINTS OF POLYGONS
        gdf3.at[i, 'x'] = xstr
        gdf3.at[i, 'y'] = ystr

    df1 = pd.DataFrame(gdf3.drop(columns='geometry'))

    #### SAVE TO OUT.CSV. REST OF PROCESS IS THROUGH VBA
    df1.to_csv(fname.replace(base, "out.csv"), index_label = 'ind') ### INDEX_LABEL CONTAINS 'GROUP ID AND SUB ID FOR VBA MACRO
    number_of_rows = len(df1.index)
    number_of_columns = len(df1.columns)

    with open(fname.replace(base, "summary.csv"), 'w', encoding='UTF8') as f:
        writer = csv.writer(f)
        writer.writerow(['row', number_of_rows, 'column', number_of_columns])


#### MAIN OF PYTHON SCRIPT
if __name__ == '__main__':

    # GET INPUT ARGUMENTS
    args = sys.argv[1:]

    W = 800         # DEFAULT WIDTH OF OVERALL EXCEL SHAPE (IN pt)
    H = 600         # DEFAULT HEIGHT OF OVERALL EXCEL SHAPE (IN pt)
    X_off = 50      # DEFAULT HORIZONTAL OFFSET FROM TOP-LEFT CORNER OF EXCEL SHEET
    Y_off = 50      # DEFAULT VERTICAL OFFSET FROM TOP-LEFT CORNER OF EXCEL SHEET
    fname = ""      # DEFAULT FILE PATH IS BLANK and will return error
    simplify = True
    print("Process Parameters")
    print(args)
    for a in args:
        print(a)
        if a[:2].lower() == "w=":       # Get input value for Width
            W = int(a[2:])
            print(a[2:])
        if a[:2].lower() == "h=":       # Get input value for Height
            H = int(a[2:])
            print(a[2:])
        if a[:6].lower() == "x_off=":   # Get input value for horizontal offset
            X_off = int(a[6:])
            print(a[6:])
        if a[:6].lower() == "y_off=":   # Get input value for vertical offset
            Y_off = int(a[6:])
            print(a[6:])
        if a[:5].lower() == "file=":    # Get file path
            fname = a[5:]
            print(a[5:])
        if a[:9].lower() == "simplify=":    # Get file path
            simplify = False if a[9:1]=='F' else True
            print(simplify)
            print(a[9:1])
            print(a[9:])

    print("End Process Parameters")
    print(W)
    print(H)
    print(X_off)
    print(Y_off)
    print(simplify)

    # check if there is a file name as input
    if fname != "":                                         # yes, there is a file name as input
        print("All good to proceed!")
        translate_shapefile(fname, W, H, X_off, Y_off, simplify)      # run process to get shapefile as CSV
    else:
        print("Make sure you enter a valid file path")      # terminate
        fname = r'C:\Users\RMooraby\Downloads\vg-hist.utm32s.shape\vg-hist.utm32s.shape\daten\utm32s\shape\VG-Hist_1990-10-03_RBZ.shp'
        translate_shapefile(fname, W, H, X_off, Y_off, simplify)  # run process to get shapefile as CSV

    print("End of Process")
