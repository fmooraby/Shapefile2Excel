# -*- coding: utf-8 -*-
"""
Created on Fri Sep 18 18:29:14 2020

@author: RMooraby
"""

#### LIBRARIES
import geopandas as gdx
import pandas as pd
import numpy as np
import math as mt
import ntpath as p
import sys

##### FUNCTION TO TRANSLATE SHAPEFLE TO STRING OF X AND Y COORDINATES
#####     fname : file location of shafile
#####     W     : Overall width of Excel Shapefile in px
#####     H     : Overall height of Excel Shapefile in px
#####     X_off : Horizontal offset from top-left corner of Excel Sheet in px
#####     Y_off : Vertical offset from top-left corner of Excel Sheet in px
def translate_shapefile(fname, W, H, X_off, Y_off):

    #### LOAD SHAPEFILE
    fpath, base = p.split(fname) # GET FOLDER OF FILE
    gdf = gdx.read_file(fname)  # LOAD FILE INTO GEOPANDAS

    #### EXPLODE ANY MULTIPOLYGONS TO POLYGONS
    gdf = gdf.to_crs("EPSG:4326")   #SET TO WGS84 PROJECTION
    gdf3 = gdf.explode()            # EXPLODE MULTIPOLYGONS TO POLYGONS

    #### CREATE X, Y COLUMNS
    gdf3['y'] = np.nan
    gdf3['x'] = np.nan

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

    ### FOR EACH ROW IN SHAPEFILE
    for i, r in gdf3.iterrows():
        r['x'] = (X_SCALE * (np.array(r['geometry'].exterior.coords.xy[0], dtype=np.float32) - min_x)) + X_OFFSET # Convert each x coordinates in polygon to excel px value
        r['y'] = (Y_SCALE * (max_y - np.array(r['geometry'].exterior.coords.xy[1], dtype=np.float32))) + Y_OFFSET # Convert each y coordinates in polygon to excel px value

        ##### CREATE STRING LIST OF X AND Y COORDINATES OF POINTS OF POLYGONS
        gdf3['x'][i] = ', '.join(r['x'].astype('str'))
        gdf3['y'][i] = ', '.join(r['y'].astype('str'))

        gdf4['x'][i] = ', '.join(np.array(r['geometry'].exterior.coords.xy[0], dtype=np.float32).astype('str'))
        gdf4['y'][i] = ', '.join(np.array(r['geometry'].exterior.coords.xy[1], dtype=np.float32).astype('str'))

    df1 = pd.DataFrame(gdf3.drop(columns='geometry'))
    df1.to_csv(fname.replace(base, "out.csv"), index_label='ind1, ind2')

#### MAIN OF PYTHON SCRIPT
if __name__ == '__main__':

    # GET INPUT ARGUMENTS
    args = sys.argv[1:]

    W = 800         # DEFAULT WIDTH OF OVERALL EXCEL SHAPE (IN pt)
    H = 600         # DEFAULT HEIGHT OF OVERALL EXCEL SHAPE (IN pt)
    X_off = 50      # DEFAULT HORIZONTAL OFFSET FROM TOP-LEFT CORNER OF EXCEL SHEET
    Y_off = 50      # DEFAULT VERTICAL OFFSET FROM TOP-LEFT CORNER OF EXCEL SHEET
    fname = ""      # DEFAULT FILE PATH IS BLANK and will return error

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

    print("End Process Parameters")

    # check if there is a file name as input
    if fname != "":                                         # yes, there is a file name as input
        print("All good to proceed!")
        translate_shapefile(fname, W, H, X_off, Y_off)      # run process to get shapefile as CSV
    else:
        print("Make sure you enter a valid file path")      # terminate

    print("End of Process")