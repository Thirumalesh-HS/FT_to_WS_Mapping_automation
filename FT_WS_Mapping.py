# -*- coding: utf-8 -*-
"""

FT_WS_Mapping: Automation of FT failure bins to WS in IC Test Engineering

Copyright (C) 2016  Thirumalesh H S

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.


Contact Info:
-------------
thirumalesh.sreenivasa@tessolve.com or hsthirumalesh@gmail.com


Credits:
-------------
John Machin
Algorithm for retaining cell format while coping between sheets
Link: http://stackoverflow.com/a/5285650/2363712

Psidom
Algorithm for concatenating multiple lists element-wise
Link: http://stackoverflow.com/a/40913016/5283513

"""

# ====== Module for getting current working directory & deleting files ====== #
from os import getcwd, remove

# ==================== Modules for excel data crunching ===================== #
from xlrd import open_workbook
from xlutils.filter import process, XLRDReader, XLWTWriter
from xlwt import Workbook, easyxf
# =========================================================================== #


# =============================== FILES ===================================== #
# give complete path to directory where files are present
DIR = getcwd() + '/'

# input FT/WS xy coord file in csv format
input_files = {
               'WS_01.csv': (
                             'FT_DOE4_S644JL73.csv',
                           ),

              }
# =========================================================================== #


# ========================== Data crunching begins ========================== #
def data_crunch(spreadsheet, mapfile, colormap=[]):
    """
        Input: spreedsheet -> WS or FT xy coords files in csv format
               mapfile     -> xy coords xy_temp file
               colormap    -> None or dictionary of colors mapped to sbin

        Returns: mapped WS/FT xls and new colormap

    """
    # new output filename for saving crunched data in xls format
    out_file = 'out_' + spreadsheet[:-4] + '.xls'

    # process xy coordinates file
    xy = open(DIR + spreadsheet)
    xy_data = xy.readlines()
    xy.close()

    # process sbin, remove empty co-ordinates and newline char
    sbin = xy_data[42]
    sbin = sbin.replace(',', ';').split(';')
    sbin = [s for s in sbin if s]
    sbin[-1] = sbin[-1].strip('\n')

    if spreadsheet.startswith('WS'):

        # process xcoord, ycoord rows
        xcoord, ycoord = xy_data[47], xy_data[48]
        xcoord = xcoord.replace(',', ';').split(';')
        ycoord = ycoord.replace(',', ';').split(';')

        # remove empty co-ordinates
        xcoord = [x for x in xcoord if x]
        ycoord = [y for y in ycoord if y]

        # remove newline char
        xcoord[-1] = xcoord[-1].strip('\n')
        ycoord[-1] = ycoord[-1].strip('\n')

        # assign unique color to each sbin cell
        sbin_unique = sbin[1:]
        sbin_unique = list(set(sbin_unique))
        sbin_unique.sort(key=int)
        sbin_colormap = sbin_color_mapping(sbin_unique, spreadsheet)
    else:
        # executes for FT files, get xcoord and ycoord
        sbin_colormap = colormap
        x, y = [], []
        for line in xy_data:
            if line.startswith('68017') and 'PTR' in line:
                x.append(line)
            elif line.startswith('68018') and 'PTR' in line:
                y.append(line)
        xcoord = get_ft_xy_coord(x)
        ycoord = get_ft_xy_coord(y)

    # process mapping file to out_file
    map_ws = open_workbook(DIR + mapfile, formatting_info=True, on_demand=True)
    map_xcoord = [str(x) for x in range(1, 40)]
    map_ycoord = [str(y) for y in range(42, 0, -1)]

    # copy mapfile contents to newbook retaining formatting
    newbook, newstyle = retain_cell_format_copy(map_ws)
    newsheet = newbook.get_sheet(0)

    # fill sbin details in WS map
    for s, x, y in zip(sbin[1:], xcoord[1:], ycoord[1:]):
        if x == '0' and y =='0':
            pass
        else:
            cell_x, cell_y = (map_ycoord.index(y) + 1), map_xcoord.index(x)
            if spreadsheet.startswith('FT'):

                newsheet.write(cell_x, cell_y, 'FT'+str(s), sbin_colormap.get(s))
            else:
                newsheet.write(cell_x, cell_y, float(s), sbin_colormap.get(s))

    # data crunching saved in out_file
    newbook.save(DIR + out_file)
    map_ws.release_resources
    del map_ws

    return out_file, sbin_colormap
    # =======================  Data crunching Ends ========================== #


# ========================= Extract x y coords of FT ======================== #
def get_ft_xy_coord(xy):
    """
       Input: mulitple xy coords rows of FT files with missing coords in each
              row
       Returns: merged xy coords of mulitple rows

    """
    # split str to lists
    xy_temp = [ each.replace(',', ';').split(';') for each in xy ]
               
    # slice off xy_temp to get only xy-coords
    xy_temp = [ each[7:] for each in xy_temp ]

    # remove newline char
    last_elements = [each[-1].strip('\n') for each in xy_temp]
    for each, last in zip(xy_temp, last_elements):
        each[-1] = last 

    # concatenate multiple lists element-wise
    xycoord = [''.join(x) for x in zip(*xy_temp)]
    return xycoord
# =========================================================================== #


# ========================= color mapping for sbins ========================= #
def sbin_color_mapping(sbin, spreadsheet):

    """
       Input: sbin -> unique set of sbins
       Returns: colormapping for individual sbin

       Only a total of 32 sbins color mapping supported now. But, total of 56
       colors can be given, so max 56 sbins unique color mapping can be done.

       56 Colors supported by xlwt:
       ========================================================================
       # aqua, black, blue, blue_gray, bright_green, brown, coral, cyan_ega,  #
       # dark_blue, dark_blue_ega, dark_green, dark_green_ega, dark_purple,   #
       # dark_red, dark_red_ega, dark_teal, dark_yellow, gold, gray_ega,      #
       # gray25, gray40, gray50, gray80, green, ice_blue, indigo, ivory,      #
       # lavender, light_blue, light_green, light_orange, light_turquoise,    #
       # light_yellow, lime, magenta_ega, ocean_blue, olive_ega, olive_green, #
       # orange, pale_blue, periwinkle, pink, plum, purple_ega, red, rose,    #
       # sea_green, silver_ega, sky_blue, tan, teal, teal_ega, turquoise,     #
       # violet, white, yellow                                                #
       ========================================================================

    """

    print("\n Map colors in order for the unique sbins of %s \n" % spreadsheet)
    print(sbin)
    colors = [
              'yellow', 'aqua', 'blue', 'blue_gray', 'bright_green', 'brown',
              'coral', 'cyan_ega', 'dark_blue', 'dark_green', 'dark_green_ega',
              'dark_purple', 'dark_red', 'dark_yellow', 'gray_ega', 'gray25',
              'gray80', 'indigo', 'ice_blue', 'lavender', 'light_blue',
              'light_orange', 'light_yellow', 'lime', 'pink', 'plum', 'red',
              'rose', 'sea_green', 'sky_blue', 'teal', 'violet',

             ]
    sbin_colormap = {}
    req_colors = colors[:len(sbin)]
    for e, c in zip(sbin, req_colors):
        sbin_colormap[e] = easyxf("""
                                     font: name Calibri, height 160, bold on;
                                     borders: top thin, bottom thin,
                                              left thin, right thin;
                                     pattern: pattern solid, fore_color %s;
                                     alignment: wrap on, horiz centre""" % c
                                  )

    return sbin_colormap
# =========================================================================== #


# =================== generate xy_temp xls file for mapping ==================== #
def gen_map_template(spreadsheet):
    """
       Input: xy_temp xls filename
       Returns: xy_temp xls file filled with X:1 to X:39 and Y:1 to Y:42

    """
    spreadsheet = spreadsheet[:-4] + '.xls'
    map_ws = Workbook()
    map_sheet = map_ws.add_sheet(spreadsheet[:-4])

    # set col width
    col_width = 256 * 5  # 5 characters wide
    for i in range(40):
        map_sheet.col(i).width = col_width

    [map_sheet.write(0, x - 1, 'X:' + str(x)) for x in range(1, 40)]
    [map_sheet.write(43 - y, 39, 'Y:' + str(y)) for y in range(42, 0, -1)]
    map_ws.save(DIR + spreadsheet)

    return spreadsheet
# =========================================================================== #


# ==================== Function to retain cell format ======================= #
def retain_cell_format_copy(wb):
    """
      suggested algorithm by John Machin
      http://stackoverflow.com/a/5285650/2363712

      Input: workbook
      Returns: sheet formatting is retained while copying to new sheet

    """
    w = XLWTWriter()
    process(XLRDReader(wb, 'xy_temp.xls'), w)

    return w.output[0][1], w.style_list
# =========================================================================== #


# ============================= Main Function =============================== #
def main():
    """
       Calls helper functions and delete xy_temp files

    """
    for ws_file, ft_files in input_files.items():
        xytemplate = gen_map_template(ws_file)
        mapfile, colormap = data_crunch(ws_file, xytemplate)

        # process each FT to map with WS
        for each in ft_files:
            data_crunch(each, mapfile, colormap)

        # remove xy_temp created WS files
        remove(DIR + ws_file[:-4] + '.xls')


if __name__ == "__main__":
    main()
# =========================================================================== #


__version__ = "0.1.5"


"""

Rev 0.1   : Initial release Nov 25, 2016
Rev 0.1.1 : Release Nov 28, 2016
            Refactored code for mapping multiple FT files with WS files.
            Removed manual saving of mapfile
Rev 0.1.2 : Release Nov 30, 2016
            Extracts xy coords of FT files without prior Text to Column
            conversion and manual merging of xy coords rows(first 4 rows)
Rev 0.1.3 : Release Dec 1, 2016
            Merging of FT file's mulitple xy coords rows(any no. of rows) 
Rev 0.1.4 : Release Dec 2, 2016
            Fixed bug-error-reading comma separated. 
Rev 0.1.5 : Release Dec 5, 2016
            In certain cases, fusing info of die can be X=Y=0, 
            such coordinates and corresponding sbin are to be skipped.
            Now script is refactored to implement this
            
"""
