%let pgm=utl-altair-slc-get-the-color-of-a-cell-in-excel-xlsx;

%stop_submission;

Altair slc get the color of a cell in excel xlsx

Too long to post here, see github
https://github.com/rogerjdeangelis/utl-get-the-color-of-a-cell-in-excel-xlsx

Excel input
https://github.com/rogerjdeangelis/utl-altair-slc-get-the-color-of-a-cell-in-excel-xlsx/blob/main/have.xlsx


Get the color of a cell in excel xlsx

 TWO SOLUTIONS

     1 slc proc r
     2 slc proc python


PROBLEM (CELL A2 HAS BACKGROUND RED COLOR)

 1  Create excel sheet with cell A2 collored with Red RGB
    ,EE0044RGB(EE, 0, 44), which converts to decimal RGB(238, 0, 68)
    is a vivid, bright pinkish-red color


    d:/xls/have.xlsx

    -----------------------------------+
    | A1| fx                 |  COLOR  |
    ------------------------------------
    [ ] |       A            |    B    |
    ----------------- -------------------
     1  |      COLOR         |  CODE   |
     -- |--------------------+---------+
     2  |       RED          |   RGB   |
     -- |--------------------+---------+
    [COLORS]


 2  Create python and r code that reads the background color iN cell A2

      Altair SLC
                                  RGB
     Obs    COLOR    CODE    TORCH_RED_CODE

      1      RED     RGB         EE0044


StackOverflow
https://stackoverflow.com/questions/55122922/get-the-color-of-a-cell-from-xlsx-with-python

Sumit Pokhrel
https://stackoverflow.com/users/2690723/sumit-pokhrel

*_                   _
(_)_ __  _ __  _   _| |_
| | '_ \| '_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
;

proc datasets lib=

data workx.colors;
 color="RED";
 code ="RGB";
 output;
run;

%utlfkil(d:/xls/have.xlsx);

ods excel file="d:/xls/have.xlsx" options(sheet_name="colors");;
proc report data=workx.colors;
cols color code;
define color / display;
define code / display;
compute color;
  call define (_col_, "STYLE", "style=[backgroundcolor=CXEE0044]");
endcomp;
run;quit;
ods excel close;

 d:/xls/have.xlsx
 -----------------------------------+
 | A1| fx                 |  COLOR  |
 ------------------------------------
 [ ] |       A            |    B    |
 ----------------- -------------------
  1  |      COLOR         |  CODE   |
  -- |--------------------+---------+
  2  |       RED          |   RGB   |
  -- |--------------------+---------+
 [COLORS]

/*
/ |  _ __  _ __ ___   ___   _ __
| | | `_ \| `__/ _ \ / __| | `__|
| | | |_) | | | (_) | (__  | |
|_| | .__/|_|  \___/ \___| |_|
    |_|
*/

options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
proc r;
submit;
library(xlsx)
library(tidyxl);

file <- "d:/xls/have.xlsx"
rgbcolors <- data.frame(read.xlsx(file, sheetName = "colors",sheetIndex=1,
                header = TRUE, stringsAsFactors = FALSE))
cells <- xlsx_cells(file, sheet = "colors")  # Sheet-specific cells
formats <- xlsx_formats(file)                 # Workbook-wide formats

# A2 is row 3 in your output, local_format_id = 3
a2_format_id <- cells$local_format_id[cells$address == "A2"]  # Returns 3

# Extract fill color (ARGB hex)
torch_red <- formats$local$fill$patternFill$fgColor$rgb[a2_format_id]
torch_red <- substr(torch_red,3,8)
rgbcolors$torch_red_code <- torch_red
print(paste("A2 Torch Red:", torch_red))  # e.g., "FFEE0440"

"final rgbcolors with rgb code"
rgbcolors
rgbcolors<-data.frame(rgbcolors)
endsubmit;
import data=workx.rgbcolors r=rgbcolors;
run;

proc print data=workx.rgbcolors;
run;quit;

/*********************************************/
/*  Altair SLC                               */
/*                                           */
/* Obs    COLOR    CODE    TORCH_RED_CODE    */
/*                                           */
/*  1      RED     RGB         EE0044        */
/*********************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/


NOTE: AUTOEXEC processing completed

1
2         options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
3         proc r;
4         submit;
5         library(xlsx)
6         library(tidyxl);
7
8         file <- "d:/xls/have.xlsx"
9         rgbcolors <- data.frame(read.xlsx(file, sheetName = "colors",sheetIndex=1,
10                        header = TRUE, stringsAsFactors = FALSE))
11        cells <- xlsx_cells(file, sheet = "colors")  # Sheet-specific cells
12        formats <- xlsx_formats(file)                 # Workbook-wide formats
13
14        # A2 is row 3 in your output, local_format_id = 3
15        a2_format_id <- cells$local_format_id[cells$address == "A2"]  # Returns 3
16
17        # Extract fill color (ARGB hex)
18        torch_red <- formats$local$fill$patternFill$fgColor$rgb[a2_format_id]
19        torch_red <- substr(torch_red,3,8)
20        rgbcolors$torch_red_code <- torch_red
21        print(paste("A2 Torch Red:", torch_red))  # e.g., "FFEE0440"
22
23        "final rgbcolors with rgb code"
24        rgbcolors
25        rgbcolors<-data.frame(rgbcolors)
26        endsubmit;

2

NOTE: Using R version 4.5.2 (2025-10-31 ucrt) from C:\Program Files\R\R-4.5.2

NOTE: Submitting statements to R:

> library(xlsx)
> library(tidyxl);
>
> file <- "d:/xls/have.xlsx"
> rgbcolors <- data.frame(read.xlsx(file, sheetName = "colors",sheetIndex=1,
+                 header = TRUE, stringsAsFactors = FALSE))
> cells <- xlsx_cells(file, sheet = "colors")  # Sheet-specific cells
> formats <- xlsx_formats(file)                 # Workbook-wide formats
>
> # A2 is row 3 in your output, local_format_id = 3
> a2_format_id <- cells$local_format_id[cells$address == "A2"]  # Returns 3
>
> # Extract fill color (ARGB hex)
> torch_red <- formats$local$fill$patternFill$fgColor$rgb[a2_format_id]
> torch_red <- substr(torch_red,3,8)
> rgbcolors$torch_red_code <- torch_red
> print(paste("A2 Torch Red:", torch_red))  # e.g., "FFEE0440"
>
> "final rgbcolors with rgb code"
> rgbcolors
> rgbcolors<-data.frame(rgbcolors)

NOTE: Processing of R statements complete

27        import data=workx.rgbcolors r=rgbcolors;
NOTE: Creating data set 'WORKX.rgbcolors' from R data frame 'rgbcolors'
NOTE: Column names modified during import of 'rgbcolors'
NOTE: Data set "WORKX.rgbcolors" has 1 observation(s) and 3 variable(s)

28        run;
NOTE: Procedure r step took :
      real time : 1.805
      cpu time  : 0.015


29
30        proc print data=workx.rgbcolors;
31        run;quit;
NOTE: 1 observations were read from "WORKX.rgbcolors"
NOTE: Procedure print step took :
      real time : 0.000
      cpu time  : 0.000


32
33
34
35
ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 1.883
      cpu time  : 0.062

/*___                                       _   _
|___ \   _ __  _ __ ___   ___   _ __  _   _| |_| |__   ___  _ __
  __) | | `_ \| `__/ _ \ / __| | `_ \| | | | __| `_ \ / _ \| `_ \
 / __/  | |_) | | | (_) | (__  | |_) | |_| | |_| | | | (_) | | | |
|_____| | .__/|_|  \___/ \___| | .__/ \__, |\__|_| |_|\___/|_| |_|
        |_|                    |_|    |___/
*/

options set=PYTHONHOME "D:\py314";
proc python;
submit;
import numpy as np;
import pandas as pd;
import openpyxl;
from openpyxl import load_workbook;
excel_file = 'd:/xls/have.xlsx';
wb = load_workbook(excel_file, data_only = True);
sh = wb['colors'];
color_in_hex = sh['A2'].fill.start_color.index;
rgb=pd.DataFrame(list(color_in_hex[2:]), columns=['rgb']).T;
adx=rgb.iloc[0,0] + rgb.iloc[0,1] + rgb.iloc[0,2] + rgb.iloc[0,3] + rgb.iloc[0,4] + rgb.iloc[0,5];
rgb['col']=adx;
want=pd.DataFrame(rgb['col']);
print(want)
endsubmit;
run;

/*************************/
/*  Altair SLC           */
/*                       */
/* The PYTHON Procedure  */
/*                       */
/*         col           */
/* rgb  EE0044           */
/*************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

NOTE: AUTOEXEC processing completed

1         options set=PYTHONHOME "D:\py314";
2         proc python;
3         submit;
4         import numpy as np;
5         import pandas as pd;
6         import openpyxl;
7         from openpyxl import load_workbook;
8         excel_file = 'd:/xls/have.xlsx';
9         wb = load_workbook(excel_file, data_only = True);
10        sh = wb['colors'];
11        color_in_hex = sh['A2'].fill.start_color.index;
12        rgb=pd.DataFrame(list(color_in_hex[2:]), columns=['rgb']).T;
13        adx=rgb.iloc[0,0] + rgb.iloc[0,1] + rgb.iloc[0,2] + rgb.iloc[0,3] + rgb.iloc[0,4] + rgb.iloc[0,5];
14        rgb['col']=adx;
15        want=pd.DataFrame(rgb['col']);
16        print(want)
17        endsubmit;

NOTE: Submitting statements to Python:


18        run;
NOTE: Procedure python step took :
      real time : 1.151
      cpu time  : 0.015
/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
