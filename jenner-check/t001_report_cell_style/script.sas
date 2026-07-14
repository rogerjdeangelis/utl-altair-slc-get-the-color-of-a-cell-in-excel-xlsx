/* Derived from utl-altair-slc-get-the-color-of-a-cell-in-excel-xlsx.sas       */
/* The Base SAS core: build the colors table and render it with PROC REPORT,   */
/* setting cell A2's background to the same Torch Red (CXEE0044) the R/Python   */
/* solutions read back out of the workbook. ODS target is a relative file so   */
/* the run is self-contained.                                                  */

data workx_colors;
 color="RED";
 code ="RGB";
 output;
run;

ods html file="report.html";
proc report data=workx_colors;
column color code;
define color / display;
define code / display;
compute color;
  call define (_col_, "STYLE", "style=[backgroundcolor=CXEE0044]");
endcomp;
run;quit;
ods html close;

proc print data=workx_colors;
run;
