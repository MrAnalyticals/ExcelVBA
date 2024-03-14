**Capturing Table object changes using Excdl VBA**

YouTube Video:

**YouTube Video Dialog**

In a worksheet we have a table formatted as a table with one field. If I make a change to the table we can see the VBA triggered displaying a message box. If I make a change elsewhere in the sheet not in the table, we can see the event did not trigger.

In the VBA code you can see that we are making use of the List Objects method as well as the List Columns method. Then we obtain the data body range of that column. this of course does not include the headers. Then, using the intersect function we find out if the editing cell is contained within or does intersect with that previously found data body range. 

**Workbook**: https://github.com/MrAnalyticals/ExcelVBA/blob/master/Table%20Object/TableChangeEvent.xlsm

