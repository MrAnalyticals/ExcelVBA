**Capturing Table object changes using Excel VBA**

![image](https://github.com/MrAnalyticals/ExcelVBA/assets/47678539/3a54a924-28e2-4f67-b603-2521170fe28e)

YouTube Video:https://youtu.be/QAqW5U-v1WY

**YouTube Video Dialog**

In a worksheet we have a table formatted as a table with one field. If I make a change to the table we can see the VBA triggered displaying a message box. If I make a change elsewhere in the sheet not in the table, we can see the event did not trigger.

In the VBA code you can see that we are making use of the List Objects method as well as the List Columns method. Then we obtain the data body range of that column. this of course does not include the headers. Then, using the intersect function we find out if the editing cell is contained within or does intersect with that previously found data body range. 

**Workbook**: https://github.com/MrAnalyticals/ExcelVBA/blob/master/Table%20Object/TableChangeEvent.xlsm

