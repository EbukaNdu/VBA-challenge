# VBA-challenge
for pictures of the workbook solutions look at the issues section

In this repository it shows the answer to my VBA challenge which was challenging to do.

I used the instructors examples like the one from the census_data solution to figure out how to set up my work sheet syntex and determine the last rows like these below
' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        ' INSERT THE YEAR
        ' --------------------------------------------

        ' Create a Variable to Hold File Name, Last Row, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRowa = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        'MsgBox WorksheetName
got help from youtube on how to calculte percentage in excel from "excel campus Jon"

I looked at the instructor's solutions to the other class problems to understand how to run loops through the worksheets and how to set the variables.

finally i crossed checked my work with xpert learning to make sure things were correct and to understand what i was doing.

