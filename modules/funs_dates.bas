Attribute VB_Name = "funs_dates"
' Custom Excel functions for date values.

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: WEEKDAYN()
' Description: Outputs the name of the weekday for a given date.

Function WEEKDAYN(date_cell As Date)
Attribute WEEKDAYN.VB_Description = "Outputs the name of the weekday for a given date."
Attribute WEEKDAYN.VB_ProcData.VB_Invoke_Func = " \n23"
    
    WEEKDAYN = WEEKDAYNAME(date_cell)

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: YMD()
' Description: Formats a date value into the ISO standard format ("yyyy-mm-dd").

Function YMD(date_cell As Date, Optional sep As String = "-")
Attribute YMD.VB_Description = "Formats a date value into the ISO standard format (yyyy-mm-dd)."
Attribute YMD.VB_ProcData.VB_Invoke_Func = " \n23"

    YMD = Format(date_cell, "yyyy" & sep & "mm" & sep & "dd")

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: MDY()
' Description: Formats a date value into the month-day-year order ("mm/dd/yyyy").

Function MDY(date_cell As Date, Optional sep As String = "/")
Attribute MDY.VB_Description = "Formats a date value into the month-day-year order (mm/dd/yyyy)."
Attribute MDY.VB_ProcData.VB_Invoke_Func = " \n23"

    MDY = Format(date_cell, "mm" & sep & "dd" & sep & "yyyy")

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: DMY()
' Description: Formats a date value into the day-month-year order ("dd/mm/yyyy").

Function DMY(date_cell As Date, Optional sep As String = "/")
Attribute DMY.VB_Description = "Formats a date value into the day-month-year order (dd/mm/yyyy)."
Attribute DMY.VB_ProcData.VB_Invoke_Func = " \n23"

    DMY = Format(date_cell, "dd" & sep & "mm" & sep & "yyyy")

End Function

