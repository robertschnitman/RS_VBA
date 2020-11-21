Attribute VB_Name = "funs_dates"
' Custom Excel functions for date values.

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: WEEKDAYNAME()
' Description: Outputs the name of the weekday for a given date.

Function WEEKDAYNAME(date_cell As Date)
Attribute WEEKDAYNAME.VB_Description = "Outputs the name of the weekday for a given date."
Attribute WEEKDAYNAME.VB_ProcData.VB_Invoke_Func = " \n25"

    wday = Weekday(date_cell, vbSunday) ' 1 = Sunday
    
    Select Case wday
    
        Case 1
        
            output = "Sunday"
            
        Case 2
        
            output = "Monday"
            
        Case 3
            
            output = "Tuesday"
            
        Case 4
        
            output = "Wednesday"
            
        Case 5
        
            output = "Thursday"
            
        Case 6
        
            output = "Friday"
            
        Case 7
        
            output = "Saturday"
            
    End Select
    
    WEEKDAYNAME = output

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: YMD()
' Description: Formats a date value into the ISO standard format ("yyyy-mm-dd").

Function YMD(date_cell As Date)
Attribute YMD.VB_Description = "Formats a date value into the ISO standard format (yyyy-mm-dd)."
Attribute YMD.VB_ProcData.VB_Invoke_Func = " \n25"

    YMD = Format(date_cell, "yyyy-mm-dd")

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: MDY()
' Description: Formats a date value into the month-day-year order ("mm/dd/yyyy").

Function MDY(date_cell As Date)
Attribute MDY.VB_Description = "Formats a date value into the month-day-year order (mm/dd/yyyy)."
Attribute MDY.VB_ProcData.VB_Invoke_Func = " \n25"

    MDY = Format(date_cell, "mm/dd/yyyy")

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: DMY()
' Description: Formats a date value into the day-month-year order ("dd/mm/yyyy").

Function DMY(date_cell As Date)
Attribute DMY.VB_Description = "Formats a date value into the day-month-year order (dd/mm/yyyy)."
Attribute DMY.VB_ProcData.VB_Invoke_Func = " \n25"

    DMY = Format(date_cell, "dd/mm/yyyy")

End Function

