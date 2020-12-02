Attribute VB_Name = "funs_logic"
' Custom Excel logic functions

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: ISLEN0()
' Description: Test whether a cell has no characters. Similar to ISBLANK().

Function ISLEN0(cell As String)
Attribute ISLEN0.VB_Description = "Test whether a cell has no characters. Similar to ISBLANK()."
Attribute ISLEN0.VB_ProcData.VB_Invoke_Func = " \n21"

    ' We typically have to test whether there is a blank cell in our client's file, which we do by using LEN(cell) = 0.
    ISLEN0 = (Len(cell) = 0)

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: IFBLANK()
' Description: Similar to IF(), but performs an action depending on whether a cell is blank or not.

Function IFBLANK(cell As String, value_if_true, value_if_false)
Attribute IFBLANK.VB_Description = "Similar to IF(), but performs an action depending on whether a cell is blank or not."
Attribute IFBLANK.VB_ProcData.VB_Invoke_Func = " \n21"

    If ISLEN0(cell) = True Then
    
        output = value_if_true
        
    Else
    
        output = value_if_false
        
    End If
    
    IFBLANK = output

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: SKIPBLANK()
' Description: Perform an action if a cell is non-blank; otherwise, output blank.

Function SKIPBLANK(cell As String, value_if_nonblank)
Attribute SKIPBLANK.VB_Description = "Perform an action if a cell is non-blank; otherwise, output blank."
Attribute SKIPBLANK.VB_ProcData.VB_Invoke_Func = " \n21"

    If ISLEN0(cell) = True Then
    
        output = ""
        
    Else
    
        output = value_if_nonblank
        
    End If
    
    SKIPBLANK = output

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: DOIF()
' Description: Perform an action only if a condition is met; otherwise, output blank.

Function DOIF(condition As Boolean, value_if_true)
Attribute DOIF.VB_Description = "Perform an action only if an action is met; otherwise, output blank."
Attribute DOIF.VB_ProcData.VB_Invoke_Func = " \n21"

    If condition = True Then
    
        output = value_if_true
        
    Else
    
        output = ""
        
    End If
    
    DOIF = output

End Function

' Author: Robert Schnitman
' Date: 2020-11-30
' Function: ISMAC()
' Description: Test whether the user's computer is a Mac.

Function ISMAC()

    If Mac Then
    
        output = True
        
    Else
    
        output = False
        
    End If
    
    ISMAC = output

End Function
