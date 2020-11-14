Attribute VB_Name = "funs_string"
' Custom Excel functions for handling strings.

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDREPLACE()
' Description: In a cell, replace a string with another.

Function FINDREPLACE(cell As String, string_old As String, string_new As String)
    
    ' VBA's Replace() is NOT like Excel's REPLACE()!!! It is simpler.
    FINDREPLACE = Replace(cell, string_old, string_new)

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDREMOVE()
' Description: In a cell, remove a specified character(s).

Function FINDREMOVE(cell As String, char As String)

    FINDREMOVE = FINDREPLACE(cell, char, "")

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDBEFORE()
' Description: In a cell, return the text before the first specified character(s).

Function FINDBEFORE(cell As String, char As String)

    ' VBA's Instr() is like Excel's Find().
    char_pos = InStr(cell, char)
    
    
    ' Throw an error if Excel cannot find the specified char.
    ' https://www.exceltip.com/custom-functions/return-error-values-from-user-defined-functions-using-vba-in-microsoft-excel.html
    If char_pos = 0 Then
    
        ' FINDBEFORE = CVErr(xlErrNA) ' #N/A error
        
        ' Exit Function
        
        FINDBEFORE = cell
        
    Else
    
        FINDBEFORE = Left(cell, char_pos - 1)
        
    End If

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDAFTER()
' Description: In a cell, return the text after the first specified character(s).

Function FINDAFTER(cell As String, char As String)
    
    ' VBA's Instr() is like Excel's Find().
    char_pos = InStr(cell, char)
    
    ' Throw an error if Excel cannot find the specified char.
    ' https://www.exceltip.com/custom-functions/return-error-values-from-user-defined-functions-using-vba-in-microsoft-excel.html
    If char_pos = 0 Then
    
        ' FINDBEFORE = CVErr(xlErrNA) ' #N/A error
        
        ' Exit Function
        
        FINDAFTER = cell
        
    Else
    
        FINDAFTER = Mid(cell, char_pos + Len(char))
        
    End If

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDBETWEEN()
' Description: In a cell, return the text BETWEEN specified characters.

Function FINDBETWEEN(cell As String, char_start As String, char_end As String)

    ' Where does char_start start?
    num_start = InStr(cell, char_start)
        
    ' Where does char_end start?
    num_end = InStr(cell, char_end)

    ' Throw an error if Excel cannot find the specified characters.
    If num_start = 0 Or num_end = 0 Then
    
        ' https://www.exceltip.com/custom-functions/return-error-values-from-user-defined-functions-using-vba-in-microsoft-excel.html
        FINDBETWEEN = CVErr(xlErrNA) ' #N/A error
        
        Exit Function
        
    Else

        ' To get the text inbetween char_start and char_end, we need to get the positions of when char_start ends and when char_end begins.
        pos_start = num_start + Len(char_start)
        pos_end = num_end - pos_start
        
        FINDBETWEEN = Mid(cell, pos_start, pos_end)
        
    End If

End Function

' Author: Robert Schnitman
' Date: 2020-11-11
' Function: FIRSTNAME()
' Description: Get the first name (and middle name if applicable).

Function FIRSTNAME(cell As String, Optional name_order As Integer = 1)
    ' NOTES:
    '   1. name_order options
    '       1 = First Name Last Name
    '       2 = Last Name, First Name
    
    ' Remove extraneous spaces (left and right sides).
    Dim cell2 As String
    cell2 = Trim(cell) ' Have to name this cell2 because LASTNAME() also uses the "cell" argument and it will "remember" the code in FIRSTNAME().
    
    ' Regular Order
    If name_order = 1 Then
    
        'Remove suffixes
        If InStr(cell2, ",") Then
           
            Dim suffix As String
            suffix = FINDAFTER(cell2, ",")
            
            cell2 = FINDREMOVE(cell, suffix)
                
        ElseIf InStr(cell, " Jr") Then
            
            cell2 = FINDBEFORE(cell2, " Jr")
            
        ElseIf InStr(cell, " I") Then
        
            cell2 = FINDBEFORE(cell2, " I")
                
        End If
    
        ' To get the number of spaces, get the length of the whole cell and subtract the cell without spaces from it.
        ' This is so that we know whether to get the middle name as well.
        len_cell = Len(cell2)
        len_cell_no_spaces = Len(FINDREMOVE(cell2, " "))
        
        len_spaces = len_cell - len_cell_no_spaces
        
        ' In the simple case (e.g. Robert Schnitman), get the text before the space.
        If len_spaces < 2 Then
        
            FIRSTNAME = Trim(FINDBEFORE(cell2, " "))
        
        ' In the complex case (e.g. Robert Gary Schnitman), get the first and middle names separately and before concatenating them together.
        Else
            
            ' Have to use DIM to avoid VBA throwing a compile error.
            Dim first As String
            Dim middle_last As String
            Dim middle As String
            Dim last As String
            
            ' First name is before the first space.
            first = FINDBEFORE(cell2, " ") ' Robert
            
            ' Middle and last names are AFTER the first space.
            middle_last = FINDAFTER(cell2, " ") ' Gary Schnitman, Jr.
            
            ' Middle name is before the space in middle_last
            middle = FINDBEFORE(middle_last, " ") ' Gary
            
            'Last name is after the space after middle name,
            last = FINDAFTER(middle_last, " ")
            
            ' Output should be the concatenation of first and middle names.
            Dim fm As String
            
            fm = first & " " & middle ' Robert Gary
            
            FIRSTNAME = Trim(fm)
            
        End If
        
    ' Reverse order
    ElseIf name_order = 2 Then
        
        Dim out As String
        out = Trim(FINDAFTER(LASTNAME(cell2), ","))
        
        If InStr(out, "Jr ") Or InStr(out, "JR ") Or InStr(out, "I ") Or InStr(out, "i ") Then
        
            FIRSTNAME = Trim(FINDAFTER(out, " "))
            
        Else
        
            FIRSTNAME = Trim(out) ' FINDAFTER(out, " ")
            
        End If
    
    ' Error if name_order is not 1 or 2.
    Else
    
        FIRSTNAME = CVErr(xlErrValue)
        
    End If
    

End Function


' Author: Robert Schnitman
' Date: 2020-11-11
' Function: LASTNAME()
' Description: Get the last name of a person.

Function LASTNAME(cell As String, Optional name_order As Integer = 1)
    '   1. name_order options
    '       1 = First Name Last Name
    '       2 = Last Name, First Name

    ' Regular order
    If name_order = 1 Then
    
        ' Get the first name so that we know the part of the string that's the last name.
        Dim first As String
        first = FIRSTNAME(cell)
        
        ' Anything after the first name is the last name.
        last = FINDAFTER(cell, first)
        
        LASTNAME = Trim(last)
        
    ' Reverse order
    ElseIf name_order = 2 Then
    
        ' Comma situations
        If InStr(cell, ",") Then
    
            ' Get the first name so that we know the part of the string that's the last name.
            Dim first2 As String
            first2 = FIRSTNAME(cell, 2)
            
            ' Remove anything that's a part of the first name.
            Dim last2 As String
            last2 = Trim(FINDREMOVE(cell, first2))
            
            ' Additional comma left at the end behind needs to be removed.
            LASTNAME = Left(last2, Len(last2) - 1)
            
        ' Non-comma situations
        Else
            
            ' Get the first name so that we know the part of the string that's the last name.
            Dim first3 As String
            first3 = FIRSTNAME(cell, 2)
            
            ' Remove anything that's a part of the first name.
            LASTNAME = Trim(FINDREMOVE(cell, first3))
        
        End If
    
    ' Throw a value error if the name_order value is not 1 or 2.
    Else
    
        LASTNAME = CVErr(xlErrValue)
        
    End If

End Function
