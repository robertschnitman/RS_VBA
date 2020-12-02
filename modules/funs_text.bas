Attribute VB_Name = "funs_text"
' Custom Excel functions for handling strings.

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDREPLACE()
' Description: In a cell, replace a string with another.

Function FINDREPLACE(cell As String, string_old As String, string_new As String)
Attribute FINDREPLACE.VB_Description = "Find and replace a character(s)."
Attribute FINDREPLACE.VB_ProcData.VB_Invoke_Func = " \n22"
    
    ' VBA's Replace() is NOT like Excel's REPLACE()!!! It is simpler.
    FINDREPLACE = Replace(cell, string_old, string_new)

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDREMOVE()
' Description: In a cell, remove a specified character(s).

Function FINDREMOVE(cell As String, char As String)
Attribute FINDREMOVE.VB_Description = "Remove a character(s)."
Attribute FINDREMOVE.VB_ProcData.VB_Invoke_Func = " \n22"

    FINDREMOVE = FINDREPLACE(cell, char, "")

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDBEFORE()
' Description: In a cell, return the text before the first specified character(s).

Function FINDBEFORE(cell As String, char As String)
Attribute FINDBEFORE.VB_Description = "Find the substring before a specified character(s)"
Attribute FINDBEFORE.VB_ProcData.VB_Invoke_Func = " \n22"

    ' VBA's Instr() is like Excel's Find().
    char_pos = InStr(cell, char)
    
    ' If char cannot be found, throw an error.
    If char_pos = 0 Then
        
        FINDBEFORE = CVErr(xlErrNA)
        
    ' Otherwise, get everything before the specified character.
    Else
    
        FINDBEFORE = Left(cell, char_pos - 1)
        
    End If

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDAFTER()
' Description: In a cell, return the text after the first specified character(s).

Function FINDAFTER(cell As String, char As String)
Attribute FINDAFTER.VB_Description = "Find the substring after a specified character(s)"
Attribute FINDAFTER.VB_ProcData.VB_Invoke_Func = " \n22"
    
    ' VBA's Instr() is like Excel's Find().
    char_pos = InStr(cell, char)
    
    ' If char cannot be found, throw an error.
    If char_pos = 0 Then

        FINDAFTER = CVErr(xlErrNA)
        
    ' Otherwise, get everything after char.
    Else
    
        FINDAFTER = Mid(cell, char_pos + Len(char)) ' We add Len(char) in case char has multiple characters (e.g. "Robert ").
        
    End If

End Function

' Author: Robert Schnitman
' Date: 2020-11-10
' Function: FINDBETWEEN()
' Description: In a cell, return the text BETWEEN specified characters.

Function FINDBETWEEN(cell As String, char_start As String, char_end As String)
Attribute FINDBETWEEN.VB_Description = "Find the substring between two characters"
Attribute FINDBETWEEN.VB_ProcData.VB_Invoke_Func = " \n22"

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
Attribute FIRSTNAME.VB_Description = "Find the first name of a name string."
Attribute FIRSTNAME.VB_ProcData.VB_Invoke_Func = " \n22"
    ' NOTES:
    '   1. name_order options
    '       1 = First Name Last Name
    '       2 = Last Name, First Name
    '   2. Reverse-order case assumes that there is a comma.
    '   3. Be careful of compound last names (e.g. Del Mul, Van Helsing, etc.)
    
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
        
    ' Reverse order--ASSUMES THAT THERE IS A COMMA.
    ElseIf name_order = 2 Then
        
        Dim out As String
        out = Trim(FINDAFTER(LASTNAME(cell2), ","))
        
        If InStr(out, "Jr ") Or InStr(out, "JR ") Or InStr(out, "I ") Or InStr(out, "i ") Then
        
            FIRSTNAME = Trim(FINDAFTER(out, " "))
            
        Else
        
            FIRSTNAME = Trim(out)
            
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
Attribute LASTNAME.VB_Description = "Find the last name of a name string."
Attribute LASTNAME.VB_ProcData.VB_Invoke_Func = " \n22"
    '   1. name_order options
    '       1 = First Name Last Name
    '       2 = Last Name, First Name
    '   2. Be careful of compound last names (e.g. Del Mul, Van Helsing, etc.)

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

' Author: Robert Schnitman
' Date: 2020-11-16
' Function: TEXTLIKE()
' Description: Determine whether a string meets at least one given pattern.

Function TEXTLIKE(cell As String, ParamArray patterns() As Variant)
Attribute TEXTLIKE.VB_Description = "Detect a pattern-match for a string."
Attribute TEXTLIKE.VB_ProcData.VB_Invoke_Func = " \n22"
    ' ParamArray allows us to give TEXTLIKE() the ability to have multiple inputs without naming them (https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-parameter-arrays).
    
    ' Source of table below: https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/operators-and-expressions/how-to-match-a-string-against-a-pattern
    ' Characters in pattern   Matches in string
    ' ---------------------   -----------------
    ' ?                       Any single character.
    ' *                       Zero or more characters.
    ' #                       Any single digit (0-9).
    ' [ charlist ]            Any single character in charlist.
    ' [ !charlist ]           Any single character not in charlist.
    
    ' e.g TEXTLIKE("Robert Schnitman", "Robert*") ' prints TRUE.
    ' e.g TEXTLIKE("Robert Schnitman", "Craig*", "Robert*") ' prints TRUE.
    
    ' For each given pattern, see if the given string matches any of the specified patterns.
    For Each patt In patterns
    
        ' Does the string match the given pattern?
        detect = cell Like patt
        
        ' If the string matches a specified pattern, exit the loop and use the value in detect;
        '   otherwise, resume the loop until the end.
        ' If the last value is FALSE, then the detect variable will return FALSE.
        If detect = True Then
            
            Exit For
            
        End If
        
    Next
        
    ' The output of the function should be a Boolean value (TRUE/FALSE).
    TEXTLIKE = detect
    
End Function

' Author: Robert Schnitman
' Date: 2020-11-18
' Function: TEXTSTRIPWS()
' Description: Remove all spaces.

Function TEXTSTRIPWS(cell As String)
Attribute TEXTSTRIPWS.VB_Description = "Remove all spaces in a string."
Attribute TEXTSTRIPWS.VB_ProcData.VB_Invoke_Func = " \n22"
    
    TEXTSTRIPWS = FINDREMOVE(cell, " ")

End Function

' Author: Robert Schnitman
' Date: 2020-11-18
' Function: TEXTINSERT()
' Description: Insert a character at a specified position

Function TEXTINSERT(cell As String, char As String, position As Integer)
Attribute TEXTINSERT.VB_Description = "Insert a character at a specified position."
Attribute TEXTINSERT.VB_ProcData.VB_Invoke_Func = " \n22"

    ' The left side of the string should be everything up to just before the specified position.
    sideA = Left(cell, position - 1)
    
    ' The right side should be the concatenation of the specified character to insert AND whatever isn't captured by sideA
    sideB = char + Mid(cell, position)
    
    ' Output should concatenate left and right sides.
    TEXTINSERT = sideA + sideB

End Function

' Author: Robert Schnitman
' Date: 2020-11-18
' Function: TEXTREVERSE()
' Description: Reverse the order of a string.

Function TEXTREVERSE(cell As String)
Attribute TEXTREVERSE.VB_Description = "Reverse the order of a string."
Attribute TEXTREVERSE.VB_ProcData.VB_Invoke_Func = " \n22"

    TEXTREVERSE = StrReverse(cell) ' e.g. TEXTREVERSE("ABCD") = "DCBA"

End Function

' Author: Robert Schnitman
' Date: 2020-11-18
' Function: TEXTCOMPARE()
' Description: Compare two strings. Based on VBA's StrComp().

Function TEXTCOMPARE(string1, string2, Optional compare_type As Long = 1, Optional value As Boolean = False)
Attribute TEXTCOMPARE.VB_Description = "Compare two strings. Based on VBA's StrComp()"
Attribute TEXTCOMPARE.VB_ProcData.VB_Invoke_Func = " \n22"

    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strcomp-function
    ' compare_type = 1 --> Textual Comparison (ABCD = abcd) -- case insensitivity
    ' compare_type = 0 --> Binary Comparison (ABCD > abcd)
    
    ' RESULTS IF value = FALSE
    ' -1 --> string1 < string2
    '  0 --> string1 = string2
    '  1 --> string1 > string2
    
    ' By default, StrComp() outputs an integer.
    comp = StrComp(string1, string2, compare_type)
    
    If value = False Then
    
        output = comp
        
    ' If we want the "translated" value of what the integer means, then output the equivalent string.
    ElseIf value = True Then
    
        Select Case comp
        
            Case -1
                
                output = "<" ' string1 & " < " & string2
                
            Case 0
            
                output = "=" ' string1 & " = " & string2
                
            Case 1
            
                output = ">" ' string1 & " > " & string2
                
        End Select
        
    End If
    
    ' Output the desired value.
    TEXTCOMPARE = output


End Function

' Author: Robert Schnitman
' Date: 2020-11-18
' Function: TEXTJOINR()
' Description: Join a range of strings into a single string, separated by an optional delimiter.

Function TEXTJOINR(string_range As Range, Optional delimiter As String)
Attribute TEXTJOINR.VB_Description = "Join a range of strings into a single string."
Attribute TEXTJOINR.VB_ProcData.VB_Invoke_Func = " \n22"

    TEXTJOINR = Application.WorksheetFunction.TextJoin(delimiter, True, string_range)

End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: TRIML()
' Description: Trim leading spaces.

Function TRIML(cell As String)
Attribute TRIML.VB_Description = "Trim leading spaces."
Attribute TRIML.VB_ProcData.VB_Invoke_Func = " \n22"

    TRIML = LTrim(cell)
    
End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: TRIMR()
' Description: Trim trailing spaces.

Function TRIMR(cell As String)
Attribute TRIMR.VB_Description = "Trim trailing spaces."
Attribute TRIMR.VB_ProcData.VB_Invoke_Func = " \n22"

    TRIMR = RTrim(cell)
    
End Function

' Author: Robert Schnitman
' Date: 2020-11-20
' Function: TRIMLR()
' Description: Remove leading and trailing spaces.

Function TRIMLR(cell As String)
Attribute TRIMLR.VB_Description = "Trim leading and trailing spaces."
Attribute TRIMLR.VB_ProcData.VB_Invoke_Func = " \n22"

    TRIMLR = LTrim(RTrim(cell))

End Function

' Author: Robert Schnitman
' Date 2020-12-01
' Function: RXLIKE()
' Description: Test whether a regular expression pattern has been met.

Function RXLIKE(cell As String, pattern As String, Optional ignore_case As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = ignore_case
        .pattern = pattern
        
    End With
    
    ' Outputs a Boolean value (TRUE/FALSE)
    RXLIKE = regex.Test(cell)

End Function

' Author: Robert Schnitman
' Date 2020-12-01
' Function: RXREPLACE()
' Description: Replace a string based on a regular expression pattern.

Function RXREPLACE(string_old As String, string_pattern As String, string_new As String, Optional ignore_case As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = ignore_case
        .pattern = string_pattern
        
    End With
    
    ' If the string matches the given pattern, replace it with the new string; otherwise, throw an error.
    If regex.Test(string_old) = True Then
    
        output = regex.Replace(string_old, string_new)
        
    Else
    
        output = CVErr(xlErrNA)
        
    End If
    
    RXREPLACE = output

End Function

' Author: Robert Schnitman
' Date 2020-12-01
' Function: RXGET()
' Description: Extract the first text that meets a regular expression pattern.

Function RXGET(cell As String, pattern As String, Optional ignore_case As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = ignore_case
        .pattern = pattern
        
    End With
    
    ' If the string matches the pattern, produce the first match; otherwise, throw an error.
    If regex.Test(cell) = True Then
    
       Set matches = regex.Execute(cell)
       
       output = matches.Item(0)
        
    Else
    
        output = CVErr(xlErrNA)
        
    End If
    
    RXGET = output

End Function

' Author: Robert Schnitman
' Date 2020-12-01
' Function: RXGETALL()
' Description: Extract ALL text that meet a regular expression pattern.

Function RXGETALL(cell As String, pattern As String, Optional sep As String = ",", Optional ignore_case As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = ignore_case
        .pattern = pattern
        
    End With
    
    ' Get all matches.
    Set matches = regex.Execute(cell)
    
    ' https://stackoverflow.com/questions/8146485/returning-a-regex-match-in-vba-excel

    ' Join all matches in a single string, separated by a delimiter ("sep").
    For i = 0 To matches.Count - 1
    
        output = output & sep & matches.Item(i)
        
    Next
    
    ' The concatenation loop above always puts the separator in the first position of the string, so we need to remove it.
    If Len(output) <> 0 Then
    
        output = Right(output, Len(output) - Len(sep))
        
    End If
    
    RXGETALL = output


End Function

' === UNDER CONSTRUCTION === '
' Author: Robert Schnitman
' Date 2020-11-20
' Function: XMLPARSE()
' Description: Search for the value associated with an XML field.

'Function XMLPARSE(cell As String, field As String)

    'Dim tag_begin, tag_end As String
    
    'tag_begin = "<" & field & ">"
    'tag_end = "</" & field & ">"
    
    'output = FINDBETWEEN(cell, field, tag_end)

    'XMLPARSE = FINDREMOVE(output, ">")

'End Function
