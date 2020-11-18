Attribute VB_Name = "funs_lookup"
' Custom functions for looking up values (e.g. VLOOKUP, @INDEX/INDEX, etc.)

' AUTHOR: Robert Schnitman
' Date: 2020-11-12
' Function: SLOOKUP()
' Description: Lookup a value based on a row value and column name.

Function SLOOKUP(id_lookup As String, column_lookup As String, data_range As Range, Optional column_match_type As Integer = 0)
Attribute SLOOKUP.VB_Description = "Lookup a value by row value and column name."
Attribute SLOOKUP.VB_ProcData.VB_Invoke_Func = " \n20"
    ' NOTES:
    '   1. All inputs take in Ranges except column_lookup as inputs so that we can use cell references.
    '       1. column_lookup can be a cell reference if and only if the cell reference points to a string.
    '   2. This function assumes that the column headers are in Row 1 of the data range.
    '   3. The match_type input is based on the match type input for MATCH():
    '       1.  1 = Less Than = find the largest value less than or equal to query_column.
    '       2.  0 = Exact     = find the value exactly equal to query_column.
    '       3. -1 = More Than = find the smallest value greater than or equal to query_column.
    '   4. Can use pattern values.
    '
    '       Source of table below: https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/operators-and-expressions/how-to-match-a-string-against-a-pattern
    '
    '       Characters in pattern   Matches in string
    '       ---------------------   -----------------
    '       ?                       Any single character.
    '       *                       Zero or more characters.
    '       #                       Any single digit (0-9).
    '       [ charlist ]            Any single character in charlist.
    '       [ !charlist ]           Any single character not in charlist.
    '
    '       +SLOOKUP(C2, "Contrib*", A1:K6)

    With Application.WorksheetFunction

        ' The Index() function requires an Array as an input; however, we need to be able to select a range of data.
        ' So, we convert the data_range into an Array for Index() to work.
        Dim myarray As Variant
        myarray = data_range
        

        ' Finding row and column numbers for INDEX().

        ' Set up variables.
        Dim i, col_count As Integer ' For the loop.
        Dim r, c As Double ' Row and column numbers to be obtained from XMATCH().
        
        col_count = data_range.Columns.Count
        
        ' Find row number of id_lookup in data_range
        ' If id_lookup is in column, then use that column and exit loop.
        For i = 1 To col_count
        
            On Error GoTo NextCol: ' Continue to the next column if we get an error.
            
                r = .Match(id_lookup, data_range.Columns(i), 0)
                
NextCol:
            ' This is used to avoid an infinite loop.
            If i <= col_count Then

                Resume NextCol2
                
            Else
            
                ' Throw an error if we cannot find a match in any of the columns of the given data range.
                SLOOKUP = CVErr(xlErrNA)
                
                Exit Function
                
            End If
NextCol2:
        
        Next i
    
        ' Find column number of column_lookup in data_range.
        c = .Match(column_lookup, data_range.Rows(1), column_match_type) ' ASSUMES THE COLUMN HEADERS ARE IN ROW 1 OF THE DATA RANGE.
        
        SLOOKUP = .Index(myarray, r, c)
        
    End With

End Function
