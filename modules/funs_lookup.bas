Attribute VB_Name = "funs_lookup"
' Custom functions for looking up values (e.g. VLOOKUP, @INDEX/INDEX, etc.)

' AUTHOR: Robert Schnitman
' Date: 2020-11-12
' Function: INDEXMATCH()
' Description: Simplification of INDEX(..., MATCH(...), MATCH(...)) (Index-Matching).

Function INDEXMATCH(id_lookup As Range, column_lookup As Range, id_range As Range, data_range As Range)
    ' NOTES:
    '   1. data_range and id_range are required to be Ranges so that we can input ranges into this function.
    '   2. id_lookup and column_lookup are declared as Ranges so that we can use cell references.
    '   3. This function assumes that the column headers are in Row 1 of data_range.
    
    ' The Index() function requires an Array as an input; however, we need to be able to select a range of data.
    ' So, we convert the data_range into an Array for Index() to work.
    Dim myarray As Variant
    myarray = data_range
    
    With Application.WorksheetFunction
    
        ' The 2nd and 3rd inputs of Index (row number and column number, respectively) are required to be of Double type.
        Dim r, c As Double
        r = .Match(id_lookup, id_range, 0) ' 0 indicates an exact match.
        c = .Match(column_lookup, data_range.Rows(1), 0) ' ASSUMES THE COLUMN HEADERS ARE IN ROW 1 OF THE DATA RANGE.
    
        ' Output
        INDEXMATCH = .Index(myarray, r, c)
        
        ' Example: I:\Robert\Automation\Excel\Tests\INDEXMATCH.xlsx
        ' =+INDEXMATCH(Test_Data!$A$1:$E$6,Test_Data!$C$1:$C$6,$C2,Test_Data!B$1)
        ' =+INDEXMATCH(data, SSN_column, SSN_cell_to_lookup, column_name_to_lookup)
    
    End With
    
End Function

' AUTHOR: Robert Schnitman
' Date: 2020-11-12
' Function: SLOOKUP()
' Description: Lookup a value based on a row value and column name.

Function SLOOKUP(id_lookup As Range, column_lookup As String, data_range As Range, Optional column_match_type As Integer = 0)
    ' NOTES:
    '   1. All inputs take in Ranges except column_lookup as inputs so that we can use cell references.
    '       1. column_lookup can be a cell reference if and only if the cell reference points to a string.
    '   2. This function assumes that the column headers are in Row 1 of the data range.
    '   3. The match_type input is based on the match type input for MATCH():
    '       1.  1 = Less Than = find the largest value less than or equal to query_column.
    '       2.  0 = Exact     = find the value exactly equal to query_column.
    '       3. -1 = More Than = find the smallest value greater than or equal to query_column.

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