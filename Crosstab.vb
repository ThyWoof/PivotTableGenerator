' Crosstab Excel Builder - version 0.9.0907
'
' AUTHOR: Paulo Monteiro / paulo@roambi.com
'
' free as in beer but don't mess my credits

Option Explicit

' change this to suite your automation needs
Const SCHEDULE_WS = "ROAMBI", GROUP_ROW_WS = "WOLF-"

' should encapsulate these globals that store intermediate phases
Const BREAK = "[(\**/)]"

Const ROW_OUTPUT = 1, ROW_HEADER = 2, ROW_SOURCE = 3, ROW_REPORT = 4
  
Dim rowMin, rowMax, grpMin, grpMax, blkMin, blkMax, grpCnt As Integer

Dim action As String

Dim srcPivot() As String
Dim srcBlock() As String
Dim srcLabel() As String
Dim srcFMask() As String

Dim sqlRow() As String
Dim sqlAgg() As String
 
Dim cnSrcBlock() As Connection
Dim rsSrcBlock() As Recordset

'
' buildOnDemand: creates a Roambi source from the active cell
'

Sub OnDemand()
  
  Dim result As String, RD As Variant
  
  RD = getReportCandidate()
  
  result = parseSQL(RD)
  If result <> "" Then
    MsgBox "PARSER: " & result
    Exit Sub
  End If
  
  result = loadData()
  If result <> "" Then
    MsgBox "LOADER: " & result
    Exit Sub
  End If
  
  result = renderReport()
  If result <> "" Then
    MsgBox "RENDER: " & result
  End If

End Sub

'
' OnSchedule: creates 1 or more Roambi sources in the file system
'

Sub OnSchedule()

  Dim row As Integer, report, result As String, RD As Variant
  
  Worksheets(SCHEDULE_WS).Activate
  
  row = 1
  report = UCase(Left(Cells(row, 1), 5))
  Do
    If report = "TABLE" Or report = "PIVOT" Then
      
      Cells(row, 1).Select
      RD = getReportCandidate()
      
      result = parseSQL(RD)
      If result = "" Then result = loadData()
      If result = "" Then result = renderReport()

      Application.CutCopyMode = False
      Application.DisplayAlerts = False
      If result = "" Then
        ActiveWorkbook.SaveAs ThisWorkbook.Path + "\" + RD(1, 2) + ".xlsx", AccessMode:=xlExclusive, ConflictResolution:=xlLocalSessionChanges
      Else
        ActiveWorkbook.Saved = True
      End If
      ActiveWorkbook.Close
      Application.DisplayAlerts = True
    End If
      
    Worksheets(SCHEDULE_WS).Activate

    Do
      row = row + 1
      If report = "" And UCase(Cells(row, 1)) = "" Then Exit Do
      report = UCase(Left(Cells(row, 1), 5))
    Loop Until report = "TABLE" Or report = "PIVOT"
  Loop Until report = ""
  
End Sub

'
' OnScheduleWithRefresh: Refresh the live data connections and then run the automation
'
Sub OnScheduleWithRefresh()

  ThisWorkbook.RefreshAll
  OnSchedule

End Sub

'
' getReportCandidate: scan southeast from the active cell to find a candidate
'

Function getReportCandidate()
  
  ' scan from the active cell for a report candidate
  Dim baseRow, baseCol, lastRow, lastCol, currRow, currCol As Integer, currVal As String
  
  baseRow = ActiveCell.row
  baseCol = ActiveCell.Column
  
  lastRow = baseRow + 2
  currVal = Left(Cells(baseRow + 1, baseCol), 5)
  Do While currVal <> "" And currVal <> "TABLE" And currVal <> "PIVOT"
    lastRow = lastRow + 1
    currVal = Cells(lastRow + 1, baseCol)
  Loop
  
  lastCol = baseCol
  Do While Cells(baseRow + 1, lastCol + 1) <> ""
    lastCol = lastCol + 1
  Loop
  
  ' copy the definition over
  ReDim RD(1 To lastRow - baseRow + 1, 1 To lastCol - baseCol + 1) As String

  For currRow = 1 To lastRow - baseRow + 1
    For currCol = 1 To lastCol - baseCol + 1
      RD(currRow, currCol) = Cells(baseRow + currRow - 1, baseCol + currCol - 1)
    Next
  Next

  getReportCandidate = RD

End Function

'
' parseSQL: creates the SQL statements from the RD array
'
' RD - Report Definition array
'
' row 1    = PIVOT or TABLE command
' row 2    = headers
' row 3    = group columns and data blocks
' row 4..N = group rows and calculations
'

Function parseSQL(RD As Variant) As String
  
  Dim row, col As Integer
  
  ' regular expression for basic syntax check
  Dim RE As Object, M
  Set RE = CreateObject("vbscript.regexp")
  RE.IgnoreCase = True

  ' validate the layout size
  If UBound(RD, 1) < 4 Or UBound(RD, 2) < 2 Then
    parseSQL = "expected at least a 4 rows x 2 columns report definition"
    Exit Function
  End If

  ' validate the HEADER command
  Dim where As String
  
  RE.Pattern = "^\s*(PIVOT|TABLE)(?:\s+where\s+(.+?))?\s*$"
    
  Set M = RE.Execute(RD(ROW_OUTPUT, 1))
  If M.Count = 0 Then
    parseSQL = "incorrect syntax on cell(1 ,1)"
    Exit Function
  End If
    
  action = UCase(M.Item(0).submatches.Item(0))
  where = M.Item(0).submatches.Item(1)
  
  ' get the group and block boundaries
  rowMin = ROW_REPORT
  rowMax = UBound(RD, 1)
  
  blkMin = 1
  blkMax = UBound(RD, 2)
  
  ' ER for LUCA and date masks on dimensions
  ReDim srcFMask(1 To blkMax) As String
    
  RE.Pattern = "^(?:\""([^\""]+)\"")?$"
  Set M = RE.Execute(RD(ROW_SOURCE, blkMin))
  Do While M.Count > 0
    srcFMask(blkMin) = M.Item(0).submatches.Item(0)
    blkMin = blkMin + 1
    
    ' validate the block count
    If blkMin > blkMax Then
      parseSQL = "expected at least one data block definition on row 3"
      Exit Function
    End If
    Set M = RE.Execute(RD(ROW_SOURCE, blkMin))
  Loop
  
  grpMin = 1
  grpMax = blkMin - 1
  grpCnt = blkMin - grpMin
  
  ' validate the group count
  If grpCnt = 0 Then
    parseSQL = "expected at least one group defintion on row 3"
    Exit Function
  End If
  
  ' get the table, pivot, block, where and format clauses from block cells
  ReDim srcTable(blkMin To blkMax) As String
  ReDim srcPivot(blkMin To blkMax) As String
  ReDim srcBlock(blkMin To blkMax) As String
  ReDim srcWhere(blkMin To blkMax) As String
  
  For col = blkMin To blkMax
    ' [table-name] ( by [field-name] | in "block-name" ) ( where <filter> ) ( format "format-mask" )
    RE.Pattern = "^\s*" & _
      "(\[[^\]]+)\]" & _
      "(?:\s+by\s+(\[[^\]]+\])|\s+in\s+\""([^\""]+)\"")?" & _
      "(?:\s+where\s+(.+?))?" & _
      "(?:\s+format\s+\""([^\""]+)\"")?" & _
      "\s*$"
    
    Set M = RE.Execute(RD(ROW_SOURCE, col))
    If M.Count = 0 Then
      parseSQL = "incorrect syntax on row " & ROW_SOURCE & ", col " & col
      Exit Function
    End If

    ' get the individual source components
    srcTable(col) = M.Item(0).submatches.Item(0) & "$]"
    srcPivot(col) = M.Item(0).submatches.Item(1)
    srcBlock(col) = M.Item(0).submatches.Item(2)
    If where = "" Then
      If M.Item(0).submatches.Item(3) = "" Then
        srcWhere(col) = "true"
      Else
        srcWhere(col) = M.Item(0).submatches.Item(3)
      End If
    Else
      If M.Item(0).submatches.Item(3) = "" Then
        srcWhere(col) = where
      Else
        srcWhere(col) = "(" & where & ") and (" & M.Item(0).submatches.Item(3) & ")"
      End If
    End If
    
    srcFMask(col) = M.Item(0).submatches.Item(4)
    Set M = Nothing
  Next col
  
  ' get the top select list
  Dim selectList As String

  For col = grpMin To grpMax
    If selectList <> "" Then selectList = selectList & ", "
    selectList = selectList & "a.[" & RD(ROW_HEADER, col) & "]"
  Next col
  
  ' get all the headers
  ReDim srcLabel(grpMin To blkMax) As String
  
  For col = grpMin To blkMax
    srcLabel(col) = RD(ROW_HEADER, col)
  Next col
  
  ' get the select, group by, and join clauses
  ReDim columnList(rowMin To rowMax) As String
  ReDim rollupList(rowMin To rowMax) As String
  ReDim joinFields(rowMin To rowMax) As String
  
  For row = rowMin To rowMax
    For col = grpMin To grpMax
      ' ( empty | "const-name" | [field-name] | |jet-SQL| )
      RE.Pattern = "^(\""([^\""]+)\""|(\[[^\]]+\])|\|([^\|]+)\|)?$"
    
      Set M = RE.Execute(RD(row, col))
      If M.Count = 0 Then
        parseSQL = "incorrect syntax at row " & row & ", col " & col
        Exit Function
      End If
      
      ' it is a field
      If M.Item(0).submatches.Item(2) <> "" Then
        If columnList(row) <> "" Then columnList(row) = columnList(row) & ", "
        If rollupList(row) <> "" Then rollupList(row) = rollupList(row) & ", "
        If joinFields(row) <> "" Then joinFields(row) = joinFields(row) & " and "
        columnList(row) = columnList(row) & RD(row, col) & " as [" & RD(ROW_HEADER, col) & "]"
        rollupList(row) = rollupList(row) & RD(row, col)
        joinFields(row) = joinFields(row) & "a.[" & RD(ROW_HEADER, col) & "] = b.[" & RD(ROW_HEADER, col) & "]"
      
      ' it is jet SQL
      ElseIf M.Item(0).submatches.Item(3) <> "" Then
        If columnList(row) <> "" Then columnList(row) = columnList(row) & ", "
        If rollupList(row) <> "" Then rollupList(row) = rollupList(row) & ", "
        If joinFields(row) <> "" Then joinFields(row) = joinFields(row) & " and "
        columnList(row) = columnList(row) & M.Item(0).submatches.Item(3) & " as [" & RD(ROW_HEADER, col) & "]"
        rollupList(row) = rollupList(row) & M.Item(0).submatches.Item(3)
        joinFields(row) = joinFields(row) & "a.[" & RD(ROW_HEADER, col) & "] = b.[" & RD(ROW_HEADER, col) & "]"
            
      ' otherwise a constant or empty
      Else
        If columnList(row) <> "" Then columnList(row) = columnList(row) & ", "
        columnList(row) = columnList(row) & "'" & M.Item(0).submatches.Item(1) & "' as [" & RD(ROW_HEADER, col) & "]"
      End If
      
      Set M = Nothing
    Next col
  Next row
  
  ' create the sql statement to uniquely combine group members from all listed rows and blocks
  ReDim sqlRow(rowMin To rowMax) As String
  ReDim sqlAgg(blkMin - 1 To blkMax) As String
  
  For row = rowMin To rowMax
    If sqlAgg(grpMax) <> "" Then sqlAgg(grpMax) = sqlAgg(grpMax) & vbCrLf & "union "
    For col = blkMin To blkMax
      If sqlRow(row) <> "" Then sqlRow(row) = sqlRow(row) & vbCrLf & "union "
      sqlRow(row) = sqlRow(row) & "select distinct " & columnList(row) & " from " & srcTable(col) & " where " & srcWhere(col)
    Next col
    sqlAgg(grpMax) = sqlAgg(grpMax) & "select * from [" & GROUP_ROW_WS & row & "]"
  Next row
  
  ' create one sql statement per block
  Dim formula As String
  
  ReDim sqlAgg(blkMin To blkMax) As String
  
  For col = blkMin To blkMax
    For row = rowMin To rowMax
      ' [field-name] | count | ( sum | avg | first | last | min | max ) [field-name] | ratio [field-name] (of | from) [field-name] | |jet-SQL|
      RE.Pattern = "^\s*(?:" & _
        "(\""[^\""]+\"")|" & _
        "(\[[^\]]+\])|" & _
        "(count|cnt)|" & _
        "(sum|avg|first|last|min|max)\s+(\[[^\]]+\])|" & _
        "(?:ratio\s+(\[[^\]]+\])\s+(of|from)\s+(\[[^\]]+\]))|" & _
        "(?:\|([^\|]+)\|)" & _
        ")?\s*$"

      Set M = RE.Execute(RD(row, col))
      If M.Count = 0 Then
        parseSQL = "incorrect formula syntax on row " & row & " col " & col
        Exit Function
      End If
      
      ' it is a constant
      If M.Item(0).submatches.Item(0) <> "" Then
        formula = "max(" & M.Item(0).submatches.Item(0) & ")"

      ' it is a field
      ElseIf M.Item(0).submatches.Item(1) <> "" Then
        formula = "sum(" & M.Item(0).submatches.Item(1) & ")"
      
      ' it is a count
      ElseIf M.Item(0).submatches.Item(2) <> "" Then
        formula = "count(*)"
        
      ' it is a unary operator
      ElseIf M.Item(0).submatches.Item(3) <> "" Then
        formula = M.Item(0).submatches.Item(3) & "(" & M.Item(0).submatches.Item(4) & ")"
      
      ' it is a ratio of
      ElseIf UCase(M.Item(0).submatches.Item(6)) = "OF" Then
        formula = "sum(" & M.Item(0).submatches.Item(5) & ") / sum(" & M.Item(0).submatches.Item(7) & ")"

      ' it is a ratio from
      ElseIf UCase(M.Item(0).submatches.Item(6)) = "FROM" Then
        formula = M.Item(0).submatches.Item(7)
        formula = "(sum(" & M.Item(0).submatches.Item(5) & ") - sum(" & formula & ")) / sum(" & formula & ")"
      
      ' it is jet-SQL
      ElseIf M.Item(0).submatches.Item(8) <> "" Then
        formula = M.Item(0).submatches.Item(8)
      
      ' it is empty
      Else
        formula = "sum(null)"
      End If
      
      If sqlAgg(col) <> "" Then sqlAgg(col) = sqlAgg(col) & vbCrLf & "union all "
      
      ' need to consider the pivot field in the select and group by lists
      If srcPivot(col) <> "" Then
        If rollupList(row) = "" Then
          sqlAgg(col) = sqlAgg(col) & "select " & selectList & ", [Pivot], [Formula] from [" & GROUP_ROW_WS & row & "$] a, (" & vbCrLf & _
            "select " & srcPivot(col) & " as [Pivot], " & formula & " as [Formula] " & _
            "from " & srcTable(col) & " where " & srcWhere(col) & vbCrLf & _
            "group by " & srcPivot(col) & vbCrLf & ") b"
        
        Else
          sqlAgg(col) = sqlAgg(col) & "select " & selectList & ", [Pivot], [Formula] from [" & GROUP_ROW_WS & row & "$] a left join (" & vbCrLf & _
            "select " & columnList(row) & ", " & srcPivot(col) & " as [Pivot], " & formula & " as [Formula] " & _
            "from " & srcTable(col) & " where " & srcWhere(col) & vbCrLf & _
            "group by " & rollupList(row) & ", " & srcPivot(col) & vbCrLf & ") b on " & joinFields(row)
        End If

      ' otherwise use a constant as Pivot
      Else
        If rollupList(row) = "" Then
          sqlAgg(col) = sqlAgg(col) & "select " & selectList & ", '" & BREAK & "' as [Pivot], [Formula] from [" & GROUP_ROW_WS & row & "$] a, (" & vbCrLf & _
            "select " & formula & " as [Formula] " & _
            "from " & srcTable(col) & " where " & srcWhere(col) & vbCrLf & _
            ") b"
        
        Else
          sqlAgg(col) = sqlAgg(col) & "select " & selectList & ", '" & BREAK & "' as [Pivot], [Formula] from [" & GROUP_ROW_WS & row & "$] a left join (" & vbCrLf & _
            "select " & columnList(row) & ", " & formula & " as [Formula] " & _
            "from " & srcTable(col) & " where " & srcWhere(col) & vbCrLf & _
            "group by " & rollupList(row) & vbCrLf & ") b on " & joinFields(row)
        End If
      End If
    Next

    ' not SQL-92 compatible but might be easily replaced with SQL Server ROLLUP / PIVOT if a real DB is required
    sqlAgg(col) = "TRANSFORM max([Formula]) SELECT " & selectList & " FROM (" & vbCrLf & _
      sqlAgg(col) & vbCrLf & ") a GROUP BY " & selectList & " PIVOT [Pivot]"
  Next
  
  Set RE = Nothing
    
End Function

'
' loadData
'

Function loadData() As String

  Dim row, col, i As Integer

  Application.CutCopyMode = False
  Application.DisplayAlerts = False
  
  ' create temporary sheets to optimize the data retrieval (using recordset instead of SELECT INTO to avoid memory leak in Jet)
  Dim cnGroup As New Connection, rsGroup As New Recordset
  
  ' opening / closing connections within the loop to avoid a weird bug
  For row = rowMax To rowMin Step -1
    On Error Resume Next
    ThisWorkbook.Worksheets(GROUP_ROW_WS & row).Delete
    On Error GoTo 0
    
    ThisWorkbook.Worksheets.Add
    ThisWorkbook.ActiveSheet.Name = GROUP_ROW_WS & row
    
    On Error GoTo DATA_CONN_ERROR_GROUP
    cnGroup.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 8.0;Data Source=" & ThisWorkbook.FullName
    
    On Error GoTo DATA_ACCESS_ERROR_GROUP
    rsGroup.Open sqlRow(row), cnGroup, adOpenForwardOnly
    On Error GoTo 0
    
    ThisWorkbook.Worksheets(GROUP_ROW_WS & row).Cells(2, 1).CopyFromRecordset rsGroup
    
    ' apply the dimension format
    For col = 1 To grpMax
      ThisWorkbook.Worksheets(GROUP_ROW_WS & row).Columns(col).NumberFormat = srcFMask(col)
    Next col
    
    ' copy the header
    For col = 0 To rsGroup.Fields.Count - 1
      ThisWorkbook.Worksheets(GROUP_ROW_WS & row).Cells(1, col + 1) = rsGroup.Fields(col).Name
    Next col
    
    rsGroup.Close
    cnGroup.Close
    Set rsGroup = Nothing
    Set cnGroup = Nothing
  Next row
  
  ' retrieve each column group data (this is dependendent on temporary sheets created above)
  ReDim cnSrcBlock(blkMin To blkMax) As Connection
  ReDim rsSrcBlock(blkMin To blkMax) As Recordset

  For col = blkMin To blkMax
    Set cnSrcBlock(col) = New Connection
    Set rsSrcBlock(col) = New Recordset
    
    On Error GoTo DATA_CONN_ERROR_BLOCK
    cnSrcBlock(col).Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 8.0;Data Source=" & ThisWorkbook.FullName
    
    On Error GoTo DATA_ACCESS_ERROR_BLOCK

    rsSrcBlock(col).Open sqlAgg(col), cnSrcBlock(col), adOpenStatic
    On Error GoTo 0
  Next
  
  ' remove the temporary sheets
  On Error Resume Next
  For row = rowMax To rowMin Step -1
    ThisWorkbook.Worksheets(GROUP_ROW_WS & row).Delete
  Next row
  On Error GoTo 0

  Application.DisplayAlerts = True
  loadData = ""
  Exit Function
  
DATA_CONN_ERROR_GROUP:
  Set cnGroup = Nothing
  
  Application.DisplayAlerts = True
  loadData = "Connection error before group fetch"
  Exit Function
  
DATA_ACCESS_ERROR_GROUP:
  Set rsGroup = Nothing
  
  cnGroup.Close
  Set cnGroup = Nothing
  
  Application.DisplayAlerts = True
  loadData = "Error fetching group on row " & row
  Exit Function
  
DATA_CONN_ERROR_BLOCK:
  For i = blkMin To col - 1
    rsSrcBlock(i).Close
    Set rsSrcBlock(i) = Nothing
    
    cnSrcBlock(i).Close
    Set cnSrcBlock(i) = Nothing
  Next i
  
  Set cnSrcBlock(col) = Nothing
  
  Application.DisplayAlerts = True
  loadData = "Connection error before block fetch"
  Exit Function
  
DATA_ACCESS_ERROR_BLOCK:
  For i = blkMin To col - 1
    rsSrcBlock(i).Close
    Set rsSrcBlock(i) = Nothing
    
    cnSrcBlock(i).Close
    Set cnSrcBlock(i) = Nothing
  Next i
  
  Set rsSrcBlock(col) = Nothing
  
  cnSrcBlock(col).Close
  Set cnSrcBlock(col) = Nothing
  
  Application.DisplayAlerts = True
  loadData = "Error fetching block on column " & col
  Exit Function
  
End Function

'
' renderReport
'

Function renderReport()
     
  Const HEADER_COLOR = 45

  Dim GROUP_COLOR
  
  GROUP_COLOR = Array(48, 16, 15)
  
  Dim row, col, rowRS, colRS, rowPos, colPos, i As Integer
  Dim prevBlock As String
  Dim hasNullPivot As Boolean

  Application.CutCopyMode = False
  Application.DisplayAlerts = False
  Workbooks.Add.Activate
  
  ' action determines offset
  rowPos = 1 - 1 * (action = "PIVOT")
  colPos = 1
  
  ' paste each block data and keep a tab on the previous block name
  prevBlock = BREAK
  For col = blkMin To blkMax
    ' check if a NULL pivot is present
    hasNullPivot = (rsSrcBlock(col).Fields(grpCnt).Name = "<>")
    
    ' copy the data over
    rowRS = rsSrcBlock(col).RecordCount
    colRS = rsSrcBlock(col).Fields.Count + hasNullPivot
    Cells(rowPos + 1, colPos).CopyFromRecordset rsSrcBlock(col)
    
    ' discard the NULL pivot column
    If hasNullPivot Then
      Range(Cells(1, colPos + grpCnt), Cells(rowPos + rowRS, colPos + grpCnt)).Delete
    End If
    
    ' copy the header
    For i = colRS + 1 * Not hasNullPivot To 0 Step -1
      If i < grpCnt Then
        Cells(rowPos, colPos + i).Value = rsSrcBlock(col).Fields(i).Name

      ElseIf srcPivot(col) = "" Then
        Cells(rowPos, colPos + i + hasNullPivot).Value = srcLabel(col)
        
      ElseIf action = "PIVOT" Then
        Cells(rowPos, colPos + i + hasNullPivot).Value = rsSrcBlock(col).Fields(i).Name
        
      ElseIf action = "TABLE" Then
        Cells(rowPos, colPos + i + hasNullPivot).Value = srcLabel(col) & " " & rsSrcBlock(col).Fields(i).Name
      End If
    Next i

    ' set the header description, merge header cells, and apply a thin border around the block header
    If action = "PIVOT" Then
    
      ' get a border around the data
      applyThinBorder Range(Cells(rowPos + 1, colPos + grpCnt), Cells(rowPos + rowRS, colPos + colRS - 1))
 
      ' it is a pivot so give it a name, merge, paint, and apply border
      If srcPivot(col) <> "" Then
        Cells(rowPos - 1, colPos + grpCnt) = srcLabel(col)
        Range(Cells(rowPos - 1, colPos + grpCnt), Cells(rowPos - 1, colPos + colRS - 1)).Merge
        applyThinBorder Range(Cells(rowPos - 1, colPos + grpCnt), Cells(rowPos - 1, colPos + colRS - 1))
        applyThickBorder Range(Cells(rowPos - 1, colPos + grpCnt), Cells(rowPos, colPos + colRS - 1))
        Range(Cells(rowPos - 1, colPos + grpCnt), Cells(rowPos, colPos + colRS - 1)).Interior.ColorIndex = HEADER_COLOR
        prevBlock = BREAK
      
      ' open a new static block
      ElseIf srcBlock(col) <> prevBlock Then
        Cells(rowPos - 1, colPos + grpCnt) = srcBlock(col)
        applyThinBorder Cells(rowPos - 1, colPos + grpCnt)
        applyThickBorder Cells(rowPos, colPos + grpCnt)
        Range(Cells(rowPos - 1, colPos + grpCnt), Cells(rowPos, colPos + grpCnt)).Interior.ColorIndex = HEADER_COLOR
        prevBlock = srcBlock(col)
        
      ' merge static columns under the same block
      Else
        Range(Cells(rowPos - 1, colPos - 1).MergeArea.Cells(1, 1), Cells(rowPos - 1, colPos + grpCnt)).Merge
        applyThinBorder Range(Cells(rowPos - 1, colPos - 1).MergeArea.Cells(1, 1), Cells(rowPos - 1, colPos + grpCnt))
        applyThickBorder Cells(rowPos, colPos + grpCnt)
        Cells(rowPos, colPos + grpCnt).Interior.ColorIndex = HEADER_COLOR
      End If
    End If
  
    ' apply the number format
    Range(Cells(rowPos + 1, colPos + grpCnt), Cells(rowPos + rowRS, colPos + colRS - 1)).NumberFormat = srcFMask(col)
    
    ' set position for next iteration and remove row members from 2nd iteration onward
    If col = blkMin Then
      colPos = colPos + colRS
    Else
      Range(Cells(1, colPos), Cells(rowPos + rowRS, colPos + grpCnt - 1)).Delete
      colPos = colPos + colRS - grpCnt
    End If
    
    ' clean the house
    rsSrcBlock(col).Close
    Set rsSrcBlock(col) = Nothing
    
    cnSrcBlock(col).Close
    Set cnSrcBlock(col) = Nothing
  Next col
  
  ' apply the dimension format (again?)
  For col = 1 To grpMax
    Columns(col).NumberFormat = srcFMask(col)
  Next col
  
  ' get a border around the aggregation rows
  applyThinBorder Range(Cells(1, 1), Cells(rowPos + rowRS, grpCnt))
 
  ' finalize the table header
  If action = "TABLE" Then
    Range(Cells(1, 1), Cells(1, colPos - 1)).Interior.ColorIndex = HEADER_COLOR
    applyThinBorder Range(Cells(2, grpCnt + 1), Cells(rowPos + rowRS, colPos - 1))
    applyThickBorder Range(Cells(1, 1), Cells(1, colPos - 1))
    
  ' or add a border to the top left pivot area
  Else
    applyThickBorder Range(Cells(1, 1), Cells(2, grpCnt))
  End If
   
  ' only merge the groups in a pivot
  If action = "PIVOT" Then
    Dim top As Integer, prevMember, currMember As String
        
    ' moving backwards to guarantee parent total row has same color as parent merge group
    For col = grpCnt - 1 To grpMin Step -1
      row = rowPos
      top = row + 1
      prevMember = getGroupRowBreak(top, col)
      
      Do
        row = row + 1
        currMember = getGroupRowBreak(row, col)
        
        If currMember <> prevMember Then
          ' merge and paint the current row block
          With Range(Cells(top, col), Cells(row - 1, col))
            .Merge
            .VerticalAlignment = xlTop
            .Interior.ColorIndex = GROUP_COLOR(col \ 3)
          End With
        
          ' give a nice thick underline touch
          With Range(Cells(row - 1, col), Cells(row - 1, grpCnt)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
          End With
        
          ' paint first child row all the way to header end if a total level
          If Cells(top, col + 1).Value = "" Then
            Range(Cells(top, col + 1), Cells(top, grpCnt)).Interior.ColorIndex = GROUP_COLOR(col \ 3)
          End If
        
          top = row
          prevMember = currMember
        End If
      Loop While row <= rowPos + rowRS
    Next col
  End If

  ' final OCD
  Cells.Select
  Cells.EntireColumn.AutoFit
  Cells(1, 1).Select
  
  Application.DisplayAlerts = True
  
End Function

'
' getGroupRowBreak - get the concatenation of all cells from column 1 to column col at specified row
'

Function getGroupRowBreak(row, col)

  Dim i As Integer, result As String
  
  For i = 1 To col
    result = result & Cells(row, i)
  Next i
        
  getGroupRowBreak = result
  
End Function

'
' auxiliary routines
'

Sub applyThinBorder(r As Range)

  With r
    .Borders(xlEdgeTop).Weight = xlThin
    .Borders(xlEdgeLeft).Weight = xlThin
    .Borders(xlEdgeRight).Weight = xlThin
    .Borders(xlEdgeBottom).Weight = xlThin
  End With
  
End Sub

Sub applyThickBorder(r As Range)

  With r
    .Borders(xlEdgeTop).Weight = xlThin
    .Borders(xlEdgeLeft).Weight = xlThin
    .Borders(xlEdgeRight).Weight = xlThin
    .Borders(xlEdgeBottom).Weight = xlThick
  End With
  
End Sub
