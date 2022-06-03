' this code is a class module for a standard "table" in Excel. This was created to help navigate tables in Excel a lot faster and easier.
Option Explicit
Public Specs As w1DArr: Public ColName As w1DArr: Public Book As Workbook
Public ShName As String: Public ShPwd As String
Public TopLeft As cFind: Public DataOffset As Long

Public Property Get Sh() As Worksheet
    If Book Is Nothing Then Set Sh = ThisWorkbook.Sheets(ShName) Else Set Sh = Book.Sheets(ShName)
End Property

Public Property Get HeadFirstRow() As Long: HeadFirstRow = Specs(0): End Property
Public Property Get FirstRow() As Long:     FirstRow = Specs(1):     End Property
Public Property Get LastRow() As Long:      LastRow = Specs(2):      End Property
Public Property Get FirstCol() As Long:     FirstCol = Specs(3):     End Property
Public Property Get LastCol() As Long:      LastCol = Specs(4):      End Property

Public Sub Init(vShName As String, vShPwd As String _
              , vdfType, vRoot As String, vStem As String, vDataOffset As Long _
              , Optional LeavedTbl As Boolean = True, Optional CheckLastRow As Long = 1048576 _
              , Optional vBook As Workbook)
    Set Specs = New w1DArr: Set ColName = New w1DArr: Set TopLeft = New cFind
    ShName = vShName: ShPwd = vShPwd: DataOffset = vDataOffset: If Not vBook Is Nothing Then Set Book = vBook
    PrepareSheet
    TopLeft.Init vdfType, vRoot, vStem: Specs.ReDimArr 0, 4
    If TopLeft(Sh) = "" Then: Exit Sub
    With Sh.Range(TopLeft(Sh))
        Specs.SetVal 0, .Row
        Specs.SetVal 3, .Column
        Specs.SetVal 4, IIf(.Offset(0, 1).Value = "", .Column, .End(xlToRight).Column)
        Specs.SetVal 1, .Offset(DataOffset, 0).Row
        If LeavedTbl Then
            If .Offset(DataOffset).Value = "" Then
                Specs.SetVal 2, Specs(1)
            ElseIf .Offset(DataOffset + 1).Value = "" Then
                Specs.SetVal 2, Specs(1)
            Else
                Specs.SetVal 2, IIf(Abs(.Offset(DataOffset, 0).End(xlDown).Row - CheckLastRow) <= 1 _
                                      , FirstRow, .Offset(DataOffset, 0).End(xlDown).Row)
            End If
        Else
            Dim vCol As Long: For vCol = FirstCol To LastCol
                Dim vLastRow As Long: vLastRow = IIf(Abs(.Offset(DataOffset, vCol - FirstCol).End(xlDown).Row - CheckLastRow) <= 1 _
                                                   , HeadFirstRow, .Offset(DataOffset, vCol - FirstCol).End(xlDown).Row)
                Dim vDataLastRow As Long: vDataLastRow = IIf(vLastRow <= vDataLastRow, vDataLastRow, vLastRow)
            Next vCol
            Specs.SetVal 2, vDataLastRow
        End If
    End With
    For vCol = FirstCol To LastCol
        Dim ColNameStr As String: ColNameStr = Sh.Cells(HeadFirstRow, vCol).Value
        If ColName.IsIn(ColNameStr) Then
            Dim vCoin As Long: vCoin = 1
            Do While ColName.IsIn(ColNameStr & "_" & vCoin)
                vCoin = vCoin + 1
            Loop
            ColNameStr = ColNameStr & "_" & vCoin
        End If
        ColName.Add ColNameStr
    Next vCol
End Sub

Public Sub InitStd(vTblHead As String, vShName As String, vShPwd As String)
    Init vShName, vShPwd, 3, vTblHead, "1>0", 1, True
End Sub

Public Property Get xIsEmpty() As Boolean
    xIsEmpty = False
    If LastRow = FirstRow And Value(FirstRow, FirstCol) = "" Then xIsEmpty = True
End Property

Public Property Get Tag() As String
    Tag = Sh.Cells(HeadFirstRow, FirstCol).Offset(-1, 0)
End Property

Public Property Get Row(vRowInd As Long) As w1DArr
    If Not (vRowInd >= FirstRow And vRowInd <= LastRow) Then Exit Property
    Set Row = New w1DArr
    Dim vCol As Long: For vCol = FirstCol To LastCol
        Row.Add IIf(IsError(Sh.Cells(vRowInd, vCol).Value), "", Sh.Cells(vRowInd, vCol).Value)
    Next vCol
End Property

Public Property Get Column(vCol, Optional RemoveDup As Boolean = False) As w1DArr
    Set Column = New w1DArr
    Dim vRowInd As Long: For vRowInd = FirstRow To LastRow
        If Not (RemoveDup And Column.IsIn(Sh.Cells(vRowInd, Enc(vCol) + FirstCol).Value)) Then Column.Add Sh.Cells(vRowInd, Enc(vCol) + FirstCol).Value
    Next vRowInd
End Property

Public Property Get ColRng(vCol) As Range
    Set ColRng = Sh.Range(Sh.Cells(FirstRow, FirstCol + Enc(vCol)), Sh.Cells(LastRow, FirstCol + Enc(vCol)))
End Property

Public Property Get RowSlice(vRowInd As Long, vCol, Optional vEdge As Long = ArrEdge.TopEdge) As w1DArr
    If Not (vEdge = ArrEdge.BottomEdge Or vEdge = ArrEdge.TopEdge) Then Exit Property
    Set RowSlice = Row(vRowInd).Slice(Enc(vCol), vEdge)
End Property

Public Property Get RowSegment(vRowInd As Long, vCol1, vCol2) As w1DArr
    Set RowSegment = Row(vRowInd).Segment(Enc(vCol1), Enc(vCol2))
End Property

Public Property Get Block(Optional vFilterCol = "", Optional vFilterVal As String = "", Optional vCol1 = "", Optional vCol2 = "", Optional vEdge As Long = ArrEdge.TopEdge) As w2DArr
    Set Block = New w2DArr
    If vFilterCol = "" Then
        If vCol1 & vCol2 = "" Then
            Block.Init Sh.Range(Sh.Cells(FirstRow, FirstCol), Sh.Cells(LastRow, LastCol)).Value
        ElseIf w(Array(vCol1, vCol2)).IsIn("") Then
            If vEdge = ArrEdge.TopEdge Then
                Block.Init Sh.Range(Sh.Cells(FirstRow, Enc(vCol1 & vCol2) + FirstCol), Sh.Cells(LastRow, LastCol)).Value
            Else
                Block.Init Sh.Range(Sh.Cells(FirstRow, FirstRow), Sh.Cells(LastRow, Enc(vCol1 & vCol2) + FirstCol)).Value
            End If
        Else
            Block.Init Sh.Range(Sh.Cells(FirstRow, FirstCol + IIf(Enc(vCol1) < Enc(vCol2), Enc(vCol1), Enc(vCol2))), Sh.Cells(LastRow, FirstCol + IIf(Enc(vCol1) < Enc(vCol2), Enc(vCol2), Enc(vCol1)))).Value
        End If
    Else
        Dim Ind As Long: For Ind = FirstRow To LastRow
            If vFilterCol = "" Or Value(Ind, vFilterCol) = vFilterVal Then
                If vCol1 & vCol2 = "" Then
                    Block.AddLine Row(Ind)
                Else
                    If w(Array(vCol1, vCol2)).IsIn("") Then
                        Dim vCol: vCol = vCol1 & vCol2
                        Block.AddLine RowSlice(Ind, vCol, vEdge)
                    Else
                        Block.AddLine RowSegment(Ind, vCol1, vCol2)
                    End If
                End If
            End If
        Next Ind
    End If
End Property

Public Property Get TblRange(Optional IncludeTag As Boolean = False) As Range
    Set TblRange = Sh.Range(Sh.Cells(HeadFirstRow, FirstCol), Sh.Cells(LastRow, LastCol))
    If IncludeTag Then
        Set TblRange = Union(TblRange, Sh.Cells(HeadFirstRow, FirstCol).Offset(-1, 0))
    End If
End Property

Public Property Get DataRange() As Range
    Set DataRange = Sh.Range(Sh.Cells(FirstRow, FirstCol), Sh.Cells(LastRow, LastCol))
End Property

Public Property Get HeaderRange() As Range
    Set HeaderRange = Sh.Range(Sh.Cells(HeadFirstRow, FirstCol), Sh.Cells(HeadFirstRow, LastCol))
End Property

Public Property Get FullRange() As Range
    Set FullRange = Sh.Range(Sh.Cells(HeadFirstRow, FirstCol), Sh.Cells(LastRow, LastCol))
End Property

Public Property Get Value(vRowInd As Long, vCol) As String
    Value = Row(vRowInd)(Enc(vCol))
End Property

Public Sub SetVal(vSetVal As String, vRowInd As Long, vCol)
    If Not (vRowInd >= FirstRow And vRowInd <= LastRow) Then Exit Sub
    Sh.Cells(vRowInd, Enc(vCol) + FirstCol).Value = vSetVal
End Sub

Public Property Get Implode(vRowInd1 As Long, vRowInd2 As Long, vCol1, vCol2 _
                          , Optional rDelim As String = vbCrLf, Optional fDelim As String = "|")
    If Not (vRowInd1 >= FirstRow And vRowInd1 <= LastRow) Then Exit Property
    If Not (vRowInd2 >= FirstRow And vRowInd2 <= LastRow) Then Exit Property
    Dim vRowIndAdj1 As Long, vRowIndAdj2 As Long
    vRowIndAdj1 = IIf(vRowInd1 <= vRowInd2, vRowInd1, vRowInd2)
    vRowIndAdj2 = IIf(vRowInd1 <= vRowInd2, vRowInd2, vRowInd1)
    Dim Temp As w1DArr: Set Temp = New w1DArr
    Dim vRow As Long: For vRow = vRowIndAdj1 To vRowIndAdj2
        Temp.Add RowSegment(vRow, vCol1, vCol2).Implode(fDelim)
    Next vRow
    Implode = Temp.Implode(rDelim): Set Temp = Nothing
End Property

Public Function Enc(vIn) As Long
    Enc = -1: Select Case TypeName(vIn)
        Case "String"
            With ColName
                If Not .IsIn(vIn) Then Exit Function
                Dim Ind As Long: For Ind = .Bottom To .Top
                    If ColName(Ind) = vIn Then Enc = Ind: Exit For
                Next Ind
            End With
        Case Else: Enc = CLng(vIn) - FirstCol
    End Select
End Function

Public Function WriteToCSV(vSession As wConn, vTableName As String, Optional LeadRows As String = "") As Boolean: WriteToCSV = False
    Const BatchSize = 10000: Const rDelim = vbCrLf: Const fDelim = "|#|"
    Dim DataBlock As w2DArr, CSV As cTxt: Set CSV = New cTxt: CSV.Init
    If LastRow - FirstRow + 1 > BatchSize Then GoTo Chunked
    Set DataBlock = Block
    If Not BulkWriteBlockToCSV(CSV, DataBlock, LeadRows, rDelim, fDelim, False) Then
        SpoolBulkWriteErrorsToSheet CSV, DataBlock, vTableName, LeadRows
        Exit Function
    Else
        WriteToCSV = True
    End If
    Exit Function
Chunked:
    Dim qInd As Long, CheckLoop As Boolean: CheckLoop = True: For qInd = 0 To Int((LastRow - FirstRow) / BatchSize)
        Set DataBlock = New w2DArr: DataBlock.Init Sh.Range(Sh.Cells(FirstRow + qInd * BatchSize + IIf(qInd = 0, 0, 1), FirstCol), Sh.Cells(IIf(LastRow < FirstRow + (qInd + 1) * BatchSize, LastRow, FirstRow + (qInd + 1) * BatchSize), LastCol)).Value
        If Not BulkWriteBlockToCSV(CSV, DataBlock, LeadRows, rDelim, fDelim, True) Then
            SpoolBulkWriteErrorsToSheet CSV, DataBlock, vTableName & ": Batch " & qInd * BatchSize + 1 & " - " & (qInd + 1) * BatchSize
            CheckLoop = False
        End If
    Next qInd
    WriteToCSV = CheckLoop
End Function

Public Function BulkWriteBlockToCSV(vTxt As cTxt, vDataBlock As w2DArr, LeadRows As String, rDelim As String, fDelim As String, Optional Appending As Boolean = False) As Boolean: BulkWriteBlockToCSV = False
    On Error GoTo ExitFunc
    vTxt.WriteTo IIf(vTxt.Contents = "", rDelim, "") & vDataBlock.Implode(LeadRows, rDelim, CStr(fDelim)) & rDelim, False, Appending
    If Not vTxt.WriteIssue Then BulkWriteBlockToCSV = True
ExitFunc:
End Function

Public Sub SpoolBulkWriteErrorsToSheet(vTxt As cTxt, vDataBlock As w2DArr, vTableName As String, Optional LeadRows As String = "")
    On Error Resume Next
    Dim ErrSh As Worksheet, Ind As Long: Set ErrSh = ThisWorkbook.Sheets("Import Errors")
    If ErrSh Is Nothing Then
        Set ErrSh = ThisWorkbook.Sheets.Add
        ErrSh.Name = "Import Errors"
    End If
    Dim FirstRow As Long
    If ErrSh.Range("A1").End(xlDown).Row = ErrSh.Rows.Count Then
        ErrSh.Range("A1").Value = "Import ID " & LeadRows & ": " & vTableName
        For Ind = 0 To ColName.Top - ColName.Bottom
            ErrSh.Range("A1").Offset(0, Ind + 1) = ColName(Ind + ColName.Bottom)
        Next Ind
        xlStyle.RngFormat ErrSh.Range(ErrSh.Range("A1"), ErrSh.Range("A1").End(xlToRight)), "SubHead"
        FirstRow = 1
    Else
        ErrSh.Range("A1").End(xlDown).Offset(1, 0).Value = "Import ID " & LeadRows & ": " & vTableName
        For Ind = 0 To ColName.Top - ColName.Bottom
            ErrSh.Range("A1").End(xlDown).Offset(0, Ind + 1) = ColName(Ind + ColName.Bottom)
        Next Ind
        xlStyle.RngFormat ErrSh.Range(ErrSh.Range("A1").End(xlDown), ErrSh.Range("A1").End(xlDown).End(xlToRight)), "SubHead"
        FirstRow = ErrSh.Range("A1").End(xlDown).Row
    End If
    Dim Coin As Long: Coin = 0: For Ind = vDataBlock.Bottom(xAxis) To vDataBlock.Top(xAxis)
        Debug.Print Coin, vDataBlock.Row(Ind).Implode("|")
        vTxt.WriteIssue = False
        vTxt.WriteTo vDataBlock.Row(Ind).Implode("|"), , True
        If vTxt.WriteIssue Then GoTo LogRow
NextRow:
    Next Ind
    xlStyle.RngFormat ErrSh.Range(ErrSh.Range("A1").End(xlDown).Offset(-Coin + 1, 0), ErrSh.Range("A1").End(xlDown).Offset(-Coin, 0).End(xlToRight).Offset(Coin, 0)), "ROCell"
    ErrSh.Columns.EntireColumn.ColumnWidth = 5000
    ErrSh.Rows.EntireRow.AutoFit
    ErrSh.Columns.EntireColumn.AutoFit
    Exit Sub
LogRow:
    Coin = Coin + 1
    ErrSh.Range("A" & FirstRow).Offset(Coin, 0).Value = Coin
    Dim cInd As Long: For cInd = 0 To vDataBlock.Row(Ind).Top - vDataBlock.Row(Ind).Bottom
        ErrSh.Range("A" & FirstRow).Offset(Coin, cInd + 1).Value = vDataBlock.Row(Ind)(cInd + vDataBlock.Row(Ind).Bottom)
    Next cInd
    GoTo NextRow
End Sub

Private Sub PrepareSheet()
    With Sh
        On Error Resume Next
        If .ProtectContents Then .Unprotect IIf(Tbls.eFlow, "", ShPwd)
        If .AutoFilterMode Then If .FilterMode Then .ShowAllData
        ' .Rows.Hidden = False: .Columns.Hidden = False
        .Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
'        .Cells.Replace ChrW(8540), ",3/8", xlPart
'        .Cells.Replace Chr(150), Chr(45), xlPart
'        .Cells.Replace Chr(151), Chr(45), xlPart
'        .Cells.Replace Chr(188), ",1/4", xlPart
'        .Cells.Replace Chr(147), """", xlPart
'        .Cells.Replace Chr(10), "/n", xlPart
'        .Cells.Replace Chr(13), "", xlPart
'        .Cells.Replace "|", "\", xlPart
'        .Cells.Replace ChrW(9651), "Delta", xlPart
    End With
    On Error GoTo 0
End Sub

Public Function DataType(vCol) As String
    Const DTCheck = 10000
    Dim CheckCol As Variant: CheckCol = Sh.Range(Sh.Cells(FirstRow, Enc(vCol) + FirstCol), Sh.Cells(IIf(LastRow - FirstRow + 1 > DTCheck, DTCheck, LastRow), Enc(vCol) + FirstCol)).Value2
    Dim CellVal As String, EmptyCol As Boolean, Num As Boolean, IntOnly As Boolean: Num = True: IntOnly = True: EmptyCol = True
    If Not IsArray(CheckCol) Or IsEmpty(CheckCol) Then GoTo SingleRow
    Dim Ind As Long: For Ind = LBound(CheckCol, 1) To UBound(CheckCol, 1)
        Num = True: IntOnly = True: EmptyCol = True
        CellVal = IIf(IsError(CheckCol(Ind, LBound(CheckCol, 2))), "", CheckCol(Ind, LBound(CheckCol, 2)))
        If Not IsEmpty(CellVal) Then EmptyCol = False
        If Not IsNumeric(CellVal) Then
            Num = False
        Else
            If Not Int(CellVal) = CellVal Then IntOnly = False
        End If
    Next Ind
Finish:
    DataType = IIf(EmptyCol Or Not Num, "char", IIf(IntOnly, "numeric", "monetary"))
    Exit Function
SingleRow:
    CellVal = IIf(IsError(CheckCol), "", CheckCol)
    If Not IsEmpty(CellVal) Then EmptyCol = False
    If Not IsNumeric(CellVal) Then
        Num = False
    Else
        If Not Int(CellVal) = CellVal Then IntOnly = False
    End If
    GoTo Finish
End Function
