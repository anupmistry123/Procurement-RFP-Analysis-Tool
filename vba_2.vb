' this class module is for a 1 dimensional arrary. It is called a w1DArr with "w" representing "wrapped".
' the "wrapped" indicates that this class module has more properties and functions than the standard 1DArr functionality within VBA.
Option Explicit
Private uArr() As Variant
Public Enum ArrEdge: TopEdge = 1: BottomEdge = 2: End Enum

Public Property Get Top() As Long: Top = Bound(TopEdge): End Property

Public Property Get Bottom() As Long: Bottom = Bound(BottomEdge): End Property

Public Property Get Val(vInd As Long) As Variant
    If Not (vInd >= Bottom And vInd <= Top) _
    Or IsEmpty Then Exit Property
    If IsObject(uArr(vInd)) Then Set Val = uArr(vInd) Else Val = uArr(vInd)
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = True
    On Error GoTo Out
    Dim Ind As Long: Ind = UBound(uArr, 1)
    IsEmpty = False
Out:
End Property

Public Property Get Count() As Long
    Count = 0: If Not IsEmpty Then Count = Top - Bottom + 1
End Property

Private Property Get Bound(Optional vArrEdge As Long = TopEdge) As Long
    Bound = -1
    If Not IsEmpty Then
        Select Case vArrEdge
            Case TopEdge:     Bound = UBound(uArr, 1)
            Case BottomEdge:  Bound = LBound(uArr, 1)
        End Select
    End If
End Property

Public Property Get Arr() As Variant()
    If IsEmpty Then Exit Property Else Arr = uArr
End Property

Public Property Get Implode(Optional fDelim As String = "|") As String
    Implode = Join(Arr, fDelim)
End Property

Public Property Get Slice(vInd As Long, Optional vEdge = ArrEdge.TopEdge) As w1DArr
    If IsEmpty Then Exit Property
    Dim vIndAdj As Long: vIndAdj = IIf(vInd >= Top, Top, IIf(vInd <= Bottom, Bottom, vInd))
    Dim Max As Long, Min As Long: Select Case vEdge
        Case TopEdge:    Max = Top:     Min = vIndAdj
        Case BottomEdge: Max = vIndAdj: Min = Bottom
    End Select
    Dim Temp() As Variant: ReDim Temp(0 To Max - Min)
    Dim Coin As Long: For Coin = 0 To Max - Min
        Temp(Coin) = uArr(Coin + Min)
    Next Coin
    Dim tempArr As w1DArr: Set tempArr = New w1DArr: tempArr.Init Temp
    Set Slice = tempArr
End Property

Public Property Get Segment(vInd1 As Long, vInd2 As Long) As w1DArr
    If IsEmpty Then Exit Property
    Dim vIndAdj1 As Long, vIndAdj2 As Long
    vIndAdj1 = IIf(vInd1 <= vInd2, vInd1, vInd2): vIndAdj2 = IIf(vInd1 <= vInd2, vInd2, vInd1)
    Dim Min As Long, Max As Long
    Min = IIf(vIndAdj1 <= Bottom, Bottom, vIndAdj1)
    Max = IIf(vIndAdj2 >= Top, Top, vIndAdj2)
    Dim Temp() As Variant: ReDim Temp(Min To Max)
    Dim Coin As Long: For Coin = Min To Max
        Temp(Coin) = uArr(Coin)
    Next Coin
    Dim tempArr As w1DArr: Set tempArr = New w1DArr: tempArr.Init Temp
    Set Segment = tempArr
End Property

Public Sub Init(vArr)
    If Not IsArray(vArr) Then Exit Sub
    Erase uArr
    If TypeName(vArr) = "Variant()" Then
        uArr = vArr
    Else
        Dim Val: For Each Val In vArr
            Add Val, LBound(vArr, 1)
        Next Val
    End If
End Sub

Public Function IsIn(vElement) As Boolean
    IsIn = False
    If IsEmpty Then Exit Function
    Dim Element As Variant: For Each Element In Arr
        If IsObject(Element) And IsObject(vElement) Then
            If Element Is vElement Then IsIn = True
        ElseIf Not IsObject(Element) And Not IsObject(vElement) Then
            If Element = vElement Then IsIn = True
        End If
    Next Element
End Function

Public Sub Add(vElement, Optional vStart As Long = 0)
    Dim vBottom As Long: vBottom = IIf(IsEmpty, vStart, Bottom)
    Dim vTop As Long: vTop = IIf(IsEmpty, vStart - 1, Top)
    ReDim Preserve uArr(vBottom To vTop + 1) As Variant
    If IsObject(vElement) Then Set uArr(vTop + 1) = vElement Else uArr(vTop + 1) = vElement
End Sub

Public Sub SetVal(vInd As Long, vSetVal)
    If IsEmpty Or Not (vInd >= Bottom And vInd <= Top) Then Exit Sub
    If IsObject(vSetVal) Then Set uArr(vInd) = vSetVal Else uArr(vInd) = vSetVal
End Sub

Public Sub ReDimArr(vLBound As Long, vUBound As Long, Optional PreserveInd = False)
    If PreserveInd Then
        If vLBound = Bottom Then
            ReDim Preserve uArr(Bottom To vUBound)
        Else
            Dim Temp() As Variant: ReDim Temp(vLBound To vUBound) As Variant
            Dim Ind As Long: For Ind = vLBound To vUBound
                If IsObject(uArr(Ind)) Then Set Temp(Ind) = Val(Ind) Else Temp(Ind) = Val(Ind)
            Next Ind
            Erase uArr: uArr = Temp
        End If
    Else
        ReDim uArr(vLBound To vUBound)
    End If
End Sub

Public Sub Remove(vInd As Long)
    If IsEmpty Or Not (vInd >= Bottom And vInd <= Top) Then Exit Sub
    Dim Ind As Long: For Ind = vInd To Top - 1
        SetVal vInd, Val(vInd + 1)
    Next Ind
    ReDimArr Bottom, Top - 1, True
End Sub

Public Sub Clear(): Erase uArr: End Sub

Public Sub Spool()
    Debug.Print Implode(", ")
End Sub

