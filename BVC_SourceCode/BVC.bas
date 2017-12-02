Attribute VB_Name = "Module2"
Option Explicit


'''FOR PASSWORD PROTECTION'''

Sub Pswd(control As IRibbonControl)
    On Error GoTo Oboy
    ActiveSheet.Activate
    MsgBox "WARNING!!!" & Chr(10) & _
    "You are about to set Password" _
    & " for your Active Workbook" _
    & ". Please do not forget it." _
    , vbInformation, "Password Setup"
    
    CPass.Show
Oboy:
End Sub

'''SET PASSWORD USERFORM'''

Private Sub SetP_Click()
    If Pssd = "" Then
        MsgBox "Please submit your password", , _
        "Password Setup"
        Exit Sub
    End If
    With ActiveWorkbook
        If .Path = "" Then
            MsgBox "Please save the workbook first," & _
            " in order to setup a password.", , _
            "Password Setup"
            Unload Me
        Else
            Application.DisplayAlerts = False
            .SaveAs .Path & "\" & .Name, , Pssd
            MsgBox "Please do not forget your password." _
            & " You have just secured your workbook" _
            & " viewing.", vbInformation, _
            "Password Setup"
            Application.DisplayAlerts = True
            Unload Me
        End If
    End With
End Sub


'''USER DEFINE FUNCTIONS COLLECTIONS'''

Function RaCo(Optional RaColumn As String) As Long
    If RaColumn = "" Then
        RaColumn = Replace(ActiveCell.Address, "$", "")
    End If
    On Error Resume Next
    RaCo = Range(RaColumn).Column
    If Err.Number <> 0 Then
        On Error GoTo 0
    End If
    RaColumn = vbNullString
End Function

Function RaRo(Optional RaRow As String) As Long
    If RaRow = "" Then
        RaRow = Replace(ActiveCell.Address, "$", "")
    End If
    On Error Resume Next
    RaRo = Range(RaRow).Row
    If Err.Number <> 0 Then
        On Error GoTo 0
    End If
    RaRow = vbNullString
End Function

Function CelNam(Optional Crow As Long, Optional Ccol As Long) As String
    On Error Resume Next
    If Crow = 0 And Ccol = 0 Then
        CelNam = Replace(ActiveCell.Address, "$", "")
    Else
        CelNam = Replace(Cells(Crow, Ccol).Address, "$", "")
    End If
Crow = 0
Ccol = 0
End Function

Function ActNam(Optional ActSheet As Worksheet) As String
    If ActSheet Is Nothing Then
        Set ActSheet = ActiveSheet
        ActNam = ActSheet.Name
    Else
        ActNam = ActSheet.Name
    End If
Set ActSheet = Nothing
End Function

Function ToRo(Optional tR As Range) As LongPtr
    If tR Is Nothing Then
        Set tR = ActiveCell
        With tR
            ToRo = Cells(Rows.Count, .Column).End(xlUp).Row
        End With
    Else
        With tR
            ToRo = Cells(Rows.Count, .Column).End(xlUp).Row
        End With
    End If
Set tR = Nothing
End Function

Function ToCo(Optional TC As Range) As LongPtr
    If TC Is Nothing Then
        Set TC = ActiveCell
        With TC
            ToCo = Cells(.Row, Columns.Count).End(xlToLeft).Column
        End With
    Else
        With TC
            ToCo = Cells(.Row, Columns.Count).End(xlToLeft).Column
        End With
    End If
Set TC = Nothing
End Function

Function CACel() As Range
Dim Ca As LongPtr
Dim Cbl As LongPtr
Ca = WorksheetFunction.CountA(Cells(1, 1).Resize(ToRo))
Cbl = WorksheetFunction.CountA(Columns(1).Rows)
With ActiveCell
    Set CACel = Cells(ToRo + ((Cbl - Ca) + 1), .Column)
End With
Ca = 0
Cbl = 0
End Function

Function CoS() As Long
Dim CS As Shape
Dim NSh As Long
Dim ShS As Long
Dim ISh As Long
Dim St As Long
Dim k As Long
    St = ActiveCell.Column
    NSh = ActiveSheet.Shapes.Count
    ShS = ActiveCell.Row
    If ActiveCell.Column = 1 Then
    For ISh = ShS To NSh
        Set CS = ActiveSheet.Shapes(ISh)
        With CS
            If Range(.Name).Column > St Then
                k = k + 1
            Else
                Exit For
            End If
        End With
    Next ISh
    Else
    CoS = 0
    End If
    CoS = k
Set CS = Nothing
NSh = 0
ShS = 0
ISh = 0
St = 0
k = 0
End Function

Function ColS() As Integer
Dim CoSt As Integer
Dim VC As Integer
Dim k As Integer
CoSt = ActiveCell.SpecialCells(xlCellTypeLastCell).Column
    With ActiveCell
        If .Column = 1 Then
            For k = .Column To CoSt
                On Error Resume Next
                VC = WorksheetFunction. _
                CountA(.Offset(, k) _
                .Resize(CoS))
                If VC = 0 Then
                    Exit For
                End If
            Next k
        End If
    End With
    ColS = k
CoSt = 0
VC = 0
k = 0
End Function

Function fLR() As LongPtr
Dim tR As LongPtr
Dim j As LongPtr
Dim k As LongPtr
    If Right(ActiveSheet.Name, 5) = "Check" Then
        tR = Int(Cells.SpecialCells(xlCellTypeConstants) _
        .Count - 1)
    Else
        tR = Int(Cells.SpecialCells(xlCellTypeConstants) _
        .Count)
    End If
    For j = 1 To tR
        If Range("A" & j).EntireRow.Hidden = False Then
            k = 0
            k = Range("A" & j).Row
        End If
    Next j
    fLR = k
tR = 0
j = 0
k = 0
End Function



'''INITIAL START SETUP FOR BVC'''

Dim Header As Variant
Dim i As Integer
Dim Tbl As ListObjects
Dim ScrR As Range
Dim Tbl1 As ListObject
Dim Psw As String
Dim PswA As String
Dim Tx As Workbook, TXA As Workbook
Dim TXW As Worksheet, TXWA As Worksheet
Dim R1 As Range, R2 As Range, R3 As Range
Dim MoA As String
Dim Pswd As String
Dim WarnD As Variant

Sub CekBVC(control As IRibbonControl)
    'ActLib
Dim M As Variant
    On Error GoTo Oboy
    ActiveSheet.Activate

    Set TXA = ActiveWorkbook
    Set TXWA = TXA.ActiveSheet
    With TXWA
        Select Case .Cells(1, 1)
            Case Is = ""
                If .ProtectContents = False Then
                    M = MsgBox( _
                    "Do you want to clear the whole cells?" _
                    , vbOKCancel, "Bible Verses Collections")
                    Select Case M
                        Case Is = vbOK
                            .Cells.Clear
                            .Tab.ColorIndex = 5
                            Setupv
                        Case Else
                            MsgBox "This Add-In is functioning." _
                            & " Please try again in new" _
                            & " Worksheet.", , _
                            "Bible Verses Collections"
                    End Select
                Else
                    MsgBox _
                    "This Add-In is functioning." _
                    & " Please use it in new worksheet." _
                    , , "Bible Verses Collections"
                End If
            Case Is = "No."
                With TXWA
                If Left(.Name, 24) = _
                "Bible Verses Collections" Then
                    VersesColl.Show
                Else
                    MsgBox _
                    "Your Add-In is functioning." _
                    & " Please use it in new worksheet." _
                    , , "Bible Verses Collections"
                End If
                End With
            Case Else
                MsgBox "This Add-In is functioning." _
                & " Please try again in new" _
                & " Worksheet.", , "Bible Verses Collections"
        End Select
    End With
Oboy:
Set TXA = Nothing
Set TXWA = Nothing
M = 0
End Sub

Sub Setupv()
On Error GoTo Good
Set Tbl = ActiveSheet.ListObjects
    'Start in range A1
    If ActiveCell.Address <> "$A$1" Then
        Cells(1, 1).Select
    End If
    Header = Array("No.", "Verses", "Ponder")
    For i = 0 To UBound(Header)
        With Cells(1, i + 1)
            .Value = Header(i)
            .HorizontalAlignment = xlLeft
        End With
    Next i
    Columns("A:B").AutoFit
    Columns(3).ColumnWidth = 60
Set ScrR = Range("A1:C2")
On Error GoTo Good
    Tbl.Add xlSrcRange, ScrR, , xlYes
Set Tbl1 = ActiveCell.ListObject
    With Tbl1
        .Name = "BibleVerses"
        .ShowAutoFilter = False
        .TableStyle = "TableStyleMedium13"
    End With
    ShowForm
Good:
Set Tbl = Nothing
Set ScrR = Nothing
Set Tbl1 = Nothing
i = 0
End Sub


'''USERFORMS BVC'''

Dim Tbl As ListObject
Dim Fr As Variant
Dim Adr As String
Dim F As Integer
Dim i As Integer
Dim Area As Long
Dim Psw As String
Dim PswA As String
Dim Tx As Workbook, TXA As Workbook
Dim TXW As Worksheet, TXWA As Worksheet
Dim R1 As Range, R2 As Range, R3 As Range

Private Sub AddV_Click()
    If Err.Number <> 0 Then GoTo Oboy
    Call CallV_Click
Oboy:
Set Tbl = ActiveSheet.ListObjects(ActiveSheet.ListObjects.Count)
    If ActiveCell = "" Then
Set Tbl = Nothing
        Exit Sub
    Else
        Tbl.ListRows.Add
        CallV_Click
    End If
 Set Tbl = Nothing
End Sub

Private Sub Ayat_Change()
Fx1
    If ActiveCell.Column = 2 Then
        'If Not ActiveCell.Offset(, 1) = "" Then
            With ActiveCell
                .Value = Ayat.Text
                With .Offset(, -1)
                    If .Value = "" Then
                        .FormulaR1C1 = _
        "=IF(RC2=0,"""",COUNTA(R2C2:[@Verses]))"
                        .Value = .Value
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlCenter
                    End If
                End With
                .VerticalAlignment = xlCenter
            End With
        'End If
    End If
    Cells.Columns(2).AutoFit
Fx2
End Sub

Private Sub Ayat_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Ayat.Locked = False
End Sub

Private Sub CallV_Click()

Set Tbl = ActiveSheet.ListObjects(ActiveSheet.ListObjects.Count)
    Fr = Cells(Rows.Count, 1).End(xlUp).Row
    If ActiveCell.Address = "$B$2" And ActiveCell = "" Then
        Exit Sub
    Else
        On Error Resume Next
        Tbl.DataBodyRange(Fr, 2).End(xlUp).Select
        If Err.Number <> 0 Then AddV_Click
    End If
    App
Set Tbl = Nothing
Fr = 0
End Sub


Private Sub DelV_Click()
Dim M As Variant
    'Call CallV_Click
'Fx1
    With ActiveCell
        If .Column = 2 Then
            If Not .Address = "$B$2" Then
                M = MsgBox("Do you want to delete this verse" _
                & " and it's comment?", vbYesNo, "Bible Verses Collections")
                Select Case M
                    Case Is = vbYes
                        .Offset(, -1).Resize(, 3).Delete
                        ReNum
                End Select
            Else
                If .Offset(1).Value = "" And _
                Intersect(ActiveCell.Offset(1), _
                ActiveSheet.ListObjects("BibleVerses"). _
                DataBodyRange) Is Nothing Then
                    .Offset(, -1).Resize(, 3).ClearContents
                    DelCom_Click
                Else
                    MsgBox "You cannot clear the first row" _
                    & " when the records are more then 1" _
                    & ". Edit through the textboxes for changes.", _
                    , "Bible Verses Collections"
                End If
            End If
        End If
    End With
        
'Fx2
M = 0
    Call CallV_Click
End Sub

Private Sub ReNum()
Fx1
Dim M As Variant
    M = Cells(Rows.Count, 2).End(xlUp).Row
    If ActiveCell.Row = M + 1 Then GoTo bye
    For i = 2 To M
        With Cells(i, 2)
            With .Offset(, -1)
                .Clear
                .FormulaR1C1 = _
            "=IF(RC2=0,"""",COUNTA(R2C2:[@Verses]))"
                .Value = .Value
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
            If .VerticalAlignment <> xlCenter Then
                .VerticalAlignment = xlCenter
            End If
            With .Offset(, 1)
                If .WrapText = False Then
                    .WrapText = True
                End If
                If .VerticalAlignment <> xlCenter Then
                    .VerticalAlignment = xlCenter
                End If
            End With
        End With
    Next i
bye:
M = 0
i = 0
Fx2
End Sub

Private Sub MultiPage1_Change()
    If MultiPage1.SelectedItem.Name = "page1" Then
        If Not ActiveCell = "" Then
            Ayat.Locked = True
            Ayat = ActiveCell.Value
            Renungan.Locked = True
            Renungan = ActiveCell.Offset(, 1).Value
        End If
        Ayat.SetFocus
    Else
        If Not ActiveCell.Offset(, 1).Comment Is Nothing Then
            Catatan.Locked = True
            Catatan = ActiveCell.Offset(, 1).Comment.Text
        Else
            If Catatan.Locked = True Then
                Catatan.Locked = False
                Catatan = ""
            Else
                Catatan = ""
            End If
        End If
        On Error Resume Next
        Catatan.SetFocus
        If Err.Number <> 0 Then
            Ayat.SetFocus
            On Error GoTo 0
        End If
    End If
End Sub
Private Sub NumberR_Change()
    NumberR.Locked = True
    SpinButton2.Value = NumberR.Text
End Sub

Private Sub NumberR_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    NumberR = 1
End Sub

Private Sub PasteV_Click()
Dim k As Long
Dim j As Long
Dim TxR As String
Dim Ch As Characters
Fx1

    If ActiveCell <> "" And ActiveCell.Offset(, 1) <> "" Then _
    Exit Sub
On Error Resume Next
    With ActiveCell
        .Offset(, 1).Select
        ActiveSheet.Paste
        .Select
    End With
Dim Ro As Long
Dim Cn1 As String, Cn2 As String, Raj As String
Dim Nu As Long, Ln As Long, Rn As Long
Dim Ls As Long, Rs As Long
Dim Lss As Long, Rss As Long
Dim Lnn As Long, Rnn As Long
Ro = Cells(Rows.Count, RaCo(ActiveCell.Offset(, 1).Address)) _
.End(xlUp).Row - RaRo()
    For Nu = 1 To Ro
        With Cells(RaRo() + (Nu - 1), _
        RaCo(ActiveCell.Offset(, 1).Address))
            If InStr(.Value, ":") <> 0 Then
                If IsNumeric(.Characters(InStr _
                (.Value, ":") - 1, 1).Text) = False Then GoTo nex
                Ls = Len(Left(.Value, InStr(.Value, ":")))
                Ln = 1
                Do Until Ln = Ls
                    'MsgBox .Characters(Ls - Ln, 1).Text
                    If .Characters(Ls - Ln, 1).Text _
                    = " " Then
                        Lss = (Ls - 1) - Ln
                        Do Until Lnn = Lss
                            'MsgBox .Characters(Lss - Lnn, 1).Text
                            If .Characters(Lss - Lnn, 1).Text _
                            = " " Then
                                If IsNumeric(.Characters _
                                (Lss - (Lnn + 1), 1).Text) _
                                Then
                                    Cn1 = Right(Left(.Value, Ls - 1), (Lnn + 2) + Ln)
                                    'MsgBox Cn1
                                    GoTo Adioso
                                Else
                                    Cn1 = Right(Left(.Value, Ls - 1), Lnn + Ln)
                                    GoTo Adioso
                                End If
                            End If
                            Lnn = Lnn + 1
                        Loop
                        Cn1 = Right(Left(.Value, Ls - 1), Lnn + Ln)
Adioso:
                        Exit Do
                    End If
                    Ln = Ln + 1
                Loop
                
                
                Rs = Len(Mid(.Value, InStr(.Value, ":")))
                Rn = 1
                Do Until Rn = Rs
                    'MsgBox .Characters(Ls + Rn, 1).Text
                    If .Characters(Ls + Rn, 1).Text _
                    = " " Then
                        Rss = (Ls + 1) + Rn
                        Do Until Rnn = Rs
                            'MsgBox .Characters(Rss + Rnn, 1).Text
                            If .Characters(Rss + Rnn, 1).Text _
                            = " " Or _
                            .Characters(Rss + Rnn, 1).Text = "" Then
                                
                                Cn2 = Mid(.Value, Ls + 1, Rnn + Rn)
                                'MsgBox Cn2
                                GoTo Amios
                            End If
                            Rnn = Rnn + 1
                        Loop
Amios:
                        Exit Do
                    End If
                    Rn = Rn + 1
                Loop
            End If
            Raj = Cn1 & ":" & Cn2
            'MsgBox Raj
            If Len(Raj) > 1 Then
                .Offset(, -1).Value = Raj
            End If
            If Right(Raj, 10) <> "Indonesian" Then
                If Len(Raj) = Len(.Value) Then
                    .Value = ""
                ElseIf Left(.Value, Len(Raj)) = Raj Then
                    .Value = Mid(.Value, Len(Raj) + 3)
                ElseIf Right(.Value, Len(Raj)) = Raj Then
                    .Value = Left(.Value, Len(.Value) - (Len(Raj) + 2))
                End If
            Else
                If Len(Raj) = Len(.Value) - 3 Then
                    .Offset(, -1).Value = .Value
                    .Value = ""
                ElseIf Left(.Value, Len(Raj)) = Raj Then
                    .Offset(, -1).Value = Left(.Value, Len(Raj) + 3)
                    .Value = Mid(.Value, Len(Raj) + 6)
                ElseIf Right(.Value, Len(Raj)) <> Raj Then
                    .Offset(, -1).Value = Right(.Value, Len(Raj) + 3)
                    .Value = Left(.Value, Len(.Value) - (Len(Raj) + 5))
                End If
            End If

        End With
nex:
    Ls = 0
    Ln = 0
    Rs = 0
    Rn = 0
    Lnn = 0
    Rnn = 0
    Raj = vbNullString
    Cn1 = vbNullString
    Cn2 = vbNullString
Next Nu
    
    Dete
    DelC
    CallV_Click
    ReNum
Fx2
End Sub
Private Sub Dete()
Dim RoC As Long
Dim NCe As Long
Dim Lo As Long
    RoC = Cells(Rows.Count, 2).End(xlUp).Row
    For NCe = 0 To RoC - RaRo()
        With Cells(RaRo() + NCe, RaCo())
            If .Value = "" And .Offset(, 1) = "" Then
                GoTo Wow
            ElseIf .Value <> "" And .Offset(, 1) = "" Then
                If .Offset(1) = "" And .Offset(1, 1) <> "" Then
                    .Offset(, 1).Value = .Offset(1, 1).Value
                    .Offset(1, 1).Value = ""
                ElseIf .Offset(-1) = "" And .Offset(-1, 1) <> "" Then
                    .Offset(, 1).Value = .Offset(-1, 1).Value
                    .Offset(-1, 1).Value = ""
                End If
            ElseIf .Value = "" Or .Value = ":" Then
                If .Offset(1) <> "" And .Offset(1, 1) = "" Then
                    .Value = .Offset(1).Value
                    .Offset(1).Value = ""
                ElseIf .Offset(-1) <> "" And .Offset(-1, 1) = "" Then
                    .Value = .Offset(-1).Value
                    .Offset(-1).Value = ""
                End If
            End If
            'MsgBox .Value & " " & .Offset(, 1)
        End With
Wow:
    Next NCe
    DeDl
bye:
RoC = 0
NCe = 0
End Sub

Private Sub DeDl()
Dim RoC As Long
Dim NCe As Long
Dim Lo As Long
    RoC = Cells(Rows.Count, 2).End(xlUp).Row

    Do Until Lo = RoC - RaRo()
        With ActiveCell
            If .Offset(Lo) = "" Then
                On Error GoTo bye
                If .Offset(Lo).ListObject.Active = True Then
                    .Offset(Lo).EntireRow.Delete
                End If
            End If
        End With
        Lo = Lo + 1
    Loop
bye:
RoC = 0
Lo = 0
End Sub
Private Sub DelC()
Fx1
Set Tbl = ActiveSheet.ListObjects(ActiveSheet.ListObjects.Count)
    Fr = Cells(Rows.Count, 1).End(xlUp).Row
    CallV_Click
    Do Until Cells(ActiveCell.Row - i, 2) <> ""
        i = i + 1
    Loop
    'IF cells(activecell.Row-i
    Cells(ActiveCell.Row - i, 2).Select
    On Error GoTo bye
    Cells(ActiveCell.Row, 2).Offset(1, -1).Resize(i, 3).Delete
bye:
Set Tbl = Nothing
Fr = 0
i = 0
Fx2
End Sub
Private Sub Renungan_Change()
    If ActiveCell.Column = 2 Then
        'If ActiveCell.Offset(, -1) = "" Then
            With ActiveCell.Offset(, 1)
                If Renungan <> "Ponder" Then
                    .Value = Renungan.Text
                    .WrapText = True
                    .VerticalAlignment = xlCenter
                End If
            End With
        'End If
    End If
    With Columns("C")
        If .ColumnWidth <> 60 Then
            .ColumnWidth = 60
        End If
    End With
End Sub

Private Sub Renungan_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Renungan.Locked = False
End Sub

Private Sub SpeechV_Click()
Fx1
    For i = 0 To NumberR - 1
        If ActiveCell.Column = 2 Then
            Application.Speech.Speak _
            ActiveCell.Offset(i) & _
            ", " & ActiveCell.Offset(i, 1)
        End If
    Next i
Fx2
End Sub

Private Sub SpinButton1_SpinDown()
    If ActiveCell.Column = 2 Then
        If ActiveCell.Offset(1) = "" Then
            Exit Sub
        Else
            ActiveCell.Offset(1).Select
            App
        End If
    End If
End Sub

Private Sub SpinButton1_SpinUp()
    If ActiveCell.Column = 2 Then
        If ActiveCell.Offset(-1).Address = "$B$1" Then
            Exit Sub
        Else
            ActiveCell.Offset(-1).Select
            App
        End If
    End If
End Sub

Private Sub SpinButton2_Change()
    NumberR.Text = SpinButton2.Value
End Sub


Private Sub SViewer_Click()
    BVViewer.Show
End Sub

Private Sub UserForm_Initialize()
Dim Ct As LongPtr, k As LongPtr
Fx1
    ActiveSheet.Unprotect
    With ActiveWindow
        If .FreezePanes = False Then
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End If
    End With
    For Ct = 1 To Worksheets.Count
        If Left(Worksheets(Ct).Name, 24) _
        = "Bible Verses Collections" Then
            k = k + 1
        End If
    Next Ct
    If Left(ActiveSheet.Name, 24) <> "Bible Verses Collections" Then
        If k > 0 Then
            With ActiveWorkbook
                If .ProtectStructure = False Then
                    ActiveSheet.Name = "Bible Verses Collections " & k + 1
                Else
                    MsgBox "Please unlock the workbook first", vbInformation, _
                    "Bible Verses Collections"
                    End
                End If
            End With
        Else
            With ActiveWorkbook
                If .ProtectStructure = False Then
                    ActiveSheet.Name = "Bible Verses Collections"
                Else
                    MsgBox "Please unlock the workbook first", vbInformation, _
                    "Bible Verses Collections"
                    End
                End If
            End With
        End If
    End If
Ct = 0
k = 0
    CallV_Click
Fx2
End Sub

Private Sub UserForm_Terminate()
Fx1
    With ActiveSheet
        .Protect
        .EnableSelection = xlNoSelection
    End With
Fx2
End Sub
Private Sub AddCom_Click()
    If ActiveCell.Column = 2 Then
        If Not ActiveCell = "" _
        And Not ActiveCell.Offset(, 1) _
        = "" Then
            DelCom_Click
            With ActiveCell.Offset(, 1)
            On Error GoTo Good
                .AddComment
                With .Comment
                    .Text Catatan.Text
                    'found this in Contextures.com
                    'by Dana DeLouis 2000/09/16 & Tom Urtis
                    .Shape.TextFrame.AutoSize = True
                    If .Shape.Width > 200 Then
                        Area = .Shape.Width * _
                        .Shape.Height
                        .Shape.Width = 170
                        .Shape.Height = (Area / 170)
                    End If
                End With
            End With
        End If
    End If
Good:
End Sub

Private Sub Catatan_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Catatan.Locked = False
End Sub

Private Sub DelCom_Click()
    If ActiveCell.Column = 2 Then
        If ActiveCell.Offset(, 1).Comment _
        Is Nothing Then
            Exit Sub
        Else
            ActiveCell.Offset(, 1). _
            Comment.Delete
        End If
    End If
End Sub

Private Sub SpinButton3_SpinDown()
    If ActiveCell.Column = 2 Then
        If ActiveCell.Offset(1) = "" Then
            Exit Sub
        Else
            ActiveCell.Offset(1).Select
            CoAp
        End If
    End If
End Sub

Private Sub CoAp()
    If Not ActiveCell.Offset(, 1) _
    .Comment Is Nothing Then
        Catatan.Locked = True
        Catatan = ActiveCell.Offset(, 1) _
        .Comment.Text
    Else
        With Catatan
            .Locked = False
            .Text = ""
        End With
    End If
End Sub

Private Sub SpinButton3_SpinUp()
    If ActiveCell.Column = 2 Then
        If ActiveCell.Offset(-1). _
        Address = "$B$1" Then
            Exit Sub
        Else
            ActiveCell.Offset(-1).Select
            CoAp
        End If
    End If
End Sub

Private Sub Fast(SU As Boolean, DS As Boolean, C As String, EE As Boolean)
    Application.ScreenUpdating = SU
    Application.DisplayStatusBar = DS
    Application.Calculation = C
    Application.EnableEvents = EE
End Sub
Private Sub Fx1()
Call Fast(False, True, xlCalculationManual, False)
End Sub
Private Sub Fx2()
Call Fast(True, True, xlCalculationAutomatic, True)
End Sub

Private Sub App()
Fx1
    If Not ActiveCell.Offset(, 1) = "" Then
        Renungan.Locked = True
        Renungan = ActiveCell.Offset(, 1).Value
        Ayat.Locked = True
        Ayat = ActiveCell.Value
    Else
        With Renungan
            .Locked = False
            .Text = "Ponder"
        End With
        With Ayat
            .Locked = False
            .Text = ""
        End With
    End If
Fx2
End Sub


'''USERFORM FOR BVVIEWER'''

Private Sub BVer()
Dim C1 As Range
Dim C2 As Range
Dim v As String

Set C1 = ActiveCell
Set C2 = ActiveCell.Offset(, 1)
    On Error Resume Next
    v = C1.Value & Chr(10) _
    & C2.Value & Chr(10) & Chr(10) _
    & "Comments:" & Chr(10) _
    & C2.Comment.Text
    If Err.Number <> 0 Then
        On Error GoTo 0
        v = vbNullString
        v = C1.Value & Chr(10) _
        & C2.Value
        VViewer = v
        VViewer.Locked = True
    Else
        VViewer = v
        VViewer.Locked = True
    End If
Set C1 = Nothing
Set C2 = Nothing
v = vbNullString
End Sub

Private Sub BckUp_Click()
Dim objWord As Object
Dim i As Integer
Dim strValue As String
Dim RV As Range
Dim Lr As Variant
Const wdMove = 0
Const wdParagraph = 4
Const wdWord = 2
Const wdUnderlineSingle = 1
Const wdUnderlineNone = 0

    If Cells(2, 2) = "" Then Exit Sub
On Error GoTo bye
Set objWord = CreateObject("Word.Application")
    Lr = Cells(Rows.Count, 2).End(xlUp).Row
    Cells(2, 2).Select
    With objWord
        .Visible = True
        .Documents.Add
        For i = 0 To Lr - 2
            Set RV = Cells(ActiveCell.Row + i, 2)
            With RV
                If .Value <> "" Then
                    On Error Resume Next
                    strValue = _
                    RV & Chr(10) _
                    & RV.Offset(, 1) _
                    & Chr(10) _
                    & "Comments:" & Chr(10) _
                    & RV.Offset(, 1).Comment.Text
                    If Err.Number > 0 Then
                        On Error GoTo 0
                        strValue = vbNullString
                        strValue = _
                        RV & Chr(10) _
                        & RV.Offset(, 1)
                    End If
                End If
            End With
            With .Selection
                .Text = strValue & Chr(10) & Chr(10)
                With .Sentences.First
                    .Bold = True
                End With
                .StartOf 2, 0
                .MoveDown 4, 1
                .SelectCurrentAlignment
            End With
            With .Selection
                .Font.Italic = True
                .StartOf 2, 0
                .MoveDown 4, 1
                .SelectCurrentAlignment
            End With
            If .Selection.Characters.Count > 1 Then
                With .Selection
                    .Font.Italic = False
                    .Sentences.First.Bold = True
                    .StartOf 2, 0
                    .MoveDown 4, 1
                    .SelectCurrentAlignment
                End With
                With .Selection
                    .Font.Underline = 1
                    .EndOf 2, 0
                    .SelectCurrentAlignment
                End With
                .Selection.Font.Underline = 0
            Else
                With .Selection
                    .EndOf 2, 0
                    .Font.Italic = False
                End With
            End If
        Next i
        On Error Resume Next
        .Documents(1).Save
        .Quit False
    End With
bye:
    If Err.Number <> 0 And Err.Number <> 4198 Then
        MsgBox Err.Number
        MsgBox "Sorry, the Word Application is not" _
        & " available.", vbInformation, "Bible Verses Collections"
    End If
Set objWord = Nothing
i = 0
Lr = 0
strValue = vbNullString
Set RV = Nothing
End Sub

Private Sub NumberVP_Change()
    NumberVP.Locked = True
    SpinButton2.Value = NumberVP.Text
End Sub

Private Sub PrintV_Click()
Dim objWord As Object
Dim i As Integer
Dim strValue As String
Dim RV As Range
Const wdMove = 0
Const wdParagraph = 4
Const wdWord = 2
Const wdUnderlineSingle = 1
Const wdUnderlineNone = 0

    If ActiveCell = "" Then Exit Sub
On Error GoTo bye
Set objWord = CreateObject("Word.Application")
    With objWord
        .Visible = True
        .Documents.Add
        For i = 0 To NumberVP - 1
            Set RV = Cells(ActiveCell.Row + i, 2)
            With RV
                If .Value <> "" Then
                    On Error Resume Next
                    strValue = _
                    RV & Chr(10) _
                    & RV.Offset(, 1) _
                    & Chr(10) _
                    & "Comments:" & Chr(10) _
                    & RV.Offset(, 1).Comment.Text
                    If Err.Number > 0 Then
                        On Error GoTo 0
                        strValue = vbNullString
                        strValue = _
                        RV & Chr(10) _
                        & RV.Offset(, 1)
                    End If
                End If
            End With
            With .Selection
                .Text = strValue & Chr(10) & Chr(10)
                With .Sentences.First
                    .Bold = True
                End With
                .StartOf 2, 0
                .MoveDown 4, 1
                .SelectCurrentAlignment
            End With
            With .Selection
                .Font.Italic = True
                .StartOf 2, 0
                .MoveDown 4, 1
                .SelectCurrentAlignment
            End With
            If .Selection.Characters.Count > 1 Then
                With .Selection
                    .Font.Italic = False
                    .Sentences.First.Bold = True
                    .StartOf 2, 0
                    .MoveDown 4, 1
                    .SelectCurrentAlignment
                End With
                With .Selection
                    .Font.Underline = 1
                    .EndOf 2, 0
                    .SelectCurrentAlignment
                End With
                .Selection.Font.Underline = 0
            Else
                With .Selection
                    .EndOf 2, 0
                    .Font.Italic = False
                End With
            End If
        Next i
        With .Documents(1)
            .PrintOut
        End With
    End With
bye:
    If Err.Number <> 0 Then
        MsgBox "Sorry, the Word Application is not" _
        & " available.", vbInformation, "Bible Verses Collections"
    End If
Set objWord = Nothing
i = 0
strValue = vbNullString
Set RV = Nothing
End Sub

Private Sub SpinButton1_SpinDown()
    With ActiveCell
        If .Column = 2 Then
            If .Offset(-1).Address = "$B$1" Then
                Exit Sub
            Else
                VViewer.Locked = False
                .Offset(-1).Select
                BVer
            End If
        End If
    End With
End Sub

Private Sub SpinButton1_SpinUp()
    With ActiveCell
        If .Column = 2 Then
            If .Offset(1) = "" Then
                Exit Sub
            Else
                VViewer.Locked = False
                .Offset(1).Select
                BVer
            End If
        End If
    End With
End Sub

Private Sub SpinButton2_Change()
    NumberVP.Text = SpinButton2.Value
End Sub

Private Sub UserForm_Initialize()
    BVer
End Sub
