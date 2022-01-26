VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blad1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SelecteerDatumVandaag()
Dim Vandaag As Range
    Recent = Date
Nogeens:
    Set Vandaag = Range("B:B").Find(What:=Recent, LookAt:=xlWhole)
    If Vandaag Is Nothing Then
        Recent = Recent + 1
        GoTo Nogeens
        If Recent > Date + 7 Then
            MsgBox "Geen datum kunnen vinden in kolom B binnen de komende zeven dagen."
            ActiveCell.Offset(0, 0).Range("A1").Select
            'of -> Range("B1000").End(xlUp).Select
                'Selecteert de laatste cel in kolom B (geldt tot B1000)
            End
        End If
    End If
    Vandaag.Select
    Cells.Find(What:="zo", After:=ActiveCell, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Select
'   ActiveCell.Offset(-6, -1).Range("A1").Select
'       Gaat 1 cel naar links en 6 omhoog
    Application.Goto Reference:=ActiveCell, Scroll:=True
'   ActiveWindow.SmallScroll ToRight:=-20
'    of ->  ActiveCell.Offset(0, -1).Range("A1").Select
    Vandaag.Select
    ActiveCell.Offset(0, -1).Range("A1").Select
End Sub

'Below is to change height cells
'Sub sbChangeRowHeightMulti()
'Rows(2881).RowHeight = 6.8
'End Sub

'Sub ChangeFontHeight()
'    Range("A26,A34,A42").Select
'    With Selection.Font
'        .Size = 7.5
'    End With
'End Sub
