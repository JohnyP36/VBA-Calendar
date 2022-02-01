VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blad2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub HandmatigCelInhoudVerwijderen()
    ' Zorgt ervoor dat het scherm niet knippert (je ziet niet elke handeling die de macro uitvoert)
'   Application.ScreenUpdating = False
    Range("B5:E35").Select
        Selection.ClearContents
    Range("A4").Select
End Sub
