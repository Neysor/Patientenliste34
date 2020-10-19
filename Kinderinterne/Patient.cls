VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Klasse Patient
Public Zimmernummer As String
Public Name As String
Public Geburtsdatum As String
Public Geschlecht As String
Public Aufnahmedatum As String
Public Entlassungsdatum As String

'Gibt Alter als String retuor
Public Function getAlterByGeb() As String
    AlterM = DateDiff("m", DateValue(Geburtsdatum), Date)
    AlterW = DateDiff("ww", DateValue(Geburtsdatum), Date)
    AlterD = DateDiff("d", DateValue(Geburtsdatum), Date)
    
    AlterY = Math.Round(AlterM / 12, 1)
    AlterM2 = Math.Round(AlterW / 4, 1)
    AlterW2 = Math.Round(AlterD / 7, 1)
    
    ' Bis 1 Monat alles in Wochen angeben
    ' Bis 17 Monate, alles in Monaten angeben
    ' Ab dem 18. Monat (=1,5a) alles in Jahren angeben
    If AlterM2 < 1 Then
        getAlterByGeb = "~" & AlterW2 & "w"
    ElseIf AlterM2 < 18 Then
        getAlterByGeb = "~" & AlterM2 & "m"
    Else
        
        getAlterByGeb = "~" & AlterY & "a"
    End If
End Function

'Gibt true wenn Erwachsen und false wenn Kind
Public Function isAdult() As Boolean
    AlterM = DateDiff("m", DateValue(Geburtsdatum), Date)
    AlterY = Math.Round(AlterM / 12, 1)
    If AlterY > 18 Then
        isAdult = True
    Else
        isAdult = False
    End If
End Function
