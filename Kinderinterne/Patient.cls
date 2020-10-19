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

'Dirty fill 1
Public Function fastFill1(v1, v2, v3, v4, v5, v6)
    Zimmernummer = v1
    Name = v2
    Geburtsdatum = v3
    Geschlecht = v4
    Aufnahmedatum = v5
    Entlassungsdatum = v6
End Function

'Dirty fill 2
Public Function fastFill2(v1, v2, v3, v4, v5, v6, v7)
    Infos = v7
    Call fastFill1(v1, v2, v3, v4, v5, v6)
End Function
'Gibt Alter als String retuor
Public Function getAlterByGeb() As String
    alterY = DateDiff("yyyy", Date, DateValue(Geburtsdatum))
    alterM = DateDiff("m", Date, DateValue(Geburtsdatum))
    alterW = DateDiff("ww", Date, DateValue(Geburtsdatum))
    
    ' Bis 1 Monat alles in Wochen angeben
    ' Bis 17 Monate, alles in Monaten angeben
    ' Ab dem 18. Monat (=1,5a) alles in Jahren angeben
    If alterM < 1 Then
        getAlterGeb = "~" & alterW & "w"
    ElseIf alterM < 18 Then
        getAlterGeb = "~" & alterM & "m"
    Else
        getAlterGeb = "~" & alterY & "a"
    End If
End Function

'Gibt true wenn Erwachsen und false wenn Kind
Public Function isAdult() As Boolean
    alterY = DateDiff("yyyy", Date, DateValue(Geburtsdatum))
    If alterY > 18 Then
        isAdult = True
    Else
        isAdult = False
    End If
End Function
