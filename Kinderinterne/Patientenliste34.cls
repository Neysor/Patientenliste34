VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Patientenliste34
' Version 0.1
' BETA
'
' Autor: Dr. Helmut M. Steinbrecher MSc. BSc.
' Mail: Helmut.Steinbrecher@Neysor.net
'
' Lizenz: CC BY-SA 4.0
' Lizenzlink: https://creativecommons.org/licenses/by-sa/4.0/
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Bitte liebe KAV-IT, MA01 oder andere - lasst uns diese Liste :)
'
' Ich hasse VBA, so primitiv diese Sprache...  Aber es ist die
' einzige Möglichkeit im Alltag benötigte Prozesse zu automatisieren
' wenn schon wieder kein Geld für eine professionelle Umsetzung
' vorhanden ist, bzw. Ressourcen an wichtigere Projekte gebunden
' sind. Es ist doch zum weinen, dass wir mit dem Menüplan im Alltag
' auf allen Stationen arbeiten müssen...
'
' Der vorliegende Code greift nur lesend auf eine Webseite zu und
' befüllt und verändert sich selbst. Es gibt keine Schreibprozesse
' auf Dateien oder Ordner. Daher geht prinzipiell keine Gefahr von
' diesem Code aus, solange ihn niemand für seine Zwecke verändert
' hat. Daher bitte immer aus vertrauenswürdiger Quelle beziehen.
'
'   !!    ES WIRD KEINERLEI GARANTIE HIERFÜR ÜBERNOMMEN    !!
'   !!     ES HANDELT SICH HIER UM EIN PRIVATES PROJEKT    !!
'   !! ES WURDE IN DER FREIZEIT ERSTELLT - RESPEKTIERT DAS !!
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'ENUM für Readystate beim Ausleseprozess der Website
Enum READYSTATE
    READYSTATE_UNINITIALIZED = 0
    READYSTATE_LOADING = 1
    READYSTATE_LOADED = 2
    READYSTATE_INTERACTIVE = 3
    READYSTATE_COMPLETE = 4
End Enum

' Diese Funktion liest die Topasseite geeignet aus liefert
' eine Liste mit Patientendaten
Public Function runTopasRead(topasurl) As Variant
    'to refer to the running copy of Internet Explorer
    Dim ie As InternetExplorer
    'to refer to the HTML document returned
    Dim html As HTMLDocument
    'open Internet Explorer in memory, and go to website
    Set ie = New InternetExplorer
    ie.Visible = False
    'Get URL from Input
    ie.navigate Trim(topasurl)
    'Wait until IE is done loading page
    Do While ie.READYSTATE <> READYSTATE_COMPLETE
        Application.StatusBar = "Warte auf Daten ..."
        DoEvents
    Loop
    
    'show text of HTML document returned
    Set html = ie.Document
    Set patlist = html.getElementById("Liste_PatientListView")
    
    Dim Patientenliste() As patient
    '-2 weil 1. Reihe nicht zählt und weil länge 0 mitzählt
    ReDim Patientenliste(patlist.Rows.Length - 2)
    
    a = 0
    For Each itm In patlist.Rows
        Set pat = New patient
        If a = 0 Then
            'noop weil erstes element ausgelassen werden muss, das sind TH
        Else
            pat.Zimmernummer = ""
            pat.Name = itm.getElementsByTagName("td")(0).innerText
            pat.Geburtsdatum = itm.getElementsByTagName("td")(2).innerText
            pat.Geschlecht = itm.getElementsByTagName("td")(1).innerText
            pat.Aufnahmedatum = itm.getElementsByTagName("td")(5).innerText
            pat.Entlassungsdatum = itm.getElementsByTagName("td")(6).innerText
            Set Patientenliste(a - 1) = pat
        End If
        a = a + 1
    Next
    
    Set ie = Nothing
    
    runTopasRead = Patientenliste
End Function

'Diese Funktion holt sich alle Patientendaten aus dem Wordfile die
'bereits vorher ausgefüllt waren
Public Function runTableRead() As Variant
    Set pattable = ActiveDocument.Tables(1)
    Dim Patientenliste(), ReturnPatientenliste() As patient
    ReDim Patientenliste(pattable.Rows.Count - 1)
    
    a = 0 'laufvariable
    b = 0 'laufvariable
    For Each itm In pattable.Rows
        'Gehe alle durch und nehme nur leute auf mit Gebdatum in Spalte
        gebdat = itm.Cells(4).Range.Text
        If gebdat = Chr(13) & Chr(7) Then 'idiotischerweise sind diese 2 zeichen in einer leeren zelle drin
            'noop, weil leer
        Else
            'nicht leer
            Set pat = New patient
            pat.Zimmernummer = Left(itm.Cells(1).Range.Text, Len(itm.Cells(1).Range.Text) - 2)
            pat.Name = Left(itm.Cells(2).Range.Text, Len(itm.Cells(2).Range.Text) - 2)
            pat.Geburtsdatum = Left(itm.Cells(4).Range.Text, Len(itm.Cells(4).Range.Text) - 2)
            pat.Geschlecht = Left(itm.Cells(3).Range.Text, Len(itm.Cells(3).Range.Text) - 2)
            pat.Aufnahmedatum = ""
            pat.Entlassungsdatum = ""
            Set Patientenliste(b) = pat
            b = b + 1
        End If
        a = a + 1
    Next
    
    If (b = 0) Then
        'liste war leer
        runTableRead = Null
    Else
        'liste war bereits befüllt
        ReDim ReturnPatientenliste(b - 1)
        For i = 0 To b - 1
            Set ReturnPatientenliste(i) = Patientenliste(i)
        Next i
        runTableRead = ReturnPatientenliste
    End If
End Function

'Demo Daten zum Testen ausserhalb des Topas Bereiches
Public Function getDemoPatients() As Variant
    Dim Patientenliste(9) As patient
    
    For i = 0 To 9
        Set pat = New patient
        pat.Zimmernummer = ""
        pat.Name = "Demo Name Nr" & (i + 1)
        pat.Geburtsdatum = "1." & i & ".200" & i
        pat.Geschlecht = "W"
        pat.Aufnahmedatum = ""
        pat.Entlassungsdatum = ""
        Set Patientenliste(i) = pat
    Next i
    getDemoPatients = Patientenliste
    
End Function


'''EINSTIEGSPUNKT REFRESH BUTTON
Private Sub RefreshButton_Click()
    'Topas URL
    Dim dht, topasurl As String
    Application.StatusBar = "Starte Aktualisierung..."
    dht = Format(Date, "dd.mm.yyyy")
    topasurl = "http://topasportal.wienkav.at/Forms/Print/PrintStationsListe.aspx?sd=95631802&bd=" + dht + "&ed=" + dht + "&td=95631800&st=&mt=-1&xt=99999&ms=10&xs=90&si=0"
    
    'Topas Patientenliste
    'topaspatlist = runTopasRead(topasurl)
    
    'Topas Patientenliste Demo
    topaspatlist = getDemoPatients()
    
    'Tabelle Patientenliste
    tablepatlist = runTableRead()
    
    'Wenn ja, geht das Programm davon aus, dass alle Personen über 18 Eltern sind, wenn nein werden alle vorgeschlagen
    If MsgBox("Alle Patienten über 18 ignorieren als Eltern überspringen?", vbYesNo, "Elterncheck") = vbYes Then
        a = 0
        For Each pat In topaspatlist
            
        Next
    End If
    
    
    
    'TODO Listen abgleichen und einfügen
    'Entlassungen prüfen und zellen leeren
    'Neuaufnahmen hinzufügen
    'Andere Patienten unberührt lassen
        
    Application.StatusBar = "Fertig und viel Spaß - lg Dr. Helmut Steinbrecher MSc. BSc."
    
End Sub


'Tauscht 2 Zeilen miteinander
Public Sub ZeilenTausch(zimmer1, zimmer2)
        
    If (zimmer1 = zimmer2) Then
        'noop
    Else
        'hole tabelle
        Set pattable = ActiveDocument.Tables(1)
        zn1index = 0
        zn2index = 0
        
        
        For Each itm In pattable.Rows
            'Lese zimmernummer
            zn = Left(itm.Cells(1).Range.Text, Len(itm.Cells(1).Range.Text) - 2)
            If zn = zimmer1 Then
                zn1index = itm.Index
                zn1text = zn
            ElseIf zn = zimmer2 Then
                zn2index = itm.Index
                zn2text = zn
            End If
        Next
        
        If (zn1index <> 0 And zn2index <> 0) Then
            'Sort index higher first
                
            znhigh = IIf(zn1index > zn2index, zn1index, zn2index)
            znhight = IIf(zn1index > zn2index, zn1text, zn2text)
            znlow = IIf(zn1index > zn2index, zn2index, zn1index)
            znlowt = IIf(zn1index > zn2index, zn2text, zn1text)
            
            'Zimmertext ändern
            pattable.Rows(znhigh).Range.Text = znlowt
            pattable.Rows(znlow).Range.Text = znhight
            
            Application.CutCopyMode = False
            Application.ScreenUpdating = False
            
            'Cut / Paste
            pattable.Rows(znhigh).Range.Cut
            pattable.Rows(znlow).Range.PasteAndFormat (wdFormatOriginalFormatting)
            
            'Cut / Paste
            pattable.Rows(znlow + 1).Range.Cut
            If (znhigh > pattable.Rows.Count) Then
                'dirty as hell
                pattable.Rows(pattable.Rows.Count).Range.PasteAndFormat (wdFormatOriginalFormatting)
                pattable.Rows(pattable.Rows.Count).Range.Cut
                pattable.Rows(pattable.Rows.Count).Range.PasteAndFormat (wdFormatOriginalFormatting)
            Else
                pattable.Rows(znhigh).Range.PasteAndFormat (wdFormatOriginalFormatting)
            End If
            
            Application.CutCopyMode = False
            Application.ScreenUpdating = True
            
        Else
            ret = MsgBox("Konnte eines der Zimmer nicht finden. Vertippt?" & vbNewLine & "Abbruch", vbExclamation, "Fehler")
        End If
    End If
End Sub

'''EINSTIEGSPUNKT ZIMMER TAUSCHEN BUTTON
Private Sub ZimmerTauschButton_Click()
    'Zimmer tauchen
    TauschenForm1.ResetForm
    TauschenForm1.Show

    If TauschenForm1.ZimmerInput1.TextLength > 0 And TauschenForm1.ZimmerInput2.TextLength > 0 Then
        Call ZeilenTausch(TauschenForm1.ZimmerInput1.Value, TauschenForm1.ZimmerInput2.Value)
    Else
        ret = MsgBox("Eingabefehler bei Zimmernummern" & vbNewLine & "Abbruch", vbExclamation, "Fehler")
    End If
End Sub
