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
' Version 0.6
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
    
    ''Weil KAV etwas am IE gemacht hat, muss ich zuerst alle schliessen
    ''Dann muss ich das Fenster oeffnen
    ''Dann das Fenster im System wiederfinden wenn es fertig geladen ist
    ''Dann erst kann ich es wieder auslesen
    
    'schliesse alle internet explorer
    For Each GetIE In CreateObject("Shell.Application").Windows() 'Loop to find
        If (Not GetIE Is Nothing) And GetIE.Name = "Internet Explorer" Then
            GetIE.Quit
            Exit For
        End If
    Next GetIE
        
    'oeffne ein neues fenster
    Set ie = New InternetExplorer
    ie.navigate Trim(topasurl)
    'ie referenz verloren wegen behinderter KAV IT
    
    MsgBox "Bitte druecken wenn Internet Explorer sichtbar und fertig geladen"
        
    For Each myIE In CreateObject("Shell.Application").Windows() 'Loop to find
        If (Not myIE Is Nothing) And myIE.Name = "Internet Explorer" Then Exit For
    Next myIE
    
    Set html = myIE.Document
    
    'Netterweise hat der KAV die Liste gleich schön mit einer ID versehen
    Set patlist = html.getElementById("Liste_PatientListView")
    
    Dim Patientenliste() As Patient
    '-2 weil 1. Reihe nicht zählt und weil länge 0 mitzählt
    ReDim Patientenliste(patlist.Rows.Length - 2)
    
    a = 0
    For Each itm In patlist.Rows
        Set pat = New Patient
        If a = 0 Then
            'noop weil erstes element ausgelassen werden muss, das sind TH
        Else
            pat.Zimmernummer = ""
            pat.Name = itm.getElementsByTagName("td")(0).innerText
            pat.Geburtsdatum = itm.getElementsByTagName("td")(2).innerText
            pat.Geschlecht = itm.getElementsByTagName("td")(1).innerText
            pat.Aufnahmedatum = itm.getElementsByTagName("td")(5).innerText
            pat.Entlassungsdatum = Trim(itm.getElementsByTagName("td")(6).innerText)
            Set Patientenliste(a - 1) = pat
        End If
        a = a + 1
    Next
    
    'wichtig fürs resourcenfreigen !
    Set ie = Nothing
    
    runTopasRead = Patientenliste
End Function

'Diese Funktion holt sich alle Patientendaten aus dem Wordfile die
'bereits vorher ausgefüllt waren
Public Function runTableRead() As Variant
    Set pattable = ActiveDocument.Tables(1)
    Dim Patientenliste(), ReturnPatientenliste() As Patient
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
            Set pat = New Patient
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
    Dim Patientenliste(9) As Patient
    
    For i = 0 To 9
        Set pat = New Patient
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


'leert eine zeile = entässt einen Patienten
Public Sub entlassePatient(zimmernr)
    'hole tabelle
    Set pattable = ActiveDocument.Tables(1)
    znindex = 0
    
    For Each itm In pattable.Rows
        'Lese zimmernummer
        zn = Left(itm.Cells(1).Range.Text, Len(itm.Cells(1).Range.Text) - 2)
        If zn = zimmernr Then
            znindex = itm.Index
            Exit For
        End If
    Next
    
    If (znindex <> 0) Then
        'Zeile auswählen
        pattable.Rows(znindex).Select
        'Zeile leeren
        Selection.Delete
        'Zimmernummer neu eintragen
        pattable.Rows(znindex).Range.Text = zimmernr
    Else
        'sollte nicht passieren
    End If
End Sub

'fügt einen Patient in ein Zimmer hinzu true wenn funktioniert hat
Public Function aufnahmePatient(zimmernr, pat) As Boolean
    aufnahmePatient = True
    Set pattable = ActiveDocument.Tables(1)
    znindex = 0
    
    For Each itm In pattable.Rows
        'Lese zimmernummer
        zn = Left(itm.Cells(1).Range.Text, Len(itm.Cells(1).Range.Text) - 2)
        If zn = zimmernr Then
            znindex = itm.Index
            Exit For
        End If
    Next
    
    If (znindex <> 0) Then
        
        'Zeile auswählen
        pattable.Rows(znindex).Select
        'Zeile leeren
        Selection.Delete
        
        'Zimmernummer neu eintragen
        pattable.Rows(znindex).Cells(1).Range.Text = zimmernr
        'Name eintragen
        pattable.Rows(znindex).Cells(2).Range.Text = pat.Name
        'Geschlecht eintragen
        pattable.Rows(znindex).Cells(3).Range.Text = pat.Geschlecht
        'Geburtsdatum eintragen
        pattable.Rows(znindex).Cells(4).Range.Text = pat.Geburtsdatum
        'Alter eintragen
        pattable.Rows(znindex).Cells(5).Range.Text = pat.getAlterByGeb()
    Else
        'kein Zimmer gefunden
        aufnahmePatient = False
    End If
End Function


'''EINSTIEGSPUNKT REFRESH BUTTON
Private Sub RefreshButton_Click()
    'Topas URL
    Dim dht, topasurl As String
    Application.StatusBar = "Starte Aktualisierung..."
    dht = Format(Date, "dd.mm.yyyy")
    'Das ist die Topas Drucken URL mit dem aktuellen Datum
    topasurl = "http://topasportal.wienkav.at/Forms/Print/PrintStationsListe.aspx?sd=95631601&bd=" + dht + "&ed=" + dht + "&td=95631600&st=&mt=-1&xt=99999&ms=10&xs=90&si=0"
    
    'Topas Patientenliste
    topaspatlist = runTopasRead(topasurl)
    
    'Topas Patientenliste Demo
    'topaspatlist = getDemoPatients()
    
    'Tabelle Patientenliste
    tablepatlist = runTableRead()
    
    'Patienten entlassen?
    entlassen = IIf((MsgBox("Entlassungen durchführen?", vbYesNo, "Entlassungen") = vbYes), True, False)
    
    'wenn Patienten in der Tabelle sind
    If Not (IsNull(tablepatlist)) Then
        'Gehe alle Patienten in der Tabelle durch ob sie auch in der Topas Liste vorkommen
        'aber nur entlassen wenn entlassen = true!
        'Die liste muss aber dennoch durchlaufen werden, da sonst bei den Neuaufnahmen der Abgleich nicht funktioniert
        
        For Each pat In tablepatlist
            
            'Suche den Patient in Liste
            tpindex = -1
            i = 0
            For Each pata In topaspatlist
                If pat.Name = pata.Name And pat.Geburtsdatum = pata.Geburtsdatum Then
                    tpindex = i
                    Exit For
                End If
                i = i + 1
            Next
            
            If tpindex < 0 Then
                'Patient konnte nicht in der Liste gefunden werden
                If entlassen Then
                    Call entlassePatient(pat.Zimmernummer)
                End If
            ElseIf topaspatlist(tpindex).Entlassungsdatum <> "" Then
               'Es steht ein Entlassungsdatum im Topas
                If entlassen Then
                    Call entlassePatient(pat.Zimmernummer)
                End If
            Else
                'Der Patient steht noch im Topas ohne Entlassungsdatum
                'wir fügen daher die bekannte zimmernummer hinzu
                topaspatlist(tpindex).Zimmernummer = pat.Zimmernummer
            End If
            
        Next
        
    End If
    
    'Wenn ja, geht das Programm davon aus, dass alle Personen über 18 Eltern sind, wenn nein werden alle vorgeschlagen
    ignore18 = False
    
    'Neuaufnahmen hinzufügen
    For Each pat In topaspatlist
        If pat.Zimmernummer <> "" Then
            'nichts tun, da schon in liste
        ElseIf ignore18 And pat.isAdult Then
            'nichts tun da alle über 18 ignoriert werden sollen und patient über 18 ist
        ElseIf pat.Entlassungsdatum <> "" Then
            'Entlassungsdatum ausgefüllt
        Else
            'nochoffen hilft bei Fehleingaben den Patienten noch einmal anzuzeigen.
            nochoffen = True
            
            'nur wenn das Fenster mit X oder ignorieren geschlossen wird, kommt der nächste Patient
            'wenn ablegen geklickt wird und ein fehler passiert, dann ist das eine Endlosschleife :)
            Do While nochoffen
                
                PatientenAblageForm1.ResetForm
                
                PatientenAblageForm1.PatLabel = PatientenAblageForm1.PatLabel & pat.Name
                PatientenAblageForm1.GebLabel = PatientenAblageForm1.GebLabel & pat.Geburtsdatum & " -> " & pat.getAlterByGeb()
                PatientenAblageForm1.SexLabel = PatientenAblageForm1.SexLabel & pat.Geschlecht
                PatientenAblageForm1.AufLabel = PatientenAblageForm1.AufLabel & pat.Aufnahmedatum
                PatientenAblageForm1.EntLabel = PatientenAblageForm1.EntLabel & pat.Entlassungsdatum
                
                PatientenAblageForm1.Show
                
                If PatientenAblageForm1.retval Then
                    'Ablegen
                    If PatientenAblageForm1.ZimmerTextBox1.TextLength > 0 Then
                        nochoffen = Not (aufnahmePatient(PatientenAblageForm1.ZimmerTextBox1.Value, pat))
                    Else
                        nochoffen = True
                        ret = MsgBox("Kein Zimmer ausgefüllt" & vbNewLine & "Neuer Versuch", vbExclamation, "Fehler")
                    End If
                Else
                    'Ignorieren
                        nochoffen = False
                End If
                
            Loop
        End If
    Next
        
    Application.StatusBar = "Fertig und viel Spaß - lg Helmut"
    
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
            'Wichtig da beim ausschneiden und einfügen ja dynamisch die Tabellengröße
            'und der index verschoben werden
            znhigh = IIf(zn1index > zn2index, zn1index, zn2index)
            znhight = IIf(zn1index > zn2index, zn1text, zn2text)
            znlow = IIf(zn1index > zn2index, zn2index, zn1index)
            znlowt = IIf(zn1index > zn2index, zn2text, zn1text)
            
            'Zimmertext ändern
            pattable.Rows(znhigh).Range.Text = znlowt
            pattable.Rows(znlow).Range.Text = znhight
            
            'Cut / Paste
            pattable.Rows(znhigh).Range.Cut
            pattable.Rows(znlow).Range.PasteAndFormat (wdFormatOriginalFormatting)
            
            'Cut / Paste
            pattable.Rows(znlow + 1).Range.Cut
            If (znhigh > pattable.Rows.Count) Then
                'dirty as hell but this is Word...
                pattable.Rows(pattable.Rows.Count).Range.PasteAndFormat (wdFormatOriginalFormatting)
                pattable.Rows(pattable.Rows.Count).Range.Cut
                pattable.Rows(pattable.Rows.Count).Range.PasteAndFormat (wdFormatOriginalFormatting)
            Else
                pattable.Rows(znhigh).Range.PasteAndFormat (wdFormatOriginalFormatting)
            End If
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

Private Sub EntlassenButton_Click()
    'Nahezu copy und paste von aktualisieren, kleine änderungen weil hier nur entlassen werden soll
    Application.StatusBar = "Starte Entlassungen..."
    'Topas URL
    Dim dht, topasurl As String

    dht = Format(Date, "dd.mm.yyyy")
    'Das ist die Topas Drucken URL mit dem aktuellen Datum
    topasurl = "http://topasportal.wienkav.at/Forms/Print/PrintStationsListe.aspx?sd=95631601&bd=" + dht + "&ed=" + dht + "&td=95631600&st=&mt=-1&xt=99999&ms=10&xs=90&si=0"
    
    'Topas Patientenliste
    topaspatlist = runTopasRead(topasurl)
    
    'Tabelle Patientenliste
    tablepatlist = runTableRead()
    
    'wenn Patienten in der Tabelle sind
    If Not (IsNull(tablepatlist)) Then
        'Gehe alle Patienten in der Tabelle durch ob sie auch in der Topas Liste vorkommen
        
        For Each pat In tablepatlist
            
            'Suche den Patient in Liste
            tpindex = -1
            i = 0
            For Each pata In topaspatlist
                If pat.Name = pata.Name And pat.Geburtsdatum = pata.Geburtsdatum Then
                    tpindex = i
                    Exit For
                End If
                i = i + 1
            Next
            
            If tpindex < 0 Then
                'Patient konnte nicht in der Liste gefunden werden
                Call entlassePatient(pat.Zimmernummer)
            ElseIf topaspatlist(tpindex).Entlassungsdatum <> "" Then
               'Es steht ein Entlassungsdatum im Topas
                Call entlassePatient(pat.Zimmernummer)
            Else
                'noop
            End If
        Next
    End If
    Application.StatusBar = "Entlassungen durchgeführt - lg Helmut Steinbrecher"
End Sub
Private Sub InfoButton1_Click()
    InfoForm.Show
End Sub
