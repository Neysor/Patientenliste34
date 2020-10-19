VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PatientenAblageForm1 
   Caption         =   "PatientenAblage"
   ClientHeight    =   4500
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   4540
   OleObjectBlob   =   "PatientenAblageForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "PatientenAblageForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public retval As Boolean

'Reset Form
Public Sub ResetForm()
    PatLabel.Caption = "Patient: "
    GebLabel.Caption = "Geburtsdatum: "
    SexLabel.Caption = "Geschlecht: "
    AufLabel.Caption = "Aufnahmedatum: "
    EntLabel.Caption = "Entlassungsdatum: "
    ZimmerTextBox1.Value = ""
    ZimmerTextBox1.SetFocus
    retval = False
End Sub

Private Sub AblegenCommandButton1_Click()
    retval = True
    PatientenAblageForm1.Hide
End Sub

Private Sub IgnoreCommandButton2_Click()
    retval = False
    PatientenAblageForm1.Hide
End Sub
