VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TauschenForm1 
   Caption         =   "Patient verschieben"
   ClientHeight    =   2570
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3670
   OleObjectBlob   =   "TauschenForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "TauschenForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Reset Form
Public Sub ResetForm()
    ZimmerInput1.Value = ""
    ZimmerInput2.Value = ""
    ZimmerInput1.SetFocus
End Sub

Private Sub TauschenButton_Click()
    TauschenForm1.Hide
End Sub
