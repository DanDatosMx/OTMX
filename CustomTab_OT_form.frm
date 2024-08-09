VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomTab_OT_form 
   Caption         =   "Data Management MX"
   ClientHeight    =   3324
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4320
   OleObjectBlob   =   "CustomTab_OT_form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "CustomTab_OT_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Owner_lbl_Click()

End Sub

Private Sub UserForm_Activate()

Me.Owner_lbl.Caption = "Oficina de Transformación México" & Chr(10) & "Data Management MX"
Me.Date_lbl.Caption = Year(Now())

End Sub
