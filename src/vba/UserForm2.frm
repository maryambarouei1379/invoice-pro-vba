VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   4660
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   8520
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sabt_Click()
'Range("a2") = moshtary_code.Value
'Range("b2") = code_melli.Value
'Range("c2") = last_name.Value
'Range("d2") = ostan.Value
'Range("e2") = address.Value
Cells(Range("g1").Value, 1).Value = moshtary_code.Value
Cells(Range("g1").Value - 1, 2).Value = code_melli.Value
Cells(Range("g1").Value - 1, 3).Value = last_name.Value
Cells(Range("g1").Value - 1, 4).Value = ostan.Value
Cells(Range("g1").Value - 1, 5).Value = address.Value
Cells(Range("g1").Value - 1, 6).Value = tamas.Value
moshtary_code.Value = ""
code_melli.Value = ""
last_name.Value = ""
ostan.Value = ""
address.Value = ""
tamas.Value = ""


End Sub

