VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4390
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   7100
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bahaye_vahed_Change()
If tedad.Value <> "" And bahaye_vahed <> "" And takhfif.Value <> "" Then
mablagh_kol.Value = (tedad.Value * bahaye_vahed.Value) - takhfif.Value
End If
End Sub

Private Sub kharidar_code_Change()
Range("l1").Value = kharidar_code.Value
kharidar_name.Value = Range("m1").Value
tamas.Value = Range("n1").Value
address.Value = Range("o1").Value

End Sub

Private Sub taed_click()

Range("h1").Value = shomare_factor.Value
Range("h2").Value = tarikh_factor.Value
Range("c7").Value = kharidar_name.Value
Range("f7").Value = kharidar_code.Value
Range("h7").Value = tamas.Value
Range("c8").Value = address.Value
Cells(Range("j1").Value, 2) = kala_code.Value
Cells(Range("j1").Value - 1, 3) = kala_name.Value
Cells(Range("j1").Value - 1, 4) = tedad.Value
Cells(Range("j1").Value - 1, 5) = bahaye_vahed.Value
Cells(Range("j1").Value - 1, 6) = takhfif.Value
Cells(Range("j1").Value - 1, 7) = mablagh_kol.Value

shomare_factor.Value = ""
tarikh_factor.Value = ""
kharidar_name.Value = ""
kharidar_code.Value = ""
tamas.Value = ""
address.Value = ""
kala_code.Value = ""
kala_name.Value = ""
tedad.Value = ""
bahaye_vahed.Value = ""
takhfif.Value = ""
mablagh_kol.Value = ""
End Sub


Private Sub takhfif_Change()
If tedad.Value <> "" And bahaye_vahed <> "" And takhfif.Value <> "" Then
mablagh_kol.Value = (tedad.Value * bahaye_vahed.Value) - takhfif.Value
End If
End Sub

Private Sub tedad_Change()
If tedad.Value <> "" And bahaye_vahed <> "" And takhfif.Value <> "" Then
mablagh_kol.Value = (tedad.Value * bahaye_vahed.Value) - takhfif.Value
End If
End Sub


