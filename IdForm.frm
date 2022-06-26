VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IdForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6972
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7224
   OleObjectBlob   =   "IdForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IdForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0

Private Sub Save_Click()
Dim book As Workbook, pg As Worksheet, last_line As Long
Dim nm As String, sur_nm As String, adr As String, phn As String

'Set book = Workbooks("VBA General Introduction.xlsm")
Set book = Workbooks(ThisWorkbook.name)
Set pg = book.Worksheets("New User")
last_line = pg.Range("A" & Rows.Count).End(xlUp).Row + 1

nm = Me.TextBox1.Value
sur_nm = Me.TextBox2.Value
adr = Me.TextBox3.Value
phn = Me.TextBox4.Value

array_list = Array("", nm, sur_nm, adr, phn)

For i = 1 To 4
    pg.Cells(last_line, i).Value = array_list(i)
Next i

Call Clear_Click
MsgBox "Saved"
Me.TextBox1.SetFocus

Set book = Nothing
Set pg = Nothing
End Sub

Private Sub Clear_Click()
Dim i As Integer

For i = 1 To 4
    Me.Controls("TextBox" & i).Value = ""
    'Space(1)
    'MsgBox "TextBox" & i & " is clear"
Next i

End Sub


Private Sub UserForm_Initialize()
    'IdForm
    Me.Caption = "Hello Earth"
    
End Sub
