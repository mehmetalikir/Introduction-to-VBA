VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IdForm2 
   Caption         =   "UserForm1"
   ClientHeight    =   7968
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9468.001
   OleObjectBlob   =   "IdForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IdForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
Dim last_row As Long, sh As Worksheet, wb As Workbook

Set wb = Workbooks(ThisWorkbook.Name)
Set sh = wb.Worksheets("New User")
last_row = sh.Range("A" & Rows.Count).End(xlUp).Row
sh.Activate

With Me.ListBox1
    .ColumnHeads = False
    .ColumnCount = 4
    .ColumnWidths = "90;110;70;80"
    .RowSource = "A2:D" & last_row
End With

Set wb = Nothing
Set sh = Nothing


End Sub
Private Sub ListBox1_Click()

End Sub

Private Sub Save_Data_Click()
Dim wb As Workbook, sh As Worksheet, last_row As Long, i As Integer
Dim nm As String, snm As String, addr As String, phn As String

Set wb = Workbooks(ThisWorkbook.Name)
Set sh = wb.Worksheets("New User")
last_row = sh.Range("A" & Rows.Count).End(3).Row + 1

nm = Me.TextBox1.Value
snm = Me.TextBox2.Value
addr = Me.TextBox3.Value
phn = Me.TextBox4.Value

sh.Activate
sh.Range("A" & last_row).Value = nm
sh.Range("B" & last_row).Value = snm
sh.Range("C" & last_row).Value = addr
sh.Range("D" & last_row).Value = phn

'ThisWorkbook.Save

wb.Save

UserForm_Initialize

Call Clear_Data_Click

'Me.TextBox1.SetFocus

Set wb = Nothing
Set sh = Nothing

End Sub
Private Sub Clear_Data_Click()
For i = 1 To 4
    Me.Controls("TextBox" & i).Value = ""
Next i
Me.TextBox1.SetFocus
End Sub

