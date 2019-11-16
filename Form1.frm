VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Function Cetak Gambar Pola"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

a = Int(Text1.Text)

List1.Clear
For x = 1 To a
    If x Mod 2 = 0 Then
        Call Cetak_barisgenap
    Else
        Call Cetak_barisganjil
    End If
Next
End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub

Function Cetak_barisganjil()
a = Int(Text1.Text)
For x = 1 To a
  b = "= "
  c = " * "
  If x < 2 Then
  d = b
  ElseIf x < 3 Then
  d = b + c
  ElseIf x < 4 Then
  d = b + c + b
  ElseIf x < 5 Then
  d = b + c + b + c
  ElseIf x < 6 Then
  d = b + c + b + c + b
  End If
Next
List1.AddItem d
End Function

Function Cetak_barisgenap()
a = Int(Text1.Text)
For x = 1 To a
  b = "= "
  c = " * "
  If x < 2 Then
  e = c
  ElseIf x < 3 Then
  e = c + b
  ElseIf x < 4 Then
  e = c + b + b
  ElseIf x < 5 Then
  e = c + b + b + b
  ElseIf x < 6 Then
  e = c + b + b + b + c
  End If
Next
List1.AddItem e
End Function
