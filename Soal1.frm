VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Total Element Array"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   300
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Nilai"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "HITUNG"
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total(), i, angka As Integer

Private Sub Command1_Click()
angka = CInt(Combo1.Text)
Call hitungArray
End Sub

Private Sub Command2_Click()
ReDim Total(1 To 5)
List1.Clear
Combo1.Clear

For i = 1 To 5
    Combo1.AddItem i
    List1.AddItem "Hasil Penjumlahan data Array : (" & i & ") = " & Total(i)
Next i

End Sub

Private Sub Form_Load()
ReDim Total(1 To 5)
For i = 1 To 5
    Combo1.AddItem i
    List1.AddItem "Hasil Penjumlahan data Array : (" & i & ") = " & Total(i)
Next i
Combo1.ListIndex = 0
End Sub

Function hitungArray()
angka = CInt(Combo1.Text)
Total(angka) = 15 - angka
If Total(angka) <> "" Then
    List1.Clear
    For i = 1 To UBound(Total)
        List1.AddItem "Hasil Penjumlahan data Array : (" & i & ") = " & Total(i)
    Next i
End If
End Function
