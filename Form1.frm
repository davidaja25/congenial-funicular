VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rinci Kembalian"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1320
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Text            =   " "
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rincian Kembalian :"
      Height          =   2895
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   3735
      Begin VB.Frame Frame3 
         Caption         =   "Rincian Kembalian :"
         Height          =   2895
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   3735
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   1080
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "RESET"
            Height          =   495
            Left            =   2880
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   1080
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "buah"
            Height          =   375
            Left            =   2400
            TabIndex        =   39
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "buah"
            Height          =   375
            Left            =   2400
            TabIndex        =   38
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "buah"
            Height          =   375
            Left            =   2400
            TabIndex        =   37
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "buah"
            Height          =   375
            Left            =   2400
            TabIndex        =   36
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "buah"
            Height          =   375
            Left            =   2400
            TabIndex        =   35
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "buah"
            Height          =   375
            Left            =   2400
            TabIndex        =   34
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Rp. 50.000,-"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Rp. 20.000,-"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Rp. 10.000,-"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Rp. 5.000,-"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Rp. 2.000,-"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Rp. 500,-"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Sisa Satuan"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   2520
            Width           =   1095
         End
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "RESET"
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Sisa Satuan"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Rp. 500,-"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Rp. 2.000,-"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Rp. 5.000,-"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Rp. 10.000,-"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Rp. 20.000,-"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Rp. 50.000,-"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   " "
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaksi"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Total Kembalian "
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Total Bayar"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Total Belanja"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim kembalian As Currency

Call hitung

End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = 0
    Text12.Text = 0
    Text5.Text = 0
    Text6.Text = 0
    Text7.Text = 0
    Text8.Text = 0
    Text11.Text = 0
    Text10.Text = 0
End Sub

Private Sub Form_Load()
 Text1.Text = ""
    Text2.Text = ""
    Text3.Text = 0
    Text12.Text = 0
    Text5.Text = 0
    Text6.Text = 0
    Text7.Text = 0
    Text8.Text = 0
    Text11.Text = 0
    Text10.Text = 0

End Sub

Function hitung()
Text3.Text = Text2.Text - Text1.Text
kembalian = Val(Text3.Text)
Text10.Text = -1

Do Until Val(Text10.Text) + 1 > 0
    If kembalian >= 50000 Then
        Text12.Text = kembalian \ 50000
        kembalian = kembalian Mod 50000
    ElseIf kembalian >= 20000 Then
        Text5.Text = kembalian \ 20000
        kembalian = kembalian Mod 20000
    ElseIf kembalian >= 10000 Then
        Text6.Text = kembalian \ 10000
        kembalian = kembalian Mod 10000
    ElseIf kembalian >= 5000 Then
        Text7.Text = kembalian \ 5000
        kembalian = kembalian Mod 5000
    ElseIf kembalian >= 2000 Then
        Text8.Text = kembalian \ 2000
        kembalian = kembalian Mod 2000
    ElseIf kembalian >= 500 Then
        Text11.Text = kembalian \ 500
        kembalian = kembalian Mod 500
    Else
        Text10.Text = kembalian 'sisa pecahan
    End If
 Loop
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call hitung
End If
End Sub
