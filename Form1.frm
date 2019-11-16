VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Potongan Harga"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Potongan Diskon"
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "RESET"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Rp."
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Rp."
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Rp."
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Total Kembalian"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Diskon"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Total Pembayaran"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaksi"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Kode Voucher"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Total Belanja"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function voucherAsik()
    Text2.Text = "DUMBWAYSASIK"
    Text3.Text = 50 / 100 * Val(Text1.Text)
    If Val(Text3.Text) <= 20000 Then
        Text4.Text = Val(Text1.Text - Text3.Text)
    Else
        Text3.Text = 20000
        Text4.Text = Val(Text1.Text) - 20000
    End If
    Text5.Text = Val(Text1.Text - Text4.Text)
    Label6.Caption = "50% (maks.20.000,-)"

End Function
Function voucherMantap()
    Text2.Text = "DUMBWAYSMANTAP"
    Text3.Text = 30 / 100 * Val(Text1.Text)
    If Val(Text3.Text) <= 40000 Then
        Text4.Text = Val(Text1.Text - Text3.Text)
    Else
        Text3.Text = 40000
        Text4.Text = Val(Text1.Text) - 40000
    End If
    Text5.Text = Val(Text1.Text - Text4.Text)
    Label6.Caption = "30% (maks.40.000,-)"

End Function
Function noVoucher()
    Text2.Text = " - "
    Text3.Text = 0
    Text4.Text = Text1.Text
    Text5.Text = Text4.Text
    Label6.Caption = " - "

End Function

Private Sub Command1_Click()
     If Val(Text1.Text) < 20000 Then
        Call noVoucher
    ElseIf Val(Text1.Text) < 50000 Then
        Call voucherAsik
    ElseIf Val(Text1.Text) >= 50000 Then
        Call voucherMantap
    End If

End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = 0
    Text4.Text = 0
    Text5.Text = 0
    Label6.Caption = ""
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = 0
    Text4.Text = 0
    Text5.Text = 0
    Label6.Caption = ""
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     If Val(Text1.Text) < 20000 Then
        Call noVoucher
    ElseIf Val(Text1.Text) < 50000 Then
        Call voucherAsik
    ElseIf Val(Text1.Text) >= 50000 Then
        Call voucherMantap
    End If
End If
End Sub
