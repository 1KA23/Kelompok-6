VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000080&
   Caption         =   "Form Login"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   1215
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         Caption         =   "KLINIK APAYA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   5
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      TabIndex        =   4
      Top             =   7920
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   5160
      Width           =   3135
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   1
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "KLINIK APAYA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   11280
      TabIndex        =   6
      Top             =   1920
      Width           =   15
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3360
      TabIndex        =   0
      Top             =   3480
      Width           =   2145
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Terbuka()

frmMenuUtama.MnRegister.Enabled = False
frmMenuUtama.MnLogin.Enabled = False
frmMenuUtama.MnLogout.Enabled = True
frmMenuUtama.MnPendaftaran.Enabled = True
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLogin_Click()
Call CariData
End Sub

Private Sub Form_Activate()
txtUsername = ""
txtPassword = ""
txtUsername.SetFocus
End Sub

Function CariData()
Call BukaDB


RSUsers.Open "Select * From tbl_Users where Username = '" & txtUsername & "' and Password = '" & txtPassword & "'", Koneksi
If RSUsers.EOF Then
    MsgBox "Username atau Password Salah!"
    txtUsername.SetFocus
    Else
    Unload Me
    frmMenuUtama.Show
    Call Terbuka
End If
End Function

