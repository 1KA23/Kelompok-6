VERSION 5.00
Begin VB.Form frmMenuUtama 
   Caption         =   "Menu Utama Klinik Apaya"
   ClientHeight    =   10335
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   Picture         =   "FormMenuUtama.frx":0000
   ScaleHeight     =   1245
   ScaleMode       =   0  'User
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnFile 
      Caption         =   "File"
      Begin VB.Menu MnRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu MnLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu MnLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu MnKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu MnPendaftaran 
      Caption         =   "Pendaftaran"
      Begin VB.Menu MnDaftar 
         Caption         =   "Daftar"
      End
   End
End
Attribute VB_Name = "frmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call Terkunci
End Sub

Private Sub MnDaftar_Click()
frmPendaftaranKlinik.Show vbModal
End Sub

Private Sub MnRegister_Click()
frmRegister.Show vbModal
End Sub

Private Sub MnKeluar_Click()
End
End Sub

Sub Terkunci()
MnLogin.Enabled = True
MnLogout.Enabled = False
MnPendaftaran.Enabled = False
End Sub

Private Sub MnLogin_Click()
frmLogin.Show vbModal
End Sub

Private Sub MnLogout_Click()
Call Terkunci
End Sub

