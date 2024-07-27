VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPendaftaranKlinik 
   BackColor       =   &H00000080&
   Caption         =   "Form2"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15645
   LinkTopic       =   "Form2"
   ScaleHeight     =   10635
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12720
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10560
      Top             =   6960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   1080
      TabIndex        =   14
      Top             =   7560
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   13
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   12
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   11
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   6120
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   8
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      Caption         =   "No Telepon"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   4
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Caption         =   "Usia"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "frmPendaftaranKlinik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Command1.Caption = "Input" Then
    Command1.Caption = "Simpan"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "Batal"
    Else
    If Text1 = "" Or Combo1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim TambahData
    TambahData = "Insert into tbl_Pendaftaran values ('" & Text1 & "','" & Combo1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
    Koneksi.Execute TambahData
    MsgBox "Tambah Data Berhasil"
    Call KondisiAwal
    Call MunculData
    End If
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Edit" Then
    Command2.Caption = "Simpan"
    Command1.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "Batal"
    Else
    If Text1 = "" Or Combo1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim EditData
    EditData = "Update tbl_Pendaftaran set JenisKelamin = '" & Combo1.Text & "', Usia = '" & Text2 & "', Alamat = '" & Text3 & "', NoTelepon = '" & Text4 & "' where Nama='" & Text1 & "'"
    Koneksi.Execute EditData
    MsgBox "Update Data Berhasil"
    Call KondisiAwal
    Call MunculData
    End If
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Hapus" Then
    Command3.Caption = "Delete"
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Caption = "Batal"
    Else
    If Text1 = "" Or Combo1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "Silahkan isi data terlebih dahulu"
    
    Else
    Call BukaDB
    Dim HapusData As String
    HapusData = "Delete from tbl_Pendaftaran where Nama='" & Text1 & "'"
    Koneksi.Execute HapusData
    MsgBox "Hapus Data Berhasil"
    Call KondisiAwal
    Call MunculData
    End If
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Tutup" Then
    Me.Hide
    Else
    Call KondisiAwal
End If
End Sub

Private Sub Command5_Click()

With CrystalReport1

.ReportFileName = App.Path & "\Report1.rpt"

.Connect = App.Path & "KlinikDB.mdb"

.SQLQuery = "SELECT * FROM tbl_Pendaftaran"

.Destination = crptToWindow

.WindowState = crptMaximized
.Action = 1
End With

End Sub

Private Sub Form_Load()
Call KondisiAwal
Call MunculData
End Sub

Sub KondisiAwal()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Combo1.Clear
Combo1.AddItem "Laki-laki"
Combo1.AddItem "Perempuan"
Command1.Caption = "Input"
Command2.Caption = "Edit"
Command3.Caption = "Hapus"
Command4.Caption = "Tutup"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Sub MunculData()
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "tbl_Pendaftaran"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BukaDB
    RSPendaftaran.Open "Select * From tbl_Pendaftaran where Nama = '" & Text1 & "'", Koneksi
    If Not RSPendaftaran.EOF Then
    Combo1 = RSPendaftaran!JenisKelamin
    Text2 = RSPendaftaran!Usia
    Text3 = RSPendaftaran!Alamat
    Text4 = RSPendaftaran!NoTelepon
    Else
    MsgBox "Data Tidak Ada!"
    End If
End If
End Sub


