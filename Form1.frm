VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Nama 
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   477
   ScaleMode       =   2  'Point
   ScaleWidth      =   710.25
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtpTglLahir 
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Top             =   3480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Format          =   189595649
      CurrentDate     =   46205
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   10320
      TabIndex        =   21
      Top             =   1080
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid dgMahasiswa 
      Height          =   5655
      Left            =   6840
      TabIndex        =   20
      Top             =   2280
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9975
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
   Begin VB.CommandButton Command5 
      Caption         =   "EXPORT"
      Height          =   495
      Left            =   12600
      TabIndex        =   19
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RESET"
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   8640
      Width           =   1215
   End
   Begin VB.ComboBox cmbJenjang 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   7560
      Width           =   2775
   End
   Begin VB.ComboBox cmbFakultas 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   6840
      Width           =   2775
   End
   Begin VB.ComboBox cmdJurusan 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox txtAlamat 
      Height          =   1335
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtNIM 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtUmur 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "FORM DATA MAHASISWA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   25
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "LIST DATA MAHASISWA"
      Height          =   255
      Left            =   6960
      TabIndex        =   23
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label search 
      Caption         =   "Search :"
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Jenjang   :"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Fakultas  :"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Jurusan   :"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Alamat    :"
      Height          =   1215
      Left            =   1320
      TabIndex        =   11
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label tgllahir 
      Caption         =   "Tgl Lahir :"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label NIM 
      Caption         =   "NIM        :"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Umur 
      Caption         =   "Umur      :"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Nama 
      Caption         =   "Nama     :"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Nama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelectedID As Long

Private Sub Form_Load()
    Call BukaKoneksi: dgMahasiswa.MarqueeStyle = 3: Call LoadJurusan: Call RefreshData
End Sub

Private Sub RefreshData()
    Dim q As String: q = "SELECT m.id, m.nama, m.umur, m.nim, m.tgl_lahir, m.alamat, j.nama_jurusan, j.fakultas, j.jenjang FROM Mahasiswa m LEFT JOIN Jurusan j ON m.id_jurusan = j.id_jurusan "
    If txtCari.Text <> "" Then q = q & "WHERE m.nama ILIKE '%" & txtCari.Text & "%' OR m.nim LIKE '%" & txtCari.Text & "%' "
    q = q & "ORDER BY m.id DESC"
    Set rs = New ADODB.Recordset: rs.CursorLocation = adUseClient: rs.Open q, conn, 3, 1: Set dgMahasiswa.DataSource = rs
    If dgMahasiswa.Columns.Count > 0 Then dgMahasiswa.Columns(0).Visible = False
End Sub

Private Sub LoadJurusan()
    Dim rsJ As New ADODB.Recordset: cmdJurusan.Clear
    rsJ.Open "SELECT id_jurusan, nama_jurusan FROM Jurusan ORDER BY nama_jurusan", conn, 3, 1
    Do While Not rsJ.EOF: cmdJurusan.AddItem rsJ!nama_jurusan: cmdJurusan.ItemData(cmdJurusan.NewIndex) = rsJ!id_jurusan: rsJ.MoveNext: Loop
End Sub

Private Sub cmdJurusan_Click()
    If cmdJurusan.ListIndex < 0 Then Exit Sub
    Dim rsJ As New ADODB.Recordset: rsJ.Open "SELECT fakultas, jenjang FROM Jurusan WHERE id_jurusan = " & cmdJurusan.ItemData(cmdJurusan.ListIndex), conn, 3, 1
    If Not rsJ.EOF Then cmbFakultas.Text = rsJ!fakultas: cmbJenjang.Text = rsJ!jenjang
End Sub

Private Sub dtpTglLahir_Change(): txtUmur.Text = DateDiff("yyyy", dtpTglLahir.Value, Date): End Sub

Private Sub txtCari_Change(): Call RefreshData: End Sub

Private Sub dgMahasiswa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If Not rs.EOF And Not rs.BOF Then
        SelectedID = rs!id: txtNama.Text = rs!Nama: txtUmur.Text = rs!Umur: txtNIM.Text = rs!NIM
        dtpTglLahir.Value = rs!tgl_lahir: txtAlamat.Text = rs!alamat: cmdJurusan.Text = rs!nama_jurusan
    End If
End Sub

Private Sub Command3_Click() ' SAVE
    If cmdJurusan.ListIndex = -1 Then Exit Sub
    conn.Execute "INSERT INTO Mahasiswa (nama, umur, nim, tgl_lahir, alamat, id_jurusan) VALUES ('" & txtNama.Text & "', " & Val(txtUmur.Text) & ", '" & txtNIM.Text & "', '" & Format(dtpTglLahir.Value, "YYYY-MM-DD") & "', '" & txtAlamat.Text & "', " & cmdJurusan.ItemData(cmdJurusan.ListIndex) & ")"
    Call Command1_Click: Call RefreshData
End Sub

Private Sub Command2_Click() ' UPDATE
    If SelectedID = 0 Or cmdJurusan.ListIndex = -1 Then Exit Sub
    conn.Execute "UPDATE Mahasiswa SET nama='" & txtNama.Text & "', umur=" & Val(txtUmur.Text) & ", nim='" & txtNIM.Text & "', tgl_lahir='" & Format(dtpTglLahir.Value, "YYYY-MM-DD") & "', alamat='" & txtAlamat.Text & "', id_jurusan=" & cmdJurusan.ItemData(cmdJurusan.ListIndex) & " WHERE id=" & SelectedID
    Call Command1_Click: Call RefreshData
End Sub

Private Sub Command4_Click() ' DELETE
    If SelectedID = 0 Then Exit Sub
    conn.Execute "DELETE FROM Mahasiswa WHERE id=" & SelectedID
    Call Command1_Click: Call RefreshData
End Sub

Private Sub Command1_Click() ' RESET
    SelectedID = 0: txtNama.Text = "": txtUmur.Text = "": txtNIM.Text = "": dtpTglLahir.Value = Date
    txtAlamat.Text = "": cmdJurusan.ListIndex = -1: cmbFakultas.Text = "": cmbJenjang.Text = ""
End Sub

Private Sub Command5_Click() ' EXPORT
    On Error Resume Next: Dim fNum As Integer, i As Integer, d As String: fNum = FreeFile
    Open App.Path & "\export_mahasiswa.xls" For Output As #fNum
    For i = 1 To dgMahasiswa.Columns.Count - 1: d = d & dgMahasiswa.Columns(i).Caption & vbTab: Next: Print #fNum, d
    If Not rs.BOF Then rs.MoveFirst
    Do While Not rs.EOF: d = "": For i = 1 To dgMahasiswa.Columns.Count - 1: d = d & rs.Fields(i).Value & vbTab: Next: Print #fNum, d: rs.MoveNext: Loop
    Close #fNum: MsgBox "Export sukses!", vbInformation: Call RefreshData
End Sub
