VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menampilkan Program Default File"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Tampilkan"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GetAssociatedProgram(ByVal Extension _
As String) As String
    Dim Path As String
    Dim FileName As String
    Dim nRet As Long
    Const MAX_PATH As Long = 260
    'Buat sebuah file temporal
    Path = String$(MAX_PATH, 0)
    If GetTempPath(MAX_PATH, Path) Then
        FileName = String$(MAX_PATH, 0)
        If GetTempFileName(Path, "~", 0, FileName) Then
            FileName = Left$(FileName, _
                InStr(FileName, vbNullChar) - 1)
            'Rename file untuk menambahkan ekstensi
            Name FileName As Left$(FileName, _
                InStr(FileName, ".")) & Extension
                FileName = Left$(FileName, _
                InStr(FileName, ".")) & Extension
            'Ambil assosiasi ekstensi EXE
            Path = String$(MAX_PATH, 0)
            Call FindExecutable(FileName, _
                vbNullString, Path)
            GetAssociatedProgram = Left$( _
                Path, InStr(Path, vbNullChar) - 1)
            'Hapus file temporal
            Kill FileName
        End If
    End If
End Function

Private Sub Command1_Click()
    MsgBox GetAssociatedProgram(Combo1.Text)
End Sub

Private Sub Form_Load()
  With Combo1
    .AddItem "TXT"
    .AddItem "DOC"
    .AddItem "XLS"
    .AddItem "JPG"
    .AddItem "BMP"
    .AddItem "GIF"
    .AddItem "DAT"
    .AddItem "MP3"
    .Text = "TXT"
  End With
End Sub


