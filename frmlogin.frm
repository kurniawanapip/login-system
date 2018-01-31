VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "login dulu"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtpassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtusername 
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Masukan Username Dan Password"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Private Sub cmdlogin_Click()
Call konekdb

If txtusername.Text = "" Then
MsgBox "username masih kosong!", vbCritical, "perhatian"
txtusername.SetFocus
ElseIf txtpassword.Text = "" Then
MsgBox "password masih kosong!", vbCritical, "perhatian"
txtpassword.SetFocus
Else
sql = "select * from login where username='" & txtusername.Text & "' and password='" & txtpassword.Text & "'"
RsAdmin.Open (sql), konek
    If RsAdmin.EOF Then
    MsgBox "username atau password salah !", vbExclamation, "gagal !"
    txtusername.Text = ""
    txtpassword.Text = ""
    txtusername.SetFocus
    Else
    Unload Me
    Form2.Show
    MsgBox "anda berhasil login !", vbInformation, "berhasil !"
    End If
    
End If
End Sub
