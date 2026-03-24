VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmdecode 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ceaser Shift Cipher"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkocult 
      BackColor       =   &H00C0C0C0&
      Caption         =   "oculta/mostra texto"
      Height          =   810
      Left            =   7155
      TabIndex        =   8
      Top             =   2835
      Width           =   855
   End
   Begin VB.CommandButton cmdopen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abrir"
      Height          =   495
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1695
      Width           =   900
   End
   Begin VB.CommandButton cmdsalvar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salvar"
      Height          =   495
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   825
      Width           =   900
   End
   Begin MSComDlg.CommonDialog cdlsave 
      Left            =   7440
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDepth 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   285
      MaxLength       =   3
      TabIndex        =   3
      Top             =   6195
      Width           =   1215
   End
   Begin VB.CommandButton cmddecode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "desencriptar"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6270
      Width           =   1215
   End
   Begin VB.CommandButton cmdencode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Encriptar"
      Height          =   495
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6270
      Width           =   1215
   End
   Begin VB.TextBox txtEnc 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      IMEMode         =   3  'DISABLE
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   6705
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Chave númerica"
      Height          =   195
      Left            =   285
      TabIndex        =   5
      Top             =   5940
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Texto Encode / Decode"
      Height          =   330
      Left            =   2625
      TabIndex        =   4
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "frmdecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkocult_Click()
If chkocult.Value = 1 Then
    txtEnc.PasswordChar = "*"
Else
    txtEnc.PasswordChar = ""
End If
End Sub

Private Sub cmdencode_Click()
'para evitar erros caso a pessoa não digite a chave
On Error GoTo fim
txtEnc.Text = Encode(txtEnc.Text, CInt(txtDepth.Text))
fim:
End Sub


Private Sub cmddecode_Click()
'para evitar erros caso a pessoa não digite a chave
On Error GoTo fim
txtEnc.Text = Decode(txtEnc.Text, CInt(txtDepth.Text))
fim:
End Sub
'instruções para abrir um arquivo .crip
Private Sub cmdopen_Click()
Dim nome As String
cdlsave.Filter = "Files (*.crip)|*.crip"
cdlsave.DefaultExt = "crip"
cdlsave.DialogTitle = "Abrir Arquivo"
cdlsave.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
cdlsave.ShowOpen
On Error GoTo fim
Open cdlsave.FileName For Input As #1
    Input #1, Receb
             txtEnc.Text = Receb
Close 1
fim:
End Sub
'instruções para salvar um arquivo .crip
Private Sub cmdsalvar_Click()
'salva o arquivo
Dim nome As String
cdlsave.Filter = "Files (*.crip)|*.crip"
cdlsave.DefaultExt = "crip"
cdlsave.DialogTitle = "Salvar Arquivo"
cdlsave.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
cdlsave.ShowSave
On Error GoTo fim
Open cdlsave.FileName For Output As #1
    Write #1, txtEnc.Text
Close 1
fim:
End Sub

'bloueio de teclas permitindo apenas números e backspace para evitar erros
Private Sub txtDepth_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Then
    Exit Sub
Else
    KeyAscii = 0
    Beep
End If
End Sub
