VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Edita_perfil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agenda - Editar as informações do seu perfil"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   DrawMode        =   12  'Nop
   Icon            =   "Edita_perfil.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   690
      Index           =   3
      Left            =   90
      TabIndex        =   2
      Top             =   3105
      Width           =   6075
      Begin MSAdodcLib.Adodc tabela_perfil 
         Height          =   330
         Left            =   195
         Top             =   510
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"Edita_perfil.frx":1082
         OLEDBString     =   $"Edita_perfil.frx":1169
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from Colaboradores"
         Caption         =   "Tabela Colaborador"
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
      Begin VB.CommandButton salvaredicaoperfil 
         Caption         =   "Salvar"
         Height          =   360
         Left            =   3495
         TabIndex        =   14
         Top             =   225
         Width           =   2190
      End
      Begin VB.CommandButton cancelaredicaoperfil 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   195
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1785
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   1305
      Width           =   6060
      Begin VB.TextBox confirmanovasenha 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2205
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1215
         Width           =   3480
      End
      Begin VB.TextBox novasenha 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2205
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   735
         Width           =   3480
      End
      Begin VB.TextBox senha_atual_perfil 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2205
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   315
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Redigite a Nova Senha:"
         Height          =   225
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1245
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Nova Senha:"
         Height          =   225
         Index           =   3
         Left            =   990
         TabIndex        =   10
         Top             =   780
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Senha Atual:"
         Height          =   225
         Index           =   2
         Left            =   1035
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6090
      Begin VB.TextBox nome_perfil 
         Height          =   300
         Left            =   885
         TabIndex        =   4
         Top             =   720
         Width           =   4830
      End
      Begin VB.TextBox chapa_perfil 
         Height          =   300
         Left            =   885
         TabIndex        =   3
         Top             =   285
         Width           =   4830
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   225
         Index           =   1
         Left            =   255
         TabIndex        =   6
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Chapa:"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   330
         Width           =   750
      End
   End
End
Attribute VB_Name = "Edita_perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelaredicaoperfil_Click()
Unload Me
End Sub

Private Sub confirmanovasenha_LostFocus()

If novasenha.Text <> "" Then
    If novasenha.Text = confirmanovasenha.Text Then
    Else
    MsgBox "As senhas digitas não são iguais.", vbOKOnly + vbExclamation + vbSystemModal
    novasenha.Text = ""
    confirmanovasenha.Text = ""
    novasenha.SetFocus
    End If
End If

End Sub

Private Sub Form_load()

nome_perfil.MaxLength = 40
novasenha.MaxLength = 20
chapa_perfil.MaxLength = 4
confirmanovasenha.MaxLength = 20
senha_atual_perfil.MaxLength = 20



chapa_perfil.Enabled = False
chapa_perfil.Text = colaborador_chapa


tabela_perfil.RecordSource = "Select NOME, CHAPA, SENHA from COLABORADORES where ID =" & colaborador_id
tabela_perfil.Refresh
nome_perfil.Text = tabela_perfil.Recordset.Fields("NOME").Value


End Sub

Private Sub nome_perfil_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 8 ' backspace
        Case 65 To 90 'A-Z
        Case 97 To 122 'a-z
        Case 32 'blank space
        Case Else
        MsgBox "Apenas letras.", vbOKOnly + vbExclamation + vbSystemModal
            KeyAscii = 0
    End Select

End Sub
Private Sub salvaredicaoperfil_Click()

If nome_perfil.Text = "" Then
    MessageBox.Show ("Preencha o campo nome.")
    nome_perfil.SetFocus
    nome_perfil.Text = ""
    nome_perfil.BackColor = &HFF00&
    Exit Sub
End If

If tabela_perfil.Recordset.Fields("SENHA").Value <> senha_atual_perfil.Text Then
    MsgBox "A senha atual não está correta.", vbOKOnly + vbExclamation + vbSystemModal
    senha_atual_perfil.Text = ""
    senha_atual_perfil.SetFocus
    senha_atual_perfil.BackColor = &HFF00&
    Exit Sub
End If

If novasenha.Text = senha_atual_perfil.Text Then
    MsgBox "Escolha uma senha diferente da usada anteriormente.", vbOKOnly + vbExclamation + vbSystemModal
    novasenha.Text = ""
    confirmanovasenha.Text = ""
    novasenha.SetFocus
Exit Sub
End If

If novasenha.Text <> confirmanovasenha.Text Then
MsgBox "As senhas digitas não são iguais.", vbOKOnly + vbExclamation + vbSystemModal
novasenha.Text = ""
confirmanovasenha.Text = ""
novasenha.SetFocus
Exit Sub
End If

tabela_perfil.Recordset.Fields("NOME").Value = nome_perfil.Text
tabela_perfil.Recordset.Fields("SENHA").Value = novasenha.Text
tabela_perfil.Recordset.Update
tabela_perfil.Refresh

Unload Me
End Sub
Private Sub senha_atual_perfil_LostFocus()

If senha_atual_perfil.Text = "" Then
Else
        If tabela_perfil.Recordset.Fields("SENHA").Value = senha_atual_perfil.Text Then
        novasenha.SetFocus
        Else
        MsgBox "A senha atual não está correta.", vbOKOnly + vbExclamation + vbSystemModal
        senha_atual_perfil.Text = ""
        senha_atual_perfil.SetFocus
        End If
End If
End Sub
