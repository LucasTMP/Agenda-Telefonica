VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda - CIP"
   ClientHeight    =   6135
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "                            Efetuar Login                          "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   630
      TabIndex        =   0
      Top             =   2985
      Width           =   3660
      Begin VB.TextBox campochapa 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   405
         MaxLength       =   4
         TabIndex        =   4
         Top             =   390
         Width           =   2865
      End
      Begin VB.TextBox camposenha 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   420
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1215
         Width           =   2865
      End
      Begin VB.CommandButton Cmd_Sair 
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   2160
         Picture         =   "login.frx":6C35
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Sair"
         Top             =   1755
         Width           =   1125
      End
      Begin VB.CommandButton Cmd_Logar 
         Caption         =   "Logar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   405
         Picture         =   "login.frx":731F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Efetuar Login"
         Top             =   1755
         Width           =   1155
      End
      Begin VB.Line Line3 
         X1              =   3555
         X2              =   3555
         Y1              =   135
         Y2              =   2640
      End
      Begin VB.Line Line2 
         X1              =   3555
         X2              =   105
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   105
         Y1              =   120
         Y2              =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Chapa :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   375
         TabIndex        =   6
         Top             =   75
         Width           =   2790
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Senha :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   420
         TabIndex        =   5
         Top             =   915
         Width           =   2790
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   5820
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Versão: 1.0"
            TextSave        =   "Versão: 1.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6033
            Text            =   "        Atualizado em: 30/07/2019"
            TextSave        =   "        Atualizado em: 30/07/2019"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   6900
      Left            =   -90
      Picture         =   "login.frx":83A1
      ScaleHeight     =   456
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   518
      TabIndex        =   8
      Top             =   -915
      Width           =   7830
      Begin MSAdodcLib.Adodc dblogin 
         Height          =   330
         Left            =   675
         Top             =   3510
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
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
         Connect         =   $"login.frx":DB6A
         OLEDBString     =   $"login.frx":DC51
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from colaboradores"
         Caption         =   "tabela_colaboradores"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1290
         Left            =   270
         Picture         =   "login.frx":DD38
         ScaleHeight     =   86
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   297
         TabIndex        =   13
         Top             =   1155
         Width           =   4455
      End
      Begin VB.Label descri 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agenda Telefonica - CIP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   405
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3945
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agenda Telefonica - CIP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   405
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3945
      End
      Begin VB.Label descri 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CIP - Agenda Telefonica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   405
         Index           =   0
         Left            =   465
         TabIndex        =   9
         Top             =   2865
         Width           =   3975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Agenda Telefonica - CIP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3945
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    dblogin.RecordSource = "select ID, NOME, CHAPA, SENHA from COLABORADORES where CHAPA ='" & campochapa.Text & "' and SENHA ='" & camposenha.Text & "'"
    dblogin.Refresh

    If dblogin.Recordset.RecordCount > 0 Then
    colaborador_senha = campochapa.Text
    colaborador_chapa = campochapa.Text
    colaborador_nome = dblogin.Recordset.Fields("NOME").Value
    colaborador_id = dblogin.Recordset.Fields("ID").Value
    Unload Me
    Form1.Show
    Else
    MsgBox "Usuário e Senha não conferem, contate o administrador do sistema.", vbCritical + vbOKOnly + vbApplicationModal, "AVISO"
    End If
    
End If
End Sub


Private Sub Cmd_Logar_Click(Index As Integer)

dblogin.RecordSource = "select ID, NOME, CHAPA, SENHA, ADM from COLABORADORES where CHAPA ='" & campochapa.Text & "' and SENHA ='" & camposenha.Text & "'"
dblogin.Refresh

If dblogin.Recordset.RecordCount > 0 Then
colaborador_senha = campochapa.Text
colaborador_chapa = campochapa.Text
colaborador_nome = dblogin.Recordset.Fields("NOME").Value
colaborador_id = dblogin.Recordset.Fields("ID").Value
colaborador_adm = dblogin.Recordset.Fields("ADM").Value
Unload Me
Form1.Show
Else
MsgBox "Usuário e Senha não conferem, contate o administrador do sistema.", vbCritical + vbOKOnly + vbApplicationModal, "AVISO"
End If

End Sub

Private Sub Cmd_Sair_Click(Index As Integer)
Unload Me
End Sub


