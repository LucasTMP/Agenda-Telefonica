VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda - CIP"
   ClientHeight    =   11370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13800
   ControlBox      =   0   'False
   FillColor       =   &H80000001&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Reference Specialty"
      Size            =   9.75
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "agenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "agenda.frx":10CA
   ScaleHeight     =   11370
   ScaleWidth      =   13800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame dados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   10320
      TabIndex        =   45
      Top             =   1350
      Visible         =   0   'False
      Width           =   3345
      Begin VB.Line Line1 
         Index           =   2
         X1              =   195
         X2              =   3165
         Y1              =   3465
         Y2              =   3465
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   3150
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Label texto_dados2 
         BackStyle       =   0  'Transparent
         Caption         =   "tempo até abrir, numero colaboradores sonlorei ipsonlorei ipson"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   330
         TabIndex        =   47
         Top             =   2205
         Width           =   2715
      End
      Begin VB.Label texto_dados 
         BackStyle       =   0  'Transparent
         Caption         =   "lorei ipsonlorei ipsonlorei ipsonlorei ipsonlorei ipsonlorei ipsonlorei ipson"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   330
         TabIndex        =   46
         Top             =   555
         Width           =   2715
      End
   End
   Begin VB.CommandButton relatorio 
      Caption         =   "Dados Gerais"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10365
      Picture         =   "agenda.frx":B82C7
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   15
      Width           =   1690
   End
   Begin VB.CommandButton atualizar 
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   8640
      Picture         =   "agenda.frx":B89B1
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   15
      Width           =   1690
   End
   Begin VB.CommandButton perfil 
      Caption         =   "Editar Perfil"
      DragIcon        =   "agenda.frx":B909B
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   15
      Picture         =   "agenda.frx":C196D
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   15
      Width           =   1690
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisar"
      DragIcon        =   "agenda.frx":C29EF
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6915
      Picture         =   "agenda.frx":CB2C1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   1690
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10845
      Top             =   5085
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
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
      Connect         =   $"agenda.frx":CC343
      OLEDBString     =   $"agenda.frx":CC42A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from agenda"
      Caption         =   "Agenda ALL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btn_salva_alt 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10320
      Picture         =   "agenda.frx":CC511
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Alterar Número"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1690
   End
   Begin VB.Frame Framealt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alterar Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4080
      Left            =   10305
      TabIndex        =   27
      Top             =   6015
      Visible         =   0   'False
      Width           =   3385
      Begin VB.TextBox txt_telaltalt 
         DataField       =   "NUMEROS_ALTERNATIVOS"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   180
         MaxLength       =   35
         TabIndex        =   36
         Top             =   3150
         Width           =   2985
      End
      Begin VB.TextBox txt_nomealt 
         DataField       =   "NOME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         MaxLength       =   50
         TabIndex        =   29
         Top             =   915
         Width           =   2985
      End
      Begin VB.TextBox txt_telalt 
         DataField       =   "NUMERO"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         MaxLength       =   25
         TabIndex        =   28
         Top             =   2055
         Width           =   2985
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefone Alternativo:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   180
         TabIndex        =   37
         Top             =   2805
         Width           =   2805
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefone :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   31
         Top             =   1725
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   30
         Top             =   555
         Width           =   915
      End
   End
   Begin VB.Frame Framedel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Deletar Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4080
      Left            =   10305
      TabIndex        =   18
      Top             =   6000
      Visible         =   0   'False
      Width           =   3385
      Begin VB.TextBox txt_telaltdel 
         DataField       =   "NUMEROS_ALTERNATIVOS"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         MaxLength       =   35
         TabIndex        =   39
         Top             =   3150
         Width           =   2985
      End
      Begin VB.TextBox txt_teldel 
         DataField       =   "NUMERO"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         MaxLength       =   25
         TabIndex        =   20
         Top             =   2055
         Width           =   2985
      End
      Begin VB.TextBox txt_nomedel 
         DataField       =   "NOME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         MaxLength       =   50
         TabIndex        =   19
         Top             =   915
         Width           =   2985
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefone Alternativo:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   40
         Top             =   2805
         Width           =   2595
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   22
         Top             =   555
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefone :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   21
         Top             =   1725
         Width           =   1050
      End
   End
   Begin VB.CommandButton btn_cancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   12015
      Picture         =   "agenda.frx":CCBFB
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Fechar Comandos"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1690
   End
   Begin VB.Frame Frameadd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adicionar Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4080
      Left            =   10305
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   3385
      Begin VB.TextBox pessoal 
         DataField       =   "ID_COLABORADOR"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1005
         TabIndex        =   42
         Text            =   "Pessoal"
         Top             =   225
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox interno 
         DataField       =   "INTERNO"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         TabIndex        =   38
         Text            =   "Interno S/N"
         Top             =   225
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txt_telaltadd 
         DataField       =   "NUMEROS_ALTERNATIVOS"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   210
         MaxLength       =   35
         TabIndex        =   34
         Top             =   3150
         Width           =   2985
      End
      Begin VB.TextBox txt_nomeadd 
         DataField       =   "NOME"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         MaxLength       =   50
         TabIndex        =   17
         Top             =   915
         Width           =   2985
      End
      Begin VB.TextBox txt_teladd 
         DataField       =   "NUMERO"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   180
         MaxLength       =   25
         TabIndex        =   16
         Top             =   2055
         Width           =   2985
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefone Alternativo:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Top             =   2805
         Width           =   2805
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefone :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1725
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   555
         Width           =   915
      End
   End
   Begin VB.CommandButton btn_alterar 
      Caption         =   "Alterar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3465
      Picture         =   "agenda.frx":CDC7D
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Alterar Número"
      Top             =   15
      Width           =   1690
   End
   Begin VB.CommandButton btn_deletar 
      Caption         =   "Deletar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5190
      Picture         =   "agenda.frx":CDF87
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Deletar Número"
      Top             =   15
      Width           =   1690
   End
   Begin VB.CommandButton btn_add 
      Caption         =   "Adicionar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1740
      Picture         =   "agenda.frx":CE291
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Adicionar Número"
      Top             =   15
      Width           =   1690
   End
   Begin MSDataGridLib.DataGrid tel 
      Bindings        =   "agenda.frx":CF313
      Height          =   9765
      Left            =   0
      TabIndex        =   6
      Top             =   1305
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   17224
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      DefColWidth     =   206
      ColumnHeaders   =   -1  'True
      ForeColor       =   -2147483625
      HeadLines       =   2
      RowHeight       =   18
      TabAction       =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
            LCID            =   1046
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
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btn_salva_deletar 
      Caption         =   "Deletar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10305
      Picture         =   "agenda.frx":CF328
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Deletar Número"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1690
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
      Height          =   885
      Left            =   12090
      Picture         =   "agenda.frx":CF632
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sair do programa"
      Top             =   15
      Width           =   1690
   End
   Begin VB.CommandButton btn_salva_add 
      Caption         =   "Adicionar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10305
      Picture         =   "agenda.frx":CFD1C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Adicionar Número"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1690
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   11055
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   13653
            Picture         =   "agenda.frx":D0406
            Text            =   "CIP - Companhia Industrial de Peças"
            TextSave        =   "CIP - Companhia Industrial de Peças"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "14:31"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "29/08/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   5477
            Text            =   "Desenvolvimento: Lucas e Vagner."
            TextSave        =   "Desenvolvimento: Lucas e Vagner."
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
   Begin TabDlg.SSTab SSTab10 
      DragIcon        =   "agenda.frx":D7CD0
      Height          =   10125
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   945
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   17859
      _Version        =   393216
      MousePointer    =   1
      Tab             =   1
      TabHeight       =   626
      BackColor       =   -2147483647
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Tel. Internos"
      TabPicture(0)   =   "agenda.frx":D8D9A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TabStrip2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Tel. Externos"
      TabPicture(1)   =   "agenda.frx":D8DB6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TabStrip3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Tel. Pessoais"
      TabPicture(2)   =   "agenda.frx":D8DD2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TabStrip4"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab1 
         Height          =   10200
         Index           =   0
         Left            =   -75000
         TabIndex        =   3
         Top             =   360
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   17992
         _Version        =   393216
         TabOrientation  =   3
         Tabs            =   26
         TabsPerRow      =   26
         TabHeight       =   617
         TabMaxWidth     =   697
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&A"
         TabPicture(0)   =   "agenda.frx":D8DEE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Command2(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Command3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "&B"
         TabPicture(1)   =   "agenda.frx":D8E0A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "&C"
         TabPicture(2)   =   "agenda.frx":D8E26
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "&D"
         TabPicture(3)   =   "agenda.frx":D8E42
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "&E"
         TabPicture(4)   =   "agenda.frx":D8E5E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         TabCaption(5)   =   "&F"
         TabPicture(5)   =   "agenda.frx":D8E7A
         Tab(5).ControlEnabled=   0   'False
         Tab(5).ControlCount=   0
         TabCaption(6)   =   "&G"
         TabPicture(6)   =   "agenda.frx":D8E96
         Tab(6).ControlEnabled=   0   'False
         Tab(6).ControlCount=   0
         TabCaption(7)   =   "&H"
         TabPicture(7)   =   "agenda.frx":D8EB2
         Tab(7).ControlEnabled=   0   'False
         Tab(7).ControlCount=   0
         TabCaption(8)   =   "&I"
         TabPicture(8)   =   "agenda.frx":D8ECE
         Tab(8).ControlEnabled=   0   'False
         Tab(8).ControlCount=   0
         TabCaption(9)   =   "&J"
         TabPicture(9)   =   "agenda.frx":D8EEA
         Tab(9).ControlEnabled=   0   'False
         Tab(9).ControlCount=   0
         TabCaption(10)  =   "&K"
         TabPicture(10)  =   "agenda.frx":D8F06
         Tab(10).ControlEnabled=   0   'False
         Tab(10).ControlCount=   0
         TabCaption(11)  =   "&L"
         TabPicture(11)  =   "agenda.frx":D8F22
         Tab(11).ControlEnabled=   0   'False
         Tab(11).ControlCount=   0
         TabCaption(12)  =   "&M"
         TabPicture(12)  =   "agenda.frx":D8F3E
         Tab(12).ControlEnabled=   0   'False
         Tab(12).ControlCount=   0
         TabCaption(13)  =   "&N"
         TabPicture(13)  =   "agenda.frx":D8F5A
         Tab(13).ControlEnabled=   0   'False
         Tab(13).ControlCount=   0
         TabCaption(14)  =   "&O"
         TabPicture(14)  =   "agenda.frx":D8F76
         Tab(14).ControlEnabled=   0   'False
         Tab(14).ControlCount=   0
         TabCaption(15)  =   "&P"
         TabPicture(15)  =   "agenda.frx":D8F92
         Tab(15).ControlEnabled=   0   'False
         Tab(15).ControlCount=   0
         TabCaption(16)  =   "&Q"
         TabPicture(16)  =   "agenda.frx":D8FAE
         Tab(16).ControlEnabled=   0   'False
         Tab(16).ControlCount=   0
         TabCaption(17)  =   "&R"
         TabPicture(17)  =   "agenda.frx":D8FCA
         Tab(17).ControlEnabled=   0   'False
         Tab(17).ControlCount=   0
         TabCaption(18)  =   "&S"
         TabPicture(18)  =   "agenda.frx":D8FE6
         Tab(18).ControlEnabled=   0   'False
         Tab(18).ControlCount=   0
         TabCaption(19)  =   "&T"
         TabPicture(19)  =   "agenda.frx":D9002
         Tab(19).ControlEnabled=   0   'False
         Tab(19).ControlCount=   0
         TabCaption(20)  =   "&U"
         TabPicture(20)  =   "agenda.frx":D901E
         Tab(20).ControlEnabled=   0   'False
         Tab(20).ControlCount=   0
         TabCaption(21)  =   "&V"
         TabPicture(21)  =   "agenda.frx":D903A
         Tab(21).ControlEnabled=   0   'False
         Tab(21).ControlCount=   0
         TabCaption(22)  =   "&W"
         TabPicture(22)  =   "agenda.frx":D9056
         Tab(22).ControlEnabled=   0   'False
         Tab(22).ControlCount=   0
         TabCaption(23)  =   "&X"
         TabPicture(23)  =   "agenda.frx":D9072
         Tab(23).ControlEnabled=   0   'False
         Tab(23).ControlCount=   0
         TabCaption(24)  =   "&Y"
         TabPicture(24)  =   "agenda.frx":D908E
         Tab(24).ControlEnabled=   0   'False
         Tab(24).ControlCount=   0
         TabCaption(25)  =   "&Z"
         TabPicture(25)  =   "agenda.frx":D90AA
         Tab(25).ControlEnabled=   0   'False
         Tab(25).ControlCount=   0
         Begin VB.CommandButton Command3 
            Caption         =   "Command2"
            Height          =   615
            Left            =   7560
            TabIndex        =   5
            Top             =   8160
            Width           =   2415
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   615
            Index           =   2
            Left            =   8280
            TabIndex        =   4
            Top             =   9960
            Width           =   2415
         End
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   10155
         Left            =   -65100
         TabIndex        =   7
         Top             =   360
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   17912
         MultiRow        =   -1  'True
         ShowTips        =   0   'False
         Placement       =   3
         TabMinWidth     =   654
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   26
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "A"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "B"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "C"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "D"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "E"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "F"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "G"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "H"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "I"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "J"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "K"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "L"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "M"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "N"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "O"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "P"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Q"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab18 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "R"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab19 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "S"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab20 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "T"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab21 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "U"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab22 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "V"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab23 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "W"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab24 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "X"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab25 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Y"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab26 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Z"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabStrip3 
         Height          =   10155
         Left            =   9885
         TabIndex        =   8
         Top             =   360
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   17912
         MultiRow        =   -1  'True
         ShowTips        =   0   'False
         Placement       =   3
         TabMinWidth     =   673
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   26
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "A"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "B"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "C"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "D"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "E"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "F"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "G"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "H"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "I"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "J"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "K"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "L"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "M"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "N"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "O"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "P"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Q"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab18 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "R"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab19 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "S"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab20 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "T"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab21 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "U"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab22 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "V"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab23 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "W"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab24 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "X"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab25 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Y"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab26 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Z"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabStrip4 
         Height          =   10155
         Left            =   -65100
         TabIndex        =   9
         Top             =   360
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   17912
         MultiRow        =   -1  'True
         ShowTips        =   0   'False
         Placement       =   3
         TabMinWidth     =   673
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   26
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "A"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "B"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "C"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "D"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "E"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "F"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "G"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "H"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "I"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "J"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "K"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "L"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "M"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "N"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "O"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "P"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Q"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab18 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "R"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab19 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "S"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab20 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "T"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab21 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "U"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab22 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "V"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab23 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "W"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab24 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "X"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab25 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Y"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab26 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Z"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10830
      Top             =   5535
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
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
      Connect         =   $"agenda.frx":D90C6
      OLEDBString     =   $"agenda.frx":D91AD
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda where INTERNO = -1 ORDER BY NOME"
      Caption         =   "Agenda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nome :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   9120
      TabIndex        =   33
      Top             =   6840
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Bentrar_Click()

tel.Refresh

End Sub

Private Sub atualizar_Click()

Adodc1.Refresh
Adodc2.Refresh
tel.Refresh
StatusBar1.Refresh

End Sub

Private Sub btn_add_Click()
Adodc1.Refresh
Adodc2.Refresh
tel.Refresh
Adodc2.Recordset.AddNew

'Travar botoes e ações -------------------------------------------------

btn_add.Enabled = False
btn_alterar.Enabled = True
btn_deletar.Enabled = True
btn_salva_add.Visible = True
btn_salva_alt.Visible = False
btn_salva_deletar.Visible = False
btn_cancelar.Visible = True

Framealt.Visible = False
Frameadd.Visible = True
Framedel.Visible = False


'-----------------------------------------------------------------------

 If SSTab10.Item(0).Tab = 0 Then
    interno.Text = "-1"
    End If
    
    If SSTab10.Item(0).Tab = 1 Then
    interno.Text = "0"
    End If
    
    If SSTab10.Item(0).Tab = 2 Then
    interno.Text = "0"
End If
    
    
If SSTab10.Item(0).Tab = 0 Then
    pessoal.Text = "0"
    End If
    
    If SSTab10.Item(0).Tab = 1 Then
    pessoal.Text = "0"
    End If
    
    If SSTab10.Item(0).Tab = 2 Then
    pessoal.Text = colaborador_id
End If

End Sub

Private Sub btn_alterar_Click()
Adodc1.Refresh
Adodc2.Refresh
tel.Refresh
'Travar botoes e ações -------------------------------------------------

btn_add.Enabled = True
btn_alterar.Enabled = False
btn_deletar.Enabled = True
btn_salva_add.Visible = False
btn_salva_alt.Visible = True
btn_salva_deletar.Visible = False
btn_cancelar.Visible = True

Framealt.Visible = True
Frameadd.Visible = False
Framedel.Visible = False

'-----------------------------------------------------------------------


End Sub

Private Sub btn_cancelar_Click()
Adodc1.Refresh
Adodc2.Refresh
tel.Refresh
'Travar botoes e ações -------------------------------------------------

btn_add.Enabled = True
btn_alterar.Enabled = True
btn_deletar.Enabled = True
btn_salva_add.Visible = False
btn_salva_alt.Visible = False
btn_salva_deletar.Visible = False
btn_cancelar.Visible = False

Framealt.Visible = False
Frameadd.Visible = False
Framedel.Visible = False



'-----------------------------------------------------------------------
End Sub

Private Sub btn_deletar_Click()
Adodc1.Refresh
Adodc2.Refresh
tel.Refresh
'Travar botoes e ações -------------------------------------------------

btn_add.Enabled = True
btn_alterar.Enabled = True
btn_deletar.Enabled = False
btn_salva_add.Visible = False
btn_salva_alt.Visible = False
btn_salva_deletar.Visible = True
btn_cancelar.Visible = True

Framealt.Visible = False
Frameadd.Visible = False
Framedel.Visible = True


'-----------------------------------------------------------------------

End Sub

Private Sub btn_salva_add_Click()

If txt_nomeadd.Text = "" Then
MsgBox "Digite um nome para o novo contato!", vbApplicationModal + vbInformation
txt_nomeadd.SetFocus
Exit Sub
End If

If txt_teladd.Item(0).Text = "" Then
MsgBox "Digite um número para o novo contato!", vbApplicationModal + vbInformation
txt_teladd.Item(0).SetFocus
Exit Sub
End If


Adodc2.Recordset.Update
Adodc2.Refresh
Adodc1.Refresh
tel.Refresh


'Travar botoes e ações -------------------------------------------------

btn_add.Enabled = True
btn_alterar.Enabled = True
btn_deletar.Enabled = True
btn_salva_add.Visible = False
btn_salva_alt.Visible = False
btn_salva_deletar.Visible = False
btn_cancelar.Visible = False

Framealt.Visible = False
Frameadd.Visible = False
Framedel.Visible = False


'-----------------------------------------------------------------------

End Sub

Private Sub btn_salva_alt_Click()
Dim vresult As String

If txt_nomealt.Text = "" Then
MsgBox "Digite um nome para o contato!", vbApplicationModal + vbInformation
txt_nomealt.SetFocus
Exit Sub
End If

If txt_telalt.Text = "" Then
MsgBox "Digite um número para o contato!", vbApplicationModal + vbInformation
txt_telalt.SetFocus
Exit Sub
End If

vresult = MsgBox("Tem certeza que deseja alterar o número ?", vbYesNo + vbQuestion, "Confirmação de Comando!")

If vresult = vbYes Then
Adodc1.Recordset.Update
Adodc1.Refresh
tel.Refresh

'Travar botoes e ações -------------------------------------------------

btn_add.Enabled = True
btn_alterar.Enabled = True
btn_deletar.Enabled = True
btn_salva_add.Visible = False
btn_salva_alt.Visible = False
btn_salva_deletar.Visible = False
btn_cancelar.Visible = False

Framealt.Visible = False
Frameadd.Visible = False
Framedel.Visible = False


'-----------------------------------------------------------------------
MsgBox "Número alterado com êxito.", vbApplicationModal + vbInformation
Else
MsgBox "Alteração cancelada!", vbApplicationModal + vbInformation
End If

End Sub

Private Sub btn_salva_deletar_Click()
Dim vresult As String

vresult = MsgBox("Tem certeza que deseja deletar o número do" & txt_nomedel.Text & "?", vbYesNo + vbQuestion, "Confirmação de Comando!")

If vresult = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Refresh
tel.Refresh

'Travar botoes e ações -------------------------------------------------

btn_add.Enabled = True
btn_alterar.Enabled = True
btn_deletar.Enabled = True
btn_salva_add.Visible = False
btn_salva_alt.Visible = False
btn_salva_deletar.Visible = False
btn_cancelar.Visible = False

Framealt.Visible = False
Frameadd.Visible = False
Framedel.Visible = False


'-----------------------------------------------------------------------

MsgBox "Número deletado com êxito.", vbApplicationModal + vbInformation
Else
MsgBox "Ação cancelada!", vbApplicationModal + vbInformation
End If
End Sub

Private Sub Cmd_Sair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Encontrar.Show

End Sub

Private Sub Form_load()

On Error GoTo errConexao

tel.HeadFont.Bold = True
tel.HeadFont.Size = 10
tel.Font.Size = 11

txt_nomedel.Enabled = False
txt_teldel.Enabled = False
txt_telaltdel.Enabled = False


If colaborador_adm = 1 Then
    
    relatorio.Enabled = True
    btn_alterar.Enabled = True
    btn_deletar.Enabled = True

        Else

        btn_alterar.Enabled = False
        btn_deletar.Enabled = False
        relatorio.Enabled = False

End If



StatusBar1.Panels(1).Text = "Sistema logado como: " & colaborador_nome
   
    If SSTab10.Item(0).Tab = 0 Then
    
        If colaborador_adm = 1 Then

        btn_alterar.Enabled = True
        btn_deletar.Enabled = True

         Else

         btn_alterar.Enabled = False
         btn_deletar.Enabled = False

         End If

    
    Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda where INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
    Adodc1.Refresh
    tel.Refresh
    End If
    
    
    If SSTab10.Item(0).Tab = 1 Then
    Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda where INTERNO = 0 and ID_COLABORADOR = 0 ORDER BY NOME"
    Adodc1.Refresh
    tel.Refresh
    
        If colaborador_adm = 1 Then

        btn_alterar.Enabled = True
        btn_deletar.Enabled = True

            Else

            btn_alterar.Enabled = False
            btn_deletar.Enabled = False

        End If

    End If
    
    If SSTab10.Item(0).Tab = 2 Then
    
        btn_alterar.Enabled = True
        btn_deletar.Enabled = True
    
    Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda where ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
    Adodc1.Refresh
    tel.Refresh
    End If
   
   
errConexao:
   With Err
           If .Number <> 0 Then
              MsgBox "Houve um erro na conexão com o banco de dados, o sistema será encerrado.", vbCritical + vbOKOnly + vbApplicationModal, "AGENDA AVISO"
           End If
   End With
   
End Sub


Private Sub Label8_Click()

End Sub

Private Sub perfil_Click()
Edita_perfil.Show
End Sub

Private Sub relatorio_Click()

If dados.Visible = False Then
dados.Visible = True
Else
dados.Visible = False
End If

Adodc2.RecordSource = "Select * from Colaboradores"
Adodc2.Refresh
total_colaboradores = Adodc2.Recordset.RecordCount

Adodc2.RecordSource = "Select * from Colaboradores where ADM = 1"
Adodc2.Refresh
total_colaboradores_admin = Adodc2.Recordset.RecordCount

Adodc2.RecordSource = "Select * from agenda where ID_COLABORADOR <> 0"
Adodc2.Refresh
total_numeros_pessoais = Adodc2.Recordset.RecordCount

Adodc2.RecordSource = "Select * from agenda"
Adodc2.Refresh
total_numeros = Adodc2.Recordset.RecordCount



texto_dados.Caption = "A agenda possui atualmente " & total_numeros & " registros em sua base de dados, sendo " & total_numeros_pessoais & " números pessoais cadastrados."
texto_dados.Refresh

texto_dados2.Caption = "O sistema conta com " & total_colaboradores & " colaboradores cadastrados, sendo " & total_colaboradores_admin & " usuários administradores."
texto_dados2.Refresh

End Sub

Private Sub SSTab10_Click(intIndex As Integer, intPreviousTab As Integer)

'Travar botoes e ações -------------------------------------------------

Framealt.Visible = False
Frameadd.Visible = False
Framedel.Visible = False

btn_add.Enabled = True
btn_alterar.Enabled = True
btn_deletar.Enabled = True
btn_salva_add.Visible = False
btn_salva_alt.Visible = False
btn_salva_deletar.Visible = False
btn_cancelar.Visible = False

'-----------------------------------------------------------------------
 
    TabStrip2.SelectedItem = TabStrip2.Tabs(1)
    TabStrip3.SelectedItem = TabStrip3.Tabs(1)
    TabStrip4.SelectedItem = TabStrip4.Tabs(1)
 
    If SSTab10.Item(0).Tab = 0 Then
        If colaborador_adm = 1 Then
             btn_alterar.Enabled = True
             btn_deletar.Enabled = True
            Else
                btn_alterar.Enabled = False
                btn_deletar.Enabled = False
        End If
    Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda where INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
    Adodc1.Refresh
    tel.Refresh
    End If
    
    
    If SSTab10.Item(0).Tab = 1 Then
        If colaborador_adm = 1 Then
            btn_alterar.Enabled = True
            btn_deletar.Enabled = True
            Else
                 btn_alterar.Enabled = False
                 btn_deletar.Enabled = False
        End If
    Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda where INTERNO = 0 and ID_COLABORADOR = 0 ORDER BY NOME"
    Adodc1.Refresh
    tel.Refresh
    End If
    
    
    If SSTab10.Item(0).Tab = 2 Then
    btn_alterar.Enabled = True
    btn_deletar.Enabled = True
    Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda where ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
    Adodc1.Refresh
    tel.Refresh
    End If
    

End Sub


Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub

Private Sub TabStrip2_Click()

    Select Case TabStrip2.SelectedItem

       Case Is = TabStrip2.Tabs(1)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'A%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh

    
      Case Is = TabStrip2.Tabs(2)
      
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'B%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    
    Case Is = TabStrip2.Tabs(3)
        
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'C%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(4)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'D%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh

    Case Is = TabStrip2.Tabs(5)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'E%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        

    Case Is = TabStrip2.Tabs(6)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'F%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        

    Case Is = TabStrip2.Tabs(7)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'G%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
           
    Case Is = TabStrip2.Tabs(8)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'H%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    
    Case Is = TabStrip2.Tabs(9)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'I%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    
    Case Is = TabStrip2.Tabs(10)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'J%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    
    Case Is = TabStrip2.Tabs(11)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'K%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(12)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'L%' AND INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case TabStrip2.Tabs(13)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'M%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    
    Case Is = TabStrip2.Tabs(14)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'N%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    Case Is = TabStrip2.Tabs(15)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'O%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    Case Is = TabStrip2.Tabs(16)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'P%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(17)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Q%' AND INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(18)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'R%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    
    Case Is = TabStrip2.Tabs(19)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'S%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
    
    Case Is = TabStrip2.Tabs(20)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'T%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(21)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'U%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(22)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'V%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(23)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'W%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(24)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'X%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
         
    Case Is = TabStrip2.Tabs(25)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Y%' AND  INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip2.Tabs(26)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Z%' AND INTERNO = -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        

    Case Else
    
         MsgBox "Houve um erro na escolha dos filtros, contate o administrador do sistema.", vbExclamation + vbOKOnly + vbApplicationModal, "AGENDA AVISO"
         
         
End Select


End Sub

Private Sub TabStrip3_Click()

Select Case TabStrip3.SelectedItem

       Case Is = TabStrip3.Tabs(1)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'A%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh

      Case Is = TabStrip3.Tabs(2)
      
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'B%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    
    Case Is = TabStrip3.Tabs(3)
        
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'C%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip3.Tabs(4)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'D%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       

    Case Is = TabStrip3.Tabs(5)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'E%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        

    Case Is = TabStrip3.Tabs(6)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'F%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       

    Case Is = TabStrip3.Tabs(7)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'G%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
           
    Case Is = TabStrip3.Tabs(8)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'H%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip3.Tabs(9)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'I%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip3.Tabs(10)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'J%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip3.Tabs(11)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'K%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip3.Tabs(12)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'L%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case TabStrip3.Tabs(13)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'M%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip3.Tabs(14)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'N%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip3.Tabs(15)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'O%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
     
    
    Case Is = TabStrip3.Tabs(16)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'P%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
      
        
    Case Is = TabStrip3.Tabs(17)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Q%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip3.Tabs(18)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'R%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip3.Tabs(19)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'S%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip3.Tabs(20)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'T%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip3.Tabs(21)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'U%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip3.Tabs(22)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'V%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip3.Tabs(23)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'W%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
     
        
    Case Is = TabStrip3.Tabs(24)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'X%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
         
    Case Is = TabStrip3.Tabs(25)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Y%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip3.Tabs(26)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Z%' AND  INTERNO <> -1 and ID_COLABORADOR = 0 ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       

    Case Else
    
         MsgBox "Houve um erro na escolha dos filtros, contate o administrador do sistema.", vbExclamation + vbOKOnly + vbApplicationModal, "AGENDA AVISO"
         
         
End Select

End Sub

Private Sub TabStrip4_Click()

Select Case TabStrip4.SelectedItem

       Case Is = TabStrip4.Tabs(1)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'A%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh

      Case Is = TabStrip4.Tabs(2)
      
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'B%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    
    Case Is = TabStrip4.Tabs(3)
        
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'C%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip4.Tabs(4)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'D%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       

    Case Is = TabStrip4.Tabs(5)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'E%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        

    Case Is = TabStrip4.Tabs(6)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'F%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       

    Case Is = TabStrip4.Tabs(7)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'G%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
           
    Case Is = TabStrip4.Tabs(8)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'H%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip4.Tabs(9)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'I%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip4.Tabs(10)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'J%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip4.Tabs(11)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'K%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case Is = TabStrip4.Tabs(12)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'L%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
        
        
    Case TabStrip4.Tabs(13)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'M%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip4.Tabs(14)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'N%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip4.Tabs(15)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'O%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
     
    
    Case Is = TabStrip4.Tabs(16)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'P%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
      
        
    Case Is = TabStrip4.Tabs(17)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Q%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip4.Tabs(18)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'R%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip4.Tabs(19)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'S%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
    
    Case Is = TabStrip4.Tabs(20)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'T%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip4.Tabs(21)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'U%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip4.Tabs(22)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'V%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip4.Tabs(23)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'W%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
     
        
    Case Is = TabStrip4.Tabs(24)
    
        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'X%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
         
    Case Is = TabStrip4.Tabs(25)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Y%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       
        
    Case Is = TabStrip4.Tabs(26)

        Adodc1.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'Z%' AND  ID_COLABORADOR =" & colaborador_id & " ORDER BY NOME"
        Adodc1.Refresh
        Adodc2.Refresh
       

    Case Else
    
         MsgBox "Houve um erro na escolha dos filtros, contate o administrador do sistema.", vbExclamation + vbOKOnly + vbApplicationModal, "AGENDA AVISO"
         
         
End Select

End Sub

Private Sub txt_nomeadd_KeyPress(KeyAscii As Integer)

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

Private Sub txt_nomealt_KeyPress(KeyAscii As Integer)

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

