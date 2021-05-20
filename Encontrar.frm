VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Encontrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agenda - Pesquisar os contatos na agenda "
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9960
   Icon            =   "Encontrar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Encontrar.frx":1082
   ScaleHeight     =   5640
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Campo para pesquisa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   -15
      TabIndex        =   1
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton botaolimpacampos 
         Caption         =   "Limpar"
         DragIcon        =   "Encontrar.frx":B827F
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   8250
         Picture         =   "Encontrar.frx":C0B51
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1050
         Width           =   1690
      End
      Begin VB.CommandButton btn_fechar 
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   8250
         Picture         =   "Encontrar.frx":C1BD3
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Sair do programa"
         Top             =   135
         Width           =   1690
      End
      Begin VB.TextBox pesquisanumero 
         Height          =   390
         Left            =   240
         TabIndex        =   3
         Top             =   1410
         Width           =   6990
      End
      Begin VB.TextBox pesquisanome 
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   630
         Width           =   7005
      End
      Begin MSAdodcLib.Adodc bdpesquisar 
         Height          =   345
         Left            =   2835
         Top             =   225
         Visible         =   0   'False
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   609
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
         Connect         =   $"Encontrar.frx":C2C55
         OLEDBString     =   $"Encontrar.frx":C2D46
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'NULL%' "
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
      Begin VB.Label Label5 
         Caption         =   "Número:"
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
         Left            =   225
         TabIndex        =   5
         Top             =   1125
         Width           =   915
      End
      Begin VB.Label Label5 
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
         Left            =   225
         TabIndex        =   4
         Top             =   345
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid tabelapesquisar 
      Bindings        =   "Encontrar.frx":C2E37
      Height          =   4185
      Left            =   -30
      TabIndex        =   0
      Top             =   1980
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   7382
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      DefColWidth     =   208
      HeadLines       =   2
      RowHeight       =   18
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
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "RESULTADOS DA PESQUISA"
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
End
Attribute VB_Name = "Encontrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub botaolimpacampos_Click()
pesquisanome.Text = ""
pesquisanumero.Text = ""
End Sub

Private Sub btn_fechar_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dadospesquisar_Click()

End Sub


Private Sub pesquisanome_Change()



If pesquisanumero.Text = "" Then

    If pesquisanome.Text = "" Then

    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like 'NULL%' "
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    Else

    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like '" & pesquisanome.Text & "%'"
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    End If
    
Else

    If pesquisanome.Text = "" Then

    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NUMERO like '" & pesquisanumero.Text & "%' or NUMEROS_ALTERNATIVOS like '" & pesquisanumero.Text & "%'"
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    Else

    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE ( NOME like '" & pesquisanome.Text & "%' and NUMERO like '" & pesquisanumero.Text & "%') or ( NUMEROS_ALTERNATIVOS like '" & pesquisanumero.Text & "%' and NOME like '" & pesquisanome.Text & "%')"
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    End If

End If

End Sub


Private Sub pesquisanumero_Change()


If pesquisanome.Text = "" Then

    If pesquisanumero.Text = "" Then

    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NUMERO like 'NULL%' "
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    Else

    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NUMERO like '" & pesquisanumero.Text & "%' or NUMEROS_ALTERNATIVOS like '" & pesquisanumero.Text & "%'"
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    End If
    
Else

    If pesquisanumero.Text = "" Then

    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE NOME like '" & pesquisanome.Text & "%'"
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    Else
    
    bdpesquisar.RecordSource = "Select NOME, NUMERO, NUMEROS_ALTERNATIVOS from agenda WHERE ( NOME like '" & pesquisanome.Text & "%' and NUMERO like '" & pesquisanumero.Text & "%') or ( NUMEROS_ALTERNATIVOS like '" & pesquisanumero.Text & "%' and NOME like '" & pesquisanome.Text & "%')"
    bdpesquisar.Refresh
    tabelapesquisar.Refresh

    End If

End If


End Sub
Private Sub Form_load()

tabelapesquisar.HeadFont.Bold = True
tabelapesquisar.HeadFont.Size = 9
tabelapesquisar.Font.Size = 10

End Sub

