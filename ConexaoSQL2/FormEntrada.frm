VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormEntrada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   17175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtIdEntrada 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3015
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      Top             =   315
      Width           =   1200
   End
   Begin VB.TextBox TxtFornecedor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1185
      Width           =   1200
   End
   Begin VB.TextBox TxtCusto 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14175
      MaxLength       =   12
      TabIndex        =   9
      Top             =   630
      Width           =   1700
   End
   Begin VB.TextBox TxtDescontoEntrada 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14175
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1335
      Width           =   1200
   End
   Begin VB.TextBox TxtCustoFinal 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14175
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2025
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produto"
      Height          =   3540
      Left            =   9375
      TabIndex        =   22
      Top             =   2850
      Width           =   7545
      Begin VB.CommandButton BtnNovoProduto 
         Caption         =   "Novo Produto"
         Height          =   375
         Left            =   3420
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1995
      End
      Begin VB.CommandButton BtnAdicionarItem 
         Caption         =   "Adicionar"
         Height          =   630
         Left            =   4575
         TabIndex        =   8
         Top             =   2790
         Width           =   1245
      End
      Begin VB.CommandButton BtnRemoverItem 
         Caption         =   "Remover"
         Height          =   630
         Left            =   5925
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2775
         Width           =   1245
      End
      Begin VB.TextBox TxtCodItem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1350
         MaxLength       =   12
         TabIndex        =   4
         Top             =   375
         Width           =   1200
      End
      Begin VB.TextBox TxtItemQuantidade 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1365
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1875
         Width           =   1200
      End
      Begin VB.TextBox TxtDescontoItem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2730
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1875
         Width           =   1200
      End
      Begin VB.TextBox TxtNomeItem 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   375
         Width           =   4095
      End
      Begin VB.TextBox TxtCustoFinalItem 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1455
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1700
      End
      Begin VB.CommandButton BtnPesquisarProduto 
         Caption         =   "->"
         Height          =   405
         Left            =   2670
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   405
         Width           =   500
      End
      Begin VB.TextBox TxtCustoItem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1365
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1170
         Width           =   1200
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estoque"
         Height          =   1740
         Left            =   5685
         TabIndex        =   23
         Top             =   915
         Width           =   1620
         Begin VB.Line LinhaSomaEstoque 
            X1              =   195
            X2              =   1425
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label LblEstoqueFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   130
            TabIndex        =   47
            Top             =   1230
            Width           =   1275
         End
         Begin VB.Label LblQuantiaEntrada 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "quantia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   135
            TabIndex        =   46
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label LblEstoqueItem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "estoque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   130
            TabIndex        =   24
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.Label Label12 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4005
         TabIndex        =   48
         Top             =   1920
         Width           =   285
      End
      Begin VB.Label Label8 
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   33
         Top             =   465
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Quantia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   32
         Top             =   1965
         Width           =   1080
      End
      Begin VB.Label Label11 
         Caption         =   "Desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2685
         TabIndex        =   31
         Top             =   1425
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Custo Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1485
         TabIndex        =   30
         Top             =   2505
         Width           =   1530
      End
      Begin VB.Label Label14 
         Caption         =   "Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   285
         TabIndex        =   29
         Top             =   1200
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   6210
      TabIndex        =   19
      Top             =   45
      Width           =   6525
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "Cancelar"
         Height          =   585
         Left            =   2325
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   180
         Width           =   1830
      End
      Begin VB.CommandButton BtnFinalizar 
         Caption         =   "Imprimir"
         Height          =   570
         Left            =   4365
         TabIndex        =   11
         Top             =   195
         Width           =   1830
      End
      Begin VB.CommandButton BtnNovo 
         Caption         =   "Novo"
         Height          =   585
         Left            =   270
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   195
         Width           =   1830
      End
   End
   Begin VB.TextBox TxtFormaPgto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1965
      Width           =   1200
   End
   Begin VB.CommandButton BtnRecalcularCusto 
      Caption         =   "Recalcular Custos"
      Height          =   315
      Left            =   14085
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   135
      Width           =   1785
   End
   Begin VB.CommandButton btnEntrada 
      Caption         =   "->"
      Height          =   405
      Left            =   2310
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   345
      Width           =   500
   End
   Begin VB.CommandButton BtnFornecedor 
      Caption         =   "->"
      Height          =   405
      Left            =   2295
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1215
      Width           =   500
   End
   Begin VB.CommandButton BtnFormaPgto 
      Caption         =   "->"
      Height          =   405
      Left            =   2295
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1995
      Width           =   500
   End
   Begin VB.CommandButton BtnLiberarAlteracao 
      Caption         =   "*"
      Height          =   405
      Left            =   16110
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Solicitar Desbloqueio para operador ADMIN"
      Top             =   690
      Width           =   510
   End
   Begin VB.CommandButton BtnAvancaRegistro 
      Caption         =   ">"
      Height          =   555
      Left            =   5205
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   285
      Width           =   675
   End
   Begin VB.CommandButton BtnVoltaRegistro 
      Caption         =   "<"
      Height          =   555
      Left            =   4440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   675
   End
   Begin MSComctlLib.ListView ListViewItensEntrada 
      Height          =   3510
      Left            =   270
      TabIndex        =   14
      Top             =   2865
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   6191
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CodItensPedido"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nome"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantidade"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Desconto"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label15 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15465
      TabIndex        =   49
      Top             =   1365
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   540
      TabIndex        =   44
      Top             =   300
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   90
      TabIndex        =   43
      Top             =   1185
      Width           =   2100
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Custo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12825
      TabIndex        =   42
      Top             =   645
      Width           =   1185
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Desconto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12135
      TabIndex        =   41
      Top             =   1305
      Width           =   1890
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Custo Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11880
      TabIndex        =   40
      Top             =   2010
      Width           =   2130
   End
   Begin VB.Label LblFornecedorNome 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4860
      TabIndex        =   39
      Top             =   1200
      Width           =   6645
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4230
      TabIndex        =   38
      Top             =   1170
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cond Pgt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   525
      TabIndex        =   37
      Top             =   1950
      Width           =   1635
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4230
      TabIndex        =   36
      Top             =   1995
      Width           =   600
   End
   Begin VB.Label LblFormaPgto 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4860
      TabIndex        =   35
      Top             =   2025
      Width           =   6645
   End
End
Attribute VB_Name = "FormEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STATUS As Integer ' 1 -> JA FINALIZADO, 2 -> NOVO

Private Sub BtnAdicionarItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnAvancaRegistro_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdEntrada.Text & " " & _
          "SELECT TOP 1 IdEntrada FROM Entrada WHERE IdEntrada > @ID ORDER BY IdEntrada"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            ID = rs("IdEntrada")
        Else
            ID = TxtIdEntrada.Text
        End If
    rs.Close
    
    TxtIdEntrada.Text = ID
    preencheEntrada (ID)

End Sub

Private Sub BtnAvancaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnCancelar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnFornecedor_Click()
    FormBuscaCliente.FORMULARIO = "FormEntrada"
    FormBuscaCliente.Show
    FormBuscaCliente.SetFocus
    MDIFormInicio.Enabled = False
End Sub

Private Sub BtnFinalizar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnFormaPgto_Click()
    FormBuscaFormaPGTO.FORMULARIO = "FormEntrada"
    FormBuscaFormaPGTO.Show
    FormBuscaFormaPGTO.SetFocus
    MDIFormInicio.Enabled = False
End Sub

Private Sub BtnLiberarAlteracao_Click()
    
    MDIFormInicio.Enabled = False
    FormSolicitaAcesso.FORMULARIO = "FormEntrada"
    FormSolicitaAcesso.Show
    FormSolicitaAcesso.TxtLogin.SetFocus
    
End Sub

Private Sub BtnNovo_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim ID As Integer
    
    SQL = "SELECT MAX(IdEntrada)+1 AS Entrada FROM Entrada"
    
    rs.Open SQL, cn, adOpenStatic
        ID = rs("Entrada")
        TxtIdEntrada.Text = rs("Entrada")
    rs.Close
    
    ' SETO O STATUS DA TELA PARA 0 -> NOVO REGISTRO
    STATUS = 0
    preencheEntrada (ID)
    
End Sub

Private Sub BtnNovo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnNovoProduto_Click()

    FormCadastroProduto.Show
    FormCadastroProduto.BtnNovo_Click

End Sub

Private Sub BtnPesquisarProduto_Click()

    FormBuscaProduto.FORMULARIO = "FormEntrada"
    FormBuscaProduto.Show
    FormBuscaProduto.SetFocus
    MDIFormInicio.Enabled = False

End Sub

Private Sub BtnRemoverItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnVoltaRegistro_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdEntrada.Text & " " & _
          "SELECT TOP 1 IdEntrada FROM Entrada WHERE IdEntrada < @ID ORDER BY IdEntrada DESC"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            ID = rs("IdEntrada")
        Else
            ID = TxtIdEntrada.Text
        End If
    rs.Close
    
    TxtIdEntrada.Text = ID
    preencheEntrada (ID)

End Sub

Private Sub BtnVoltaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub Form_Load()
    
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    Dim SQL As String
    
    On Error GoTo TrataErro
    
    SQL = "DECLARE @maxID INT " & _
          "SELECT @maxID = MAX(IdEntrada) FROM Entrada " & _
          "SELECT @maxID AS Entrada"
             
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Entrada")) Then
            STATUS = 0
            ID = 1
            TxtIdEntrada.Text = 1
        Else
            ' SETO O STATUS DA TELA PARA 1 -> REGISTRO JÁ FINALIZADO
            STATUS = 1
            ID = rs("Entrada")
        End If
    rs.Close
    ' LIMPA O RECORDSET
    Set rs = Nothing
    
    ' LIMPA OS CAMPOS DO ITEM
    limpaCamposItem
    
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preencheEntrada (ID)
    
Exit Sub
TrataErro:
    MsgBox "Algum erro ocorreu ao carregar o Form - " & Me.Name & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub btnEntrada_Click()
    'FormBuscaPedido.Show
    'FormBuscaPedido.SetFocus
    'MDIFormInicio.Enabled = False
End Sub

Private Sub BtnAdicionarItem_Click()

    On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim SQLVerificaEntrada As String
    Dim SQLVerificarProduto As String
    Dim rsVerificarProduto As New ADODB.Recordset
    Dim descontoItem As String
    Dim EstoqueFinal As Long
    Dim CustoEntrada As Double
    Dim custoItem As String
    Dim validacaoItem As Integer
    
    If TxtCodItem = "" Then
        MsgBox "Campo Produto não pode estar vazio"
        TxtCodItem.SetFocus
        'NÃO SEGUE ADIANTE
        Exit Sub
    ElseIf TxtItemQuantidade = "" Then
        MsgBox "Campo quantidade não pode estar vazio"
        TxtItemQuantidade.SetFocus
        'NÃO SEGUE ADIANTE
        Exit Sub
    Else
        If TxtDescontoItem = "" Then
            TxtDescontoItem.Text = 0
        End If
        
        'Alterando a "," por "." para não dar erro no SQL ao adicionar o desconto
        descontoItem = Replace(TxtDescontoItem.Text, ",", ".")
        custoItem = Replace(TxtCustoItem.Text, ",", ".")
        
        SQLVerificaEntrada = "SELECT IdEntrada FROM Entrada"
        
        rs.Open SQLVerificaEntrada, cn, adOpenStatic
            If rs.EOF = True Then
                'ADICIONA O PRODUTO NA TABELA DE PRODUTOS PARA O PEDIDO 1 (PRIMEIRO REGISTRO)
                SQL = "DECLARE @CustoTotal decimal(18,2) " & _
                      "SELECT @CustoTotal = ((" & custoItem & " * " & TxtItemQuantidade.Text & ") - ((" & custoItem & " * " & TxtItemQuantidade.Text & ") * (" & descontoItem & " * 0.01))) " & _
                      "INSERT INTO ItensEntrada (IdEntrada, IdProduto, CustoProduto, QuantidadeProduto, DescontoProduto, CustoTotalProduto) " & _
                      "VALUES (1, " & TxtCodItem.Text & ", " & custoItem & ", " & TxtItemQuantidade.Text & ", " & descontoItem & ", @CustoTotal)"
            Else
                'ADICIONA O PRODUTO NA TABELA DE PRODUTOS PARA O PEDIDO (NORMALMENTE)
                SQL = "INSERT INTO ItensEntrada (IdEntrada, IdProduto, CustoProduto, QuantidadeProduto, DescontoProduto, CustoTotalProduto) " & _
                      "VALUES (" & TxtIdEntrada.Text & ", " & TxtCodItem.Text & ", " & custoItem & ", " & TxtItemQuantidade.Text & ", " & descontoItem & ", ((" & custoItem & " * " & TxtItemQuantidade.Text & ") - ((" & custoItem & " * " & TxtItemQuantidade.Text & ") * (" & descontoItem & " * 0.01)))) "
            End If
        rs.Close
        ' LIMPA O RECORDSET
        Set rs = Nothing
    
        'FAZ A VERIFICAÇÃO SE EXISTE O PRODUTO QUE SERÁ INSERIDO
        SQLVerificarProduto = "SELECT CodProduto FROM Produtos WHERE CodProduto = " & TxtCodItem.Text
        
        rsVerificarProduto.Open SQLVerificarProduto, cn, adOpenStatic
            'SE O PRODUTO EXISTE
            If rsVerificarProduto.EOF = False Then
                ' FAZ O CALCULO DO NOVO ESTOQUE
                EstoqueFinal = LblEstoqueItem.Caption + TxtItemQuantidade.Text
                
                ' CHAMA A FUNÇÃO QUE VERIFICA SE EXISTE O PRODUTO JÁ INFORMADO NOS ITENS DA ENTRADA & PASSA O SQL PARA FAZER O INSERT
                 validacaoItem = verificaItensEntrada(TxtCodItem.Text)
                 
                 If validacaoItem = 1 Then
                    cn.Execute (SQL)
                 End If
            Else
                'ERRO CASO NÃO TENHA O PRODUTO CADASTRADO NO SISTEMA
                MsgBox "Produto não cadastrado", vbCritical
            End If
        rsVerificarProduto.Close
        
        'PEGA OS DADOS DA TABLE E ADICIONA NO GRID
        atualizaListaEntrada
                
        'Zerando os valores ao inserir o produto
        limpaCamposItem
                    
        'SELECT para buscar a soma total dos valores do pedido, SE NÃO TIVER NENHUM VALOR RETORNA 0
        SQL = "DECLARE @CustoTotal DECIMAL(18,2) " & _
              "SELECT @CustoTotal = SUM(itens.CustoTotalProduto) " & _
              "FROM ItensEntrada AS itens " & _
              "WHERE itens.IdEntrada = " & TxtIdEntrada.Text & " " & _
              "SELECT @CustoTotal = ISNULL(@CustoTotal, 0) " & _
              "SELECT @CustoTotal AS CustoTotal"
        'Adicionando o valorTotal no TextField
        
        rs.Open SQL, cn, adOpenStatic
            TxtCusto.Text = Format(rs("CustoTotal"), "0.00")
        rs.Close
        
        'Verificando se o Desconto de Pedido está nulo, se sim coloca valor 0
        If TxtDescontoEntrada = "" Then
            TxtDescontoEntrada.Text = 0
        End If
        
        'Volta o foco para o codItem
        TxtCodItem.SetFocus
    End If
    
    'SEMPRE ATUALIZA O VALOR DO PEDIDO
    RecalcularCustoEntrada
Exit Sub

TrataErro:
    MsgBox " AÇÃO CANCELADA - Ocorreu um erro durante a tentativa de inserir o item: " & TxtCodItem.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub

Private Sub BtnRemoverItem_Click()
    
    On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim X As Integer
         
    If ListViewItensEntrada.ListItems.Count > 0 Then
        If MsgBox("Confirmar a exclusão deste Item?", vbYesNo) = vbYes Then
            
            SQL = "DELETE FROM ItensEntrada WHERE IdItensEntrada = " & ListViewItensEntrada.SelectedItem
            
            cn.Execute (SQL)
            
            'ATUALIZA A LISTA DE ITENS
            atualizaListaEntrada
            'RECALCULA O VALOR DO PEDIDO
            RecalcularCustoEntrada
        End If
        
        RecalcularCustoEntrada
    Else
        MsgBox ("Nenhum item Selecionado")
    End If
Exit Sub

TrataErro:
    MsgBox "Algum erro ocorreu ao tentar excluír o item - " & TxtCodItem.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim SQL As String

    If STATUS = 0 Then
        'Limpa os itens do pedido
        SQL = "DELETE FROM ItensEntrada WHERE IdEntrada = " & TxtIdEntrada.Text
        cn.Execute SQL
    End If
    
    ' HABILITA O MENU
    MDIFormInicio.Menu.Enabled = True
    MDIFormInicio.Relatorio.Enabled = True
    MDIFormInicio.Configurações.Enabled = True
    
End Sub

Private Sub ListViewItensEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtFornecedor_GotFocus()

    With TxtFornecedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Debug.Print KeyCode
    
    If KeyCode = 115 Then ' F4
        BtnFornecedor_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtFornecedor_KeyPress(KeyAscii As Integer)

    On Error GoTo TrataErro
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER key.
        ' VERIFICA SE O CAMPO NÃO ESTÁ VAZIO
        If TxtFornecedor.Text = "" Then
            MsgBox ("Necessário informar um Fornecedor")
            KeyAscii = 0
        Else
            ' CHAMA A FUNÇÃO QUE ADICIONA O NOME DO FORNECEDOR NO LABEL
            If PegaNomeFornecedor = 0 Then
                Exit Sub
            End If
            KeyAscii = 0
        End If
    End If
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro durante a tentativa de inserir o Fornecedor: " & TxtFornecedor.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub TxtCodItem_GotFocus()

    With TxtCodItem
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtCodItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    ElseIf KeyCode = 115 Then ' F4
        BtnPesquisarProduto_Click
    End If
End Sub

Private Sub TxtCodItem_LostFocus()
    
    preencheProduto
    
End Sub

Private Sub TxtDescontoItem_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
    
End Sub

Private Sub TxtDescontoEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtFormaPgto_GotFocus()

    With TxtFormaPgto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtFormaPgto_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    ElseIf KeyCode = 115 Then ' F4
        BtnFormaPgto_Click
    End If
End Sub

Private Sub TxtItemQuantidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtItemQuantidade_LostFocus()
    
    If TxtCodItem.Text = "" Then
        ' NÃO ATUALIZA O VALOR DO ITEM
    Else
        ' CHAMA A FUNÇÃO QUE ATUALIZA O VALOR DO ITEM
        atualizaCustoItem
    End If

End Sub

Private Sub TxtNomeItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtIdEntrada_GotFocus()

    With TxtIdEntrada
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtIdEntrada_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 115 Then ' F4
        btnEntrada_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If
    
End Sub

Private Sub TxtFormaPgto_KeyPress(KeyAscii As Integer)

    On Error GoTo TrataErro
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER KEY
        ' VERIFICA SE O CAMPO NÃO ESTÁ VAZIO
        If TxtFormaPgto.Text = "" Then
            MsgBox ("Necessário informar uma Forma de Pagamento")
            KeyAscii = 0
        Else
            ' CHAMA A FUNÇÃO QUE ADICIONA O NOME DA FORMA DE PAGAMENTO NO LABEL
            PegaNomeFormaPgto
            KeyAscii = 0
        End If
    End If
    
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro durante a tentativa de inserir a Forma de Pagamento: " & TxtFormaPgto.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub TxtCodItem_KeyPress(KeyAscii As Integer)

    On Error GoTo TrataErro
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER KEY
        ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS DO PRODUTO (LOST FOCUS)
        SendKeys ("{tab}")
        KeyAscii = 0
    End If
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro ao procurar o produto no banco: " & TxtCodItem & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub

Public Function preencheProduto()
    
    Dim rsVerificarProduto As New ADODB.Recordset
    Dim SQLVerificarProduto As String
    Dim nomeProduto As String
    Dim EstoqueItem As Long
    
    If TxtCodItem = "" Then
        'MsgBox "Campo Produto não pode estar vazio"
        
        'CHAMA A FUNÇÃO DE LIMPAR OS CAMPOS DO ITEM
        limpaCamposItem
    Else
        'FAZ A VERIFICAÇÃO SE EXISTE O PRODUTO QUANDO APERTAR ENTER BUSCANDO PELO NOME
        SQLVerificarProduto = "SELECT NomeProduto, StatusProduto FROM Produtos WHERE CodProduto = " & TxtCodItem
        
        rsVerificarProduto.Open SQLVerificarProduto, cn, adOpenStatic
            If rsVerificarProduto.EOF = False Then
                nomeProduto = rsVerificarProduto("NomeProduto")
                If rsVerificarProduto("StatusProduto") = 0 Then
                    'SE O PRODUTO EXISTE, PASSA PRO PROXIMO TAB E ADICIONA O NOME NO TXT NOME ITEM
                    TxtNomeItem.Text = nomeProduto
                    
                    ' ATIVA A LINHA DA SOMA E ADICIONA O TOTAL DO ESTOQUE
                    LinhaSomaEstoque.Visible = True
                    LblEstoqueFinal.Caption = EstoqueItem + 1
                    
                    ' ADICIONA O VALOR 1 NO TXT DA QUANTIDADE
                    TxtItemQuantidade.Text = 1
                    
                    'CHAMA A FUNÇÃO PARA SETAR O CUSTO DO PRODUTO
                    CustoProduto
                    
                    'CHAMA A FUNÇÃO PARA ATUALIZAR O CUSTO DO ITEM
                    atualizaCustoItem
                    'SendKeys "{tab}"
                Else
                    'ERRO CASO O PRODUTO CADASTRADO NO SISTEMA ESTEJA INATIVO
                    MsgBox "Produto " & TxtCodItem.Text & " - " & nomeProduto & " - Inativo!"
                    TxtCodItem.SetFocus
                    'CHAMA A FUNÇÃO DE LIMPAR OS CAMPOS DO ITEM
                    limpaCamposItem
                End If
            Else
                'ERRO CASO NÃO TENHA O PRODUTO CADASTRADO NO SISTEMA
                MsgBox "Produto " & TxtCodItem.Text & " não cadastrado"
                TxtCodItem.SetFocus
                'CHAMA A FUNÇÃO DE LIMPAR OS CAMPOS DO ITEM
                limpaCamposItem
            End If
        rsVerificarProduto.Close
    End If
    
End Function

Private Sub TxtIdEntrada_KeyPress(KeyAscii As Integer)
        
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
        
    If KeyAscii = 13 Then ' The ENTER key.
        If TxtIdEntrada.Text = "" Then
            MsgBox ("Necessário informar um ID de Entrada!")
            KeyAscii = 0
        Else
            If STATUS = 1 Then
                ' FUNÇÃO QUE VERIFICA SE EXISTE ESTE PEDIDO NO SQL
                buscaEntrada (TxtIdEntrada.Text)
                KeyAscii = 0
            Else
                SendKeys ("{tab}") ' MOVE PARA O PRÓXIMO CAMPO
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TxtCusto_GotFocus()

    With TxtCusto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtCusto_LostFocus()
    
    Dim descontoPedido As String
               
    'Formata o número para apenas 2 casas decimais
    TxtCusto.Text = Format(TxtCusto.Text, "0.00")
            
    'CHAMA A FUNÇÃO PARA FAZER O CALCULO DA ENTRADA
    RecalcularCustoEntrada
End Sub

'Formatação <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub TxtCustoFinal_Change()
    'Formata o número para apenas 2 casas decimais
    TxtCustoFinal.Text = Format(TxtCustoFinal.Text, "0.00")
End Sub
'Formatação >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub TxtItemQuantidade_GotFocus()
    ' selecionar o texto ao receber o foco
    With TxtItemQuantidade
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtItemQuantidade_KeyPress(KeyAscii As Integer)

    On Error GoTo TrataError

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If

    If KeyAscii = 13 Then ' The ENTER KEY
        If TxtCodItem.Text = "" Then
            MsgBox ("Necessário informar um produto!")
            TxtItemQuantidade.Text = ""
            TxtCodItem.SetFocus
            KeyAscii = 0
        Else
            If TxtItemQuantidade = "" Then
                TxtItemQuantidade.Text = 1
                SendKeys "{tab}"   ' Set the focus to the next control
                KeyAscii = 0
            ElseIf TxtItemQuantidade = 0 Then
                MsgBox "Campo quantidade não pode ser 0"
                TxtItemQuantidade.Text = ""
                KeyAscii = 0
            Else
                'CHAMA A FUNÇÃO PARA ATUALIZAR O VALOR DO ITEM
                atualizaCustoItem
                SendKeys "{tab}"   ' Set the focus to the next control.
                KeyAscii = 0       ' Ignore this key.
            End If
        End If
    End If
Exit Sub

TrataError:
    MsgBox " Ocorreu um erro ao informar a quantidade do item: " & TxtCodItem.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub

Private Sub TxtDescontoItem_KeyPress(KeyAscii As Integer)

    On Error GoTo TrataError
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER KEY
        If TxtCodItem.Text = "" Then
            MsgBox ("Necessário ter um item informado!")
            ' CHAMA A FUNÇÃO QUE LIMPA O ITEM
            limpaCamposItem
            ' VOLTA O FOCO PARA O CÓDIGO
            TxtCodItem.SetFocus
        Else
            'CHAMA A FUNÇÃO PARA ATUALIZAR O VALOR DO ITEM (LOST FOCUS)
            SendKeys "{tab}"   ' Set the focus to the next control.
            KeyAscii = 0       ' Ignore this key.
        End If
    End If
    
Exit Sub
TrataError:
    MsgBox " Ocorreu um erro ao informar o desconto do item: " & TxtCodItem.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

'GotFocus para selecionar todo o valor do desconto
Private Sub TxtDescontoItem_GotFocus()
    With TxtDescontoItem
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtDescontoItem_LostFocus()
    'Formata o número para apenas 2 casas decimais
    TxtDescontoItem.Text = Format(TxtDescontoItem.Text, "0.00")
    
    If TxtCodItem.Text = "" Then
        ' NÃO ATUALIZA O VALOR DO ITEM
    Else
        ' CHAMA A FUNÇÃO QUE ATUALIZA O VALOR DO ITEM
        atualizaCustoItem
    End If
    
End Sub

Private Sub TxtDescontoEntrada_KeyPress(KeyAscii As Integer)
    
    On Error GoTo TrataError

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If

    If KeyAscii = 13 Then ' The ENTER KEY
        SendKeys "{tab}"   ' Set the focus to the next control.
        KeyAscii = 0       ' Ignore this key.
    End If
Exit Sub
    
TrataError:
    MsgBox " Ocorreu um erro ao informar a o desconto: " & TxtDescontoEntrada.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

'GotFocus para selecionar todo o valor do desconto Total
Private Sub TxtDescontoEntrada_GotFocus()

    With TxtDescontoEntrada
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub TxtDescontoEntrada_LostFocus()
    
    Dim descontoEntrada As String
    Dim SQL As String
    Dim rs As New Recordset
    Dim CustoEntrada As Double
    
    SQL = "SELECT * FROM ConfiguracoesGerais"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            descontoEntrada = -1
        Else
            descontoEntrada = Format(rs("MaxDescontoEntrada"), "0.00")
        End If
    rs.Close
    
    If TxtCusto.Text = "" Then
        MsgBox "Necessário ter um valor informado!"
        TxtDescontoEntrada.Text = ""
        TxtCusto.SetFocus
    Else
        ' SÓ VERIFICA EM NOVOS PEDIDOS
        If STATUS = 0 Then
            If TxtDescontoEntrada.Text = "" Then
                TxtDescontoEntrada.Text = 0
            End If
            
            CustoEntrada = TxtDescontoEntrada.Text
            
            If descontoEntrada = -1 Then
                If CustoEntrada = 100 Then
                    MsgBox "Não é permitido dar uma desconto de 100%"
                    TxtDescontoEntrada.Text = 0
                    TxtDescontoEntrada.SetFocus
                ElseIf CustoEntrada > 99.994 Then
                    MsgBox "Valor de desconto inválido!"
                    TxtDescontoEntrada.Text = 0
                    TxtDescontoEntrada.SetFocus
                End If
            Else
                If CustoEntrada > descontoEntrada Then
                    MsgBox ("Desconto informado maior que " & descontoEntrada & " permitido!")
                    TxtDescontoEntrada.Text = 0
                    TxtDescontoEntrada.SetFocus
                End If
            End If
        End If
    End If
    'Formata o número para apenas 2 casas decimais
    TxtDescontoEntrada.Text = Format(TxtDescontoEntrada.Text, "0.00")
                
    'CHAMA A FUNÇÃO QUE FAZ OS CALCULOS DO VALOR
    RecalcularCustoEntrada
End Sub

Private Sub TxtCusto_KeyPress(KeyAscii As Integer)
    
    On Error GoTo TrataError

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If

    If KeyAscii = 13 Then ' The ENTER KEY
        SendKeys "{tab}"   ' Set the focus to the next control.
        KeyAscii = 0       ' Ignore this key.
    End If
Exit Sub
    
TrataError:
    MsgBox " Ocorreu um erro ao informar a o desconto: " & TxtDescontoEntrada.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub


Private Sub BtnFinalizar_Click()

    On Error GoTo TrataError

    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    Dim CustoAntigo As Double
    Dim validar As Integer
    Dim resultado As Integer
    Dim verificarImpressao As Integer
    Dim tipoOperacao As Integer
    Dim descontoEntrada As Double
    Dim Custo As Double
    Dim CustoFinal As Double
    Dim ID As Integer

    SQL = "SELECT * FROM ItensEntrada WHERE IdEntrada = " & TxtIdEntrada.Text

    If STATUS = 0 Then ' PEDIDO NOVO (FINALIZAR)
        'Fazendo verificações
        rs.Open SQL, cn, adOpenStatic
            If TxtFornecedor.Text = "" Then
                MsgBox "Campo forncedor não pode ser vazio"
                TxtFornecedor.SetFocus
            ElseIf TxtFormaPgto.Text = "" Then
                MsgBox "Campo Cond. Pagamento não pode ser vazio"
                TxtFormaPgto.SetFocus
            ElseIf rs.EOF = True Then
                MsgBox "Necessário informar ao menos um item!"
                TxtCodItem.SetFocus
            ElseIf TxtCusto.Text = "" Then
                MsgBox "Campo Custo não pode ser vazio"
                TxtCusto.SetFocus
            ElseIf TxtDescontoEntrada.Text = "" Then
                TxtDescontoEntrada.Text = 0
            ElseIf TxtCustoFinal.Text = "" Then
                MsgBox "Campo Custo Total não pode estar vazio"
                TxtCustoFinal.SetFocus
            Else
            
                If PegaNomeFornecedor = 0 Then
                    Exit Sub
                End If
            
               'SELECT para buscar a soma dos valores do pedido
                SQL = "SELECT SUM(itens.CustoTotalProduto) " & _
                      "FROM ItensEntrada AS itens " & _
                      "WHERE itens.IdEntrada = " & TxtIdEntrada.Text
                'Pega a soma dos itens
                CustoAntigo = cn.Execute(SQL).GetString
                'Valida se a soma dos itens é diferente do valor informado
                If CustoAntigo <> TxtCusto.Text Then
                    validar = MsgBox("Custo da entrada alterado manualmente, prosseguir com a finalização?", vbYesNo)
                    If validar = vbYes Then
                        'Alterando a "," por "." para não dar erro no SQL ao adicionar o desconto
                        descontoEntrada = Replace(TxtDescontoEntrada.Text, ",", ".")
                        Custo = Replace(TxtCusto.Text, ",", ".")
                        CustoFinal = Replace(TxtCustoFinal.Text, ",", ".")
                        
                        'Inserindo os valores na tabela PEDIDO
                        SQL = "INSERT INTO Entrada (EntradaIdCli, EntradaIdPgto, EntradaCusto, EntradaDesconto, EntradaCustoTotal) " & _
                              "VALUES (" & TxtFornecedor.Text & ", " & TxtFormaPgto.Text & ", " & Custo & ", " & descontoEntrada & ", " & CustoFinal & ")"
                        'Executando SQL
                        cn.Execute SQL
                        
                        verificarImpressao = MsgBox("Entrada Finalizada com sucesso! - Deseja imprimir a Entrada?", vbYesNo)
                        
                        If verificarImpressao = vbYes Then
                            imprimirEntrada
                        End If
                        
                        ' CHAMA A FUNÇÃO QUE ATUALIZA O CUSTO DOS PRODUTOS
                        atualizaCustoProduto
                            
                        ' FINALIZOU COM SUCESSO, SELECIONA A ULTIMA ENTRADA NO BANCO
                        SQL = "DECLARE @maxID INT " & _
                              "SELECT @maxID = MAX(IdEntrada) FROM Entrada " & _
                              "SELECT @maxID AS Entrada"
                          
                        rsDados.Open SQL, cn, adOpenStatic
                            ID = rsDados("Entrada")
                        rsDados.Close
                        
                        ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
                        STATUS = 1
                            
                        ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
                        preencheEntrada (ID)
                    Else
                        'Volta o foco para o TXTCUSTO
                        TxtCusto.SetFocus
                    End If
                Else
                    resultado = finalizarEntrada(TxtFornecedor.Text, TxtFormaPgto.Text, TxtCusto.Text, TxtDescontoEntrada.Text, TxtCustoFinal.Text)
                        
                    If resultado = 0 Then
                        verificarImpressao = MsgBox("Entrada Finalizado com sucesso! - Deseja imprimir a Entrada?", vbYesNo)
                        
                        If verificarImpressao = vbYes Then
                            imprimirEntrada
                        End If
                        
                        ' CHAMA A FUNÇÃO QUE ATUALIZA O CUSTO DOS PRODUTOS
                        atualizaCustoProduto
                        
                        ' FINALIZOU COM SUCESSO, SELECIONA A ULTIMA ENTRADA NO BANCO
                        SQL = "DECLARE @maxID INT " & _
                              "SELECT @maxID = MAX(IdEntrada) FROM Entrada " & _
                              "SELECT @maxID AS Entrada"
                          
                        rsDados.Open SQL, cn, adOpenStatic
                            ID = rsDados("Entrada")
                        rsDados.Close
                        
                        ' SETO O STATUS DA TELA PARA 1 -> REGISTRO JÁ FINALIZADO
                        STATUS = 1
                            
                        ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
                        preencheEntrada (ID)
                    ElseIf resultado = 1 Then
                        Exit Sub
                    End If
                End If
            End If
        rs.Close
    ElseIf STATUS = 1 Then ' REGISTRO JÁ FINALIZADO (IMPRIMIR)
        ' CHAMA A FUNÇÃO DE IMPRIMIR REGISTRO
        imprimirEntrada
    End If
Exit Sub

TrataError:
    MsgBox "Ocorreu um erro ao finalizar a entrada: " & TxtIdEntrada.Text & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub

Private Sub BtnCancelar_Click()
    
    On Error GoTo TrataErro
    
    ' CHAMA O FUNÇÃO QUE FINALIZA O FORM
    finalizaForm
    
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro ao tentar excluír a entrada N°: " & TxtIdEntrada.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub BtnRecalcularCusto_Click()
    
    On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim validar As Integer
    
    validar = MsgBox("Recalcular o custo da Entrada?", vbYesNo)
    
    If validar = vbYes Then
        'SELECT para buscar a soma dos valores do pedido
        SQL = "SELECT SUM(itens.CustoTotalProduto) " & _
              "FROM ItensEntrada AS itens " & _
              "WHERE itens.IdEntrada = " & TxtIdEntrada.Text
        
            rs.Open SQL, cn, adOpenStatic
                TxtCusto.Text = rs.GetString
            rs.Close
            
            If TxtDescontoEntrada.Text = "" Then
                descontoEntrada = 0
            Else
                'Alterando a "," por "." para não dar erro no SQL
                descontoEntrada = Replace(TxtDescontoEntrada.Text, ",", ".")
            End If
            
            
            'SELECT para buscar o valot total do pedido (com o desconto)
            SQL = "SELECT SUM(itens.CustoTotalProduto) - (SUM(itens.CustoTotalProduto) * (" & descontoEntrada & " * 0.01)) " & _
                  "FROM ItensEntrada AS itens " & _
                  "WHERE itens.IdEntrada = " & TxtIdEntrada.Text
            'Adiconando o valor final do pedido
            rs.Open SQL, cn, adOpenStatic
                TxtCustoFinal.Text = rs.GetString
            rs.Close
        'VOLTA O FOCO PARA O TXT COD ITEM
        TxtCodItem.SetFocus
    Else
        'VOLTA O FOCO PARA O TXT COD ITEM
        TxtCodItem.SetFocus
    End If
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro ao tentar recalcular o custo da Entrada: " & TxtCusto.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub

Private Function finalizarEntrada(EntradaIdCli As Integer, EntradaIdPgto As Integer, EntradaCusto As String, EntradaDesconto As String, EntradaCustoTotal As String)
        
    Dim CMD As New ADODB.Command
    
    Dim Retorno As Integer
    
    
    CMD.ActiveConnection = cn
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "finalizarEntrada"
    
    CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adInteger, adParamReturnValue, , -1)
    CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adInteger, adParamOutput, , -1)
    
    CMD.Parameters.Append CMD.CreateParameter("EntradaIdCli", adInteger, adParamInput, , EntradaIdCli)
    CMD.Parameters.Append CMD.CreateParameter("EntradaIdPgto", adInteger, adParamInput, , EntradaIdPgto)
    CMD.Parameters.Append CMD.CreateParameter("EntradaCusto", adInteger, adParamInput, , EntradaCusto)
    CMD.Parameters.Append CMD.CreateParameter("EntradaDesconto", adInteger, adParamInput, , EntradaDesconto)
    CMD.Parameters.Append CMD.CreateParameter("EntradaCustoTotal", adInteger, adParamInput, , EntradaCustoTotal)
    
    'Executa o Command (SP)
    CMD.Execute
    
    'Adiciona o retorno da SP na variavel retorno
    Retorno = CMD.Parameters("RetornoOperacao").Value
    
    ' Se retorno = 0 deu certo | 1 = deu errado
    finalizarEntrada = Retorno
End Function

Private Function RecalcularCustoEntrada()
    
    Dim Custo As Double
    Dim Desconto As Double
    Dim CustoFinal As Double
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim descontoEntrada As Double
    
    SQL = "SELECT MaxDescontoEntrada FROM ConfiguracoesGerais"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            descontoEntrada = -1
        Else
            descontoEntrada = Format(rs("MaxDescontoEntrada"), "0.00")
        End If
    rs.Close
    
    If TxtCusto.Text = "" Or TxtCusto.Text = Null Then
        TxtCusto.Text = 0
    End If
    
    If TxtDescontoEntrada.Text = "" Then
        Desconto = 0
    Else
        Desconto = TxtDescontoEntrada.Text
    End If
    
    TxtDescontoEntrada.Text = Format(TxtDescontoEntrada.Text, "0.00")
    
    If descontoEntrada = -1 Then
        If Desconto = 100 Then
            MsgBox "Não é permitido dar uma desconto de 100%"
            TxtDescontoEntrada.Text = 0
            TxtDescontoEntrada.SetFocus
        ElseIf Desconto > 99.994 Then
            MsgBox "Valor de desconto inválido!"
            TxtDescontoEntrada.Text = 0
            TxtDescontoEntrada.SetFocus
        End If
    Else
        If Desconto > descontoEntrada Then
            MsgBox ("Desconto informado maior que " & descontoEntrada & " permitido!")
            TxtDescontoEntrada.Text = 0
            TxtDescontoEntrada.SetFocus
        Else
            If TxtDescontoEntrada.Text = "" Then
                Desconto = 0
            Else
                Desconto = TxtDescontoEntrada.Text
            End If
        End If
    End If

    'Faz o calculo do desconto
    CustoFinal = (Custo - (Custo * (Desconto * 0.01)))
    
    'MOSTAR O CUSTO NA TELA
    TxtCustoFinal.Text = CustoFinal
End Function

Private Function limpaCamposItem()

    TxtCodItem.Text = ""
    TxtNomeItem.Text = ""
    TxtCustoItem.Text = ""
    TxtItemQuantidade.Text = ""
    TxtDescontoItem.Text = ""
    TxtCustoFinalItem.Text = ""
    LblEstoqueItem.Caption = ""
    LblQuantiaEntrada.Caption = ""
    LblEstoqueFinal.Caption = ""
    LinhaSomaEstoque.Visible = False

End Function

Private Function preencheEntrada(ID As Integer)
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim rsDados As New ADODB.Recordset
    Dim tipoOperacao As Integer
    
    ' SE STATUS = 0 ENTÃO PEDIDO NOVO
    If STATUS = 0 Then
        'HABILITO TUDO PARA QUE SEJA ALTERADO PARA DAR UMA NOVO PEDIDO
        TxtIdEntrada.Enabled = False
        TxtFornecedor.Enabled = True
        TxtFormaPgto.Enabled = True
        ' SER FOR ADMIN PODE ALTERAR O VALOR DO PEDIDO MANUALMENTE
        If FormLogin.ADMIN = 1 Then
            TxtCusto.Enabled = True
            TxtCustoItem.Enabled = True
            ' SE FOR ADMIN NÃO PRECISA DO BOTÃO DE LIBERAÇÃO
            BtnLiberarAlteracao.Visible = False
        ElseIf FormLogin.ADMIN = 2 Then
            TxtCusto.Enabled = False
            TxtCustoItem.Enabled = False
            ' SE NÃO FOR ADMIN, PODE SOLICITAR LIBERAÇÃO
            BtnLiberarAlteracao.Visible = True
        End If
        TxtDescontoEntrada.Enabled = True
        TxtCustoFinal.Enabled = False
        BtnAdicionarItem.Enabled = True
        BtnCancelar.Enabled = True
        BtnNovo.Enabled = False
        BtnRecalcularCusto.Enabled = True
        BtnRemoverItem.Enabled = True
        BtnAvancaRegistro.Enabled = False
        BtnVoltaRegistro.Enabled = False
        
        ' DESABILITA O MENU
        MDIFormInicio.Menu.Enabled = False
        MDIFormInicio.Relatorio.Enabled = False
        MDIFormInicio.Configurações.Enabled = False
        
        ' LIMPA TODOS OS CAMPOS
        TxtFornecedor.Text = ""
        TxtFormaPgto.Text = ""
        TxtCusto.Text = ""
        TxtDescontoEntrada.Text = ""
        TxtCustoFinal.Text = ""
        LblFornecedorNome.Caption = ""
        LblFormaPgto.Caption = ""
        ListViewItensEntrada.ListItems.Clear
        limpaCamposItem
        
        ' MUDA O NOME DO CAMPO IMPRIMIR PARA FINALIZAR
        BtnFinalizar.Caption = "Finalizar"
        
        ' DEFINO O FOCO PARA O TXT CLIENTE
        'SendKeys ("{tab}")
    ' SE STATUS = 1 ENTÃO PEDIDO JÁ FINALIZADO
    ElseIf STATUS = 1 Then
        SQL = "SELECT * " & _
              "FROM Entrada " & _
              "WHERE IdEntrada = " & ID
        
            rsDados.Open SQL, cn, adOpenStatic
                TxtIdEntrada.Text = rsDados("IdEntrada")
                TxtFornecedor.Text = rsDados("EntradaIdCli")
                TxtFormaPgto.Text = rsDados("EntradaIdFormaPgto")
                TxtCusto.Text = rsDados("EntradaCusto")
                TxtDescontoEntrada.Text = rsDados("EntradaDesconto")
                TxtCustoFinal.Text = rsDados("EntradaCustoTotal")
            rsDados.Close
        Set rsDados = Nothing
        
        ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
        STATUS = 1
         
        'DESABILITO TUDO PARA QUE NÃO SEJA ALTERADO EM PEDIDOS NO MODO DE VISUALIZAÇÃO
        TxtIdEntrada.Enabled = True
        TxtFornecedor.Enabled = False
        TxtFormaPgto.Enabled = False
        TxtCusto.Enabled = False
        TxtDescontoEntrada.Enabled = False
        TxtCustoFinal.Enabled = False
        BtnAdicionarItem.Enabled = False
        BtnNovo.Enabled = True
        BtnCancelar.Enabled = False
        BtnRecalcularCusto.Enabled = False
        BtnRemoverItem.Enabled = False
        BtnLiberarAlteracao.Visible = False
        BtnAvancaRegistro.Enabled = True
        BtnVoltaRegistro.Enabled = True
        
        ' HABILITA O MENU
        MDIFormInicio.Menu.Enabled = True
        MDIFormInicio.Relatorio.Enabled = True
        MDIFormInicio.Configurações.Enabled = True
        
        ' FORMATANDO OS CAMPOS
        TxtCusto.Text = Format(TxtCusto.Text, "0.00")
        TxtDescontoEntrada.Text = Format(TxtDescontoEntrada.Text, "0.00")
        
        ' MUDA O NOME DO CAMPO IMPRIMIR PARA FINALIZAR
        BtnFinalizar.Caption = "Imprimir"
        
        'Chama a função de atualizar a lista de itens
        atualizaListaEntrada
    
        'CHAMA A FUNÇÃO DE ADICIONAR O NOME DO FORNECEDOR NA TELA
        PegaNomeFornecedor
        
        ' CHAMA A FUNÇÃO DE ADICIONAR O NOME DA FORMA DE PAGAMENTO NA TELA
        PegaNomeFormaPgto
    End If
End Function

Public Function imprimirEntrada()

    Dim SQL As String
    Dim SQL2 As String
    Dim rs As New ADODB.Recordset
    
    ' SQL com o que vai buscas no banco
    SQL = "SELECT ent.IdEntrada AS Entrada, cli.CliNome AS Cliente, pgto.NomeFormaPgt AS FormaPGTO, ent.EntradaCusto AS Custo, ent.EntradaDesconto AS Desconto, ent.EntradaCustoTotal AS CustoTotal " & _
          "FROM Entrada AS ent " & _
          "JOIN Clientes AS cli ON ent.EntradaIdCli = cli.IdCliente " & _
          "JOIN FormaPgto AS pgto ON ent.EntradaIdFormaPgto = pgto.IdFormaPgt " & _
          "WHERE ent.IdEntrada = " & TxtIdEntrada.Text

    ' DEFINE O CABEÇALHO DO RELATÓRIO (IMPRSSÃO DE ENTRADA)
    rs.Open SQL, cn, adOpenStatic
        ImpressaoDeEntrada.Entrada = rs("Entrada")
        ImpressaoDeEntrada.Fornecedor = rs("Cliente")
        ImpressaoDeEntrada.FormaPGTO = rs("FormaPGTO")
        ImpressaoDeEntrada.Custo = rs("Custo")
        ImpressaoDeEntrada.Desconto = rs("Desconto")
        ImpressaoDeEntrada.CustoTotal = rs("CustoTotal")
    rs.Close

    ' DEFINE A CONEXÃO COM O BANCO
    ImpressaoDeEntrada.DataControlImpressaoEntrada.ConnectionString = cn
    
    SQL2 = "SELECT itens.IdProduto AS Produto, produto.NomeProduto AS Nome, itens.CustoProduto AS CustoProduto, itens.QuantidadeProduto AS QuantidadeProduto, itens.DescontoProduto AS DescontoProduto, itens.CustoTotalProduto AS CustoTotalProduto " & _
           "FROM ItensEntrada AS itens " & _
           "JOIN Produtos AS produto ON itens.IdProduto = produto.CodProduto " & _
           "WHERE itens.IdEntrada = " & TxtIdEntrada.Text & " " & _
           "ORDER BY itens.IdItensEntrada"
          
    ' PASSA A STRING QUE VAI SER EXECUTADA NO RELATÓRIO
    ImpressaoDeEntrada.DataControlImpressaoEntrada.Source = SQL2
    
    ImpressaoDeEntrada.Show

End Function

Public Function PegaNomeFornecedor()
    
    'FAZ UMA VERIFICAÇÃO, SE FOR EM UM NOVO PEDIDO, VALIDA SE O CLIENTE ESTÁ INATIVO OU EXISTE
    '0 - NOVO PEDIDO
    '1 - PEDIDO JÁ FINALIZADO
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim Retorno As Integer
    
    Retorno = 1
    
    SQL = "SELECT CliNome, CliStatus FROM Clientes WHERE IdCliente = " & TxtFornecedor.Text
    
    rs.Open SQL, cn, adOpenStatic
        If STATUS = 0 Then
            
            ' FUNÇÃO QUE VERIFICA SE O FORNECEDOR EXISTE
            If verificaFornecedor(TxtFornecedor.Text) = 0 Then
                ' SE NÃO EXISTE (0), SAI DA FUNÇÃO
                Retorno = 0
                Exit Function
            End If
            
            ' FUNÇÃO QUE VERIFICA SE O TIPO DO CLIENTE É UM FORNECEDOR
            If verificaTipoFornecedor(TxtFornecedor.Text) = 0 Then
                ' SE CLIENTE NÃO TIVER O TIPO FORNECEDOR NO CADASTRO (0), SAI DA FUNÇÃO
                Retorno = 0
                Exit Function
            End If
            
            If rs("CliStatus") = False Then
                LblFornecedorNome.Caption = rs("CliNome")
                SendKeys "{tab}" ' Set the focus to the next control.
            ElseIf rs("CliStatus") = True Then
                MsgBox "Fornecedor Inativo!"
                LblFornecedorNome.Caption = ""
                TxtFornecedor.Text = ""
                Retorno = 0
            End If
        ElseIf STATUS = 1 Then
            If rs.EOF = False Then
                LblFornecedorNome.Caption = rs("CliNome")
            End If
        End If
    rs.Close
    
    Set rs = Nothing
    
    PegaNomeFornecedor = Retorno
    
End Function

Private Function verificaFornecedor(ID As Integer)

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim Retorno As Integer
    
    SQL = "SELECT CliNome FROM Clientes WHERE IdCliente = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            MsgBox "Fornecedor não cadastrado"
            LblFornecedorNome.Caption = ""
            TxtFornecedor.Text = ""
            TxtFornecedor.SetFocus
            
            Retorno = 0
        Else
            Retorno = 1
        End If
    rs.Close
    Set rs = Nothing
    
    verificaFornecedor = Retorno

End Function


Private Function verificaTipoFornecedor(ID As Integer)

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim Retorno As Integer
    
    SQL = "SELECT TipoFornecedor FROM TipoClientes WHERE IdCliente = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs("TipoFornecedor") = True Then
            Retorno = 1
        Else
            MsgBox ("Tipo cliente inválido para essa operação!")
            LblFornecedorNome.Caption = ""
            TxtFornecedor.Text = ""
            TxtFornecedor.SetFocus
            
            Retorno = 0
        End If
    rs.Close
    Set rs = Nothing
    
    verificaTipoFornecedor = Retorno

End Function


Public Function PegaNomeFormaPgto()
    
    '0 - NOVO PEDIDO
    '1 - PEDIDO JÁ FINALIZADO
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim nomeFormaPGTO As String
    
    SQL = "SELECT NomeFormaPgt, StatusFormaPgt FROM FormaPgto WHERE IdFormaPgt = " & TxtFormaPgto
    
    rs.Open SQL, cn, adOpenStatic
    
        If STATUS = 0 Then
            If TxtFormaPgto.Text = "" Then
                MsgBox "Campo Condição de pagamento não pode ser vazio"
            Else
                If rs.EOF = False Then
                    nomeFormaPGTO = rs("NomeFormaPgt")
                    If rs("StatusFormaPgt") = True Then
                        MsgBox "Forma de Pagamento: " & TxtFormaPgto.Text & " - " & nomeFormaPGTO & " - Inativa!"
                        LblFormaPgto.Caption = ""
                        TxtFormaPgto.Text = ""
                        TxtFormaPgto.SetFocus
                    Else
                        LblFormaPgto.Caption = nomeFormaPGTO
                        SendKeys "{tab}" ' Set the focus to the next control.
                    End If
                Else
                    MsgBox "Forma de Pagamento: " & TxtFormaPgto.Text & " - não cadastrado!"
                    LblFormaPgto = ""
                    TxtFormaPgto = ""
                End If
            End If
        ElseIf STATUS = 1 Then
            nomeFormaPGTO = rs("NomeFormaPgt")
            If rs.EOF = False Then
                LblFormaPgto.Caption = nomeFormaPGTO
            End If
        End If
    rs.Close
    Set rs = Nothing
End Function

Public Function buscaEntrada(ID As Integer)
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    
    SQL = "SELECT * FROM Entrada WHERE IdEntrada = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            
        Else
            MsgBox "Entrada N° " & ID & " não encontrada!"
            
            SQL = "DECLARE @maxID INT " & _
                  "SELECT @maxID = MAX(IdEntrada) FROM Entrada " & _
                  "SELECT @maxID AS Entrada"
          
            rsDados.Open SQL, cn, adOpenStatic
                ID = rsDados("Entrada")
            rsDados.Close
        End If
    rs.Close
    
    ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
    STATUS = 1
    
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preencheEntrada (ID)
    
    SendKeys "{tab}"   ' Set the focus to the next control.
    ' VOLTA O FOCO PARA O TXT PEDIDO
    SendKeys "+{tab}" ' SHIFT TAB
    
End Function

Private Function finalizaForm()

    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    Dim SQL As String
    Dim validacao As Integer
    
    If STATUS = 0 Then ' PEDIDO NOVO
        validacao = MsgBox("Caso saia do cadastro, os dados serão perdidos! Deseja mesmo sair?", vbYesNo)
        
        If validacao = vbYes Then
            'Limpa os itens dA ENTRADA
            SQL = "DELETE FROM ItensEntrada WHERE IdEntrada = " & TxtIdEntrada.Text
            cn.Execute SQL
            
            ' VOLTA O STATUS PARA PEDIDO JÁ FINALIZADO
            STATUS = 1
            
            ' SQL PEGANDO O ULTIMO PEDIDO
            SQL = "DECLARE @maxID INT " & _
                  "SELECT @maxID = MAX(IdEntrada) FROM Entrada " & _
                  "SELECT @maxID AS Entrada"
            
            ' EXECUTA E ADICIONA O ULTIMA ENTRADA NA TELA
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("Entrada")) Then
                    STATUS = 0
                    ID = 1
                    TxtIdEntrada.Text = 1
                Else
                    ID = rs("Entrada")
                    ' SETO O STATUS DA TELA PARA 1 -> REGISTRO JÁ FINALIZADO
                    STATUS = 1
                End If
            rs.Close
            ' LIMPA O RECORDSET
            Set rs = Nothing
            
            ' CHAMA A FUNÇÃO PARA ADICIONAR OS DADOS NA TELA
            preencheEntrada (ID)
        Else
            TxtFornecedor.SetFocus
        End If
    Else ' REGISTRO JÁ FINALIZADOS
        Unload Me
    End If
    
End Function

Public Function CustoProduto()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT CustoProduto FROM Produtos WHERE CodProduto = " & TxtCodItem.Text
    
    If TxtCodItem.Text = 0 Or TxtCodItem.Text = "" Then
        'DEIXA O CAMPO EM BRANCO
        TxtCustoItem.Text = ""
    Else
        rs.Open SQL, cn, adOpenStatic
            'ADICIONA O VALOR DO PRODUTO NO CAMPO
            TxtCustoItem.Text = Format(rs("CustoProduto"), "0.00")
        rs.Close
    End If
    
End Function

Private Function atualizaCustoItem()
    
    Dim Custo As Double
    Dim Quantidade As Integer
    Dim Desconto As Double
    Dim CustoTotal As Double
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim descontoItem As Integer
    Dim EstoqueItem As Long
       
    SQL = "SELECT MaxDescontoItemEntrada FROM ConfiguracoesGerais"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            descontoItem = -1
        Else
            descontoItem = Format(rs("MaxDescontoItemEntrada"), "0.00")
        End If
    rs.Close
    
    ' PEGA O CUSTO DO PRODUTO
    If TxtCustoItem.Text = "" Then
        TxtCustoItem.Text = 0
    Else
        Custo = TxtCustoItem.Text
    End If

    ' VERIFICA CAMPOS VAZIOS
    If TxtItemQuantidade.Text = "" Then
        TxtItemQuantidade.Text = 1
    Else
        Quantidade = TxtItemQuantidade.Text
    End If
    
    If TxtDescontoItem = "" Then
        TxtDescontoItem.Text = 0
        Desconto = 0
    Else
        Desconto = TxtDescontoItem.Text
    End If

    ' SE DESCONTO ITEM = -1 UTILIZA O DESCONTO PADRAO DO SISTEMA MAX = 99,994
    If descontoItem = -1 Then
        If Desconto = 100 Then
            MsgBox ("Não é permitido o desconto de 100%")
            TxtDescontoItem.Text = ""
            Desconto = 0
            TxtDescontoItem.SetFocus
        ElseIf Desconto > 99.994 Then
            MsgBox ("Desconto inválido")
            TxtDescontoItem.Text = ""
            Desconto = 0
            TxtDescontoItem.SetFocus
        End If
    Else
        If Desconto > descontoItem Then
            MsgBox "Desconto informado maior que os " & descontoItem & " permitidos!"
            TxtDescontoItem = ""
            Desconto = 0
            TxtDescontoItem.SetFocus
        End If
    End If
    
    ' CHAMA A FUNÇÃO PARA MOSTRAR O ESTOQUE E FAZ O CALCULO DO NOVO ESTOQUE
    EstoqueItem = SPsGlobais.VerificaEstoqueProduto(TxtCodItem.Text)
    
    ' ADICIONA O ESTOQUE DO PRODUTO NO CAMPO
    LblEstoqueItem.Caption = EstoqueItem
    
    ' ADICIONA A QUANTIDADE O CAMPO DO ESTOQUE
    LblQuantiaEntrada.Caption = "+" & TxtItemQuantidade.Text
    
    ' ADICIONA O NOVO ESTOQUE
    LblEstoqueFinal.Caption = EstoqueItem + TxtItemQuantidade.Text
    
    ' EXECUTA
    cn.Execute (SQL)
    
    'CALCULA O VALOR COM BASE NOS PARAMETROS
    CustoTotal = ((Custo * Quantidade) - ((Custo * Quantidade) * (Desconto * 0.01)))
    
    'ADICIONA O VALOR CALCULADO NA TELA
    TxtCustoFinalItem.Text = Format(CustoTotal, "0.00")
End Function

Private Function atualizaListaEntrada()
    
    Dim rsDados As New ADODB.Recordset
    Dim SQL As String
    
    'SEMPRE LIMPA O LISTVIEW ANTES DE INSERIR NOVAMENTE
    ListViewItensEntrada.ListItems.Clear
    
    'PEGA OS DADOS INICIAIS DA TABLE E ADICIONA NO GRID
    SQL = "SELECT itens.IdItensEntrada AS CodItensPedido, itens.IdProduto AS Produto, produto.NomeProduto AS Nome, itens.CustoProduto AS Custo, itens.QuantidadeProduto AS Quantidade, itens.DescontoProduto AS Desconto, itens.CustoTotalProduto AS CustoTotal " & _
          "FROM ItensEntrada AS itens " & _
          "JOIN Produtos AS produto ON itens.IdProduto = produto.CodProduto " & _
          "WHERE itens.IdEntrada = " & TxtIdEntrada.Text & " " & _
          "ORDER BY itens.IdItensEntrada"
          
    rsDados.Open SQL, cn, adOpenDynamic
        Do While rsDados.EOF = False
            Set itens = ListViewItensEntrada.ListItems.Add(, , rsDados("CodItensPedido"))
            itens.SubItems(1) = rsDados("Produto")
            itens.SubItems(2) = rsDados("Nome")
            itens.SubItems(3) = rsDados("Custo")
            itens.SubItems(4) = rsDados("Quantidade")
            itens.SubItems(5) = rsDados("Desconto")
            itens.SubItems(6) = rsDados("CustoTotal")
        
            'Move Para o próximo registro
            rsDados.MoveNext
        Loop
    rsDados.Close
End Function

Private Sub TxtCustoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtCustoFinal_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtCustoItem_GotFocus()

    With TxtCustoItem
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtCustoItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtCustoItem_KeyPress(KeyAscii As Integer)

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If

    If KeyAscii = 13 Then ' ENTER
        SendKeys ("{tab}")
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtCustoItem_LostFocus()

    'Formata o número para apenas 2 casas decimais
    TxtCustoItem.Text = Format(TxtCustoItem.Text, "0.00")

    If TxtCodItem.Text = "" Then
        ' NÃO ATUALIZA O VALOR DO ITEM
    Else
        ' CHAMA A FUNÇÃO QUE ATUALIZA O CUSTO DO ITEM
        atualizaCustoItem
    End If

End Sub

Private Function verificaItensEntrada(ID As Integer)

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim Retorno As Integer
    Dim custoItem, quantidadeItem, descontoItem, custoFinalItem As String
    
    SQL = "SELECT IdProduto FROM ItensEntrada WHERE IdEntrada = " & TxtIdEntrada.Text & " AND IdProduto = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            custoItem = Replace(TxtCustoItem.Text, ",", ".")
            quantidadeItem = Replace(TxtItemQuantidade.Text, ",", ".")
            descontoItem = Replace(TxtDescontoItem.Text, ",", ".")
            custoFinalItem = Replace(TxtCustoFinalItem.Text, ",", ".")
            
            SQL = "UPDATE ItensEntrada " & _
                  "SET CustoProduto = " & custoItem & ", QuantidadeProduto = " & quantidadeItem & ", DescontoProduto = " & descontoItem & ", CustoTotalProduto = " & custoFinalItem & " " & _
                  "WHERE IdEntrada = " & TxtIdEntrada.Text & " AND IdProduto = " & ID
                  
            Debug.Print SQL
            cn.Execute (SQL)
            
            Retorno = 0
        Else
            Retorno = 1
        End If
    rs.Close
    Set rs = Nothing
    
    verificaItensEntrada = Retorno

End Function

Private Function atualizaCustoProduto()

    Dim CMD As New ADODB.Command
    
    CMD.ActiveConnection = cn
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "AtualizarCustoProduto"
    
    CMD.Parameters.Append CMD.CreateParameter("ID", adInteger, adParamInput, , TxtIdEntrada.Text)
    
    CMD.Execute

End Function
