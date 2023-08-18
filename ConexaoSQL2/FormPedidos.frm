VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos"
   ClientHeight    =   6870
   ClientLeft      =   2610
   ClientTop       =   4650
   ClientWidth     =   17025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   17025
   Begin VB.CommandButton BtnVoltaRegistro 
      Caption         =   "<"
      Height          =   555
      Left            =   4110
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   405
      Width           =   675
   End
   Begin VB.CommandButton BtnAvancaRegistro 
      Caption         =   ">"
      Height          =   555
      Left            =   4875
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   405
      Width           =   675
   End
   Begin VB.CommandButton BtnLiberarAlteracao 
      Caption         =   "*"
      Height          =   405
      Left            =   15780
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Solicitar Desbloqueio para operador ADMIN"
      Top             =   810
      Width           =   510
   End
   Begin MSComctlLib.ListView ListViewItensPedido 
      Height          =   3510
      Left            =   240
      TabIndex        =   33
      Top             =   2985
      Width           =   8865
      _ExtentX        =   15637
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
   Begin VB.CommandButton BtnFormaPgto 
      Caption         =   "->"
      Height          =   405
      Left            =   1965
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2115
      Width           =   500
   End
   Begin VB.CommandButton BtnCliente 
      Caption         =   "->"
      Height          =   405
      Left            =   1965
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1335
      Width           =   500
   End
   Begin VB.CommandButton btnPedido 
      Caption         =   "->"
      Height          =   405
      Left            =   1980
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   465
      Width           =   500
   End
   Begin VB.CommandButton BtnRecalcularValor 
      Caption         =   "Recalcular Valor"
      Height          =   315
      Left            =   13755
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   180
      Width           =   1785
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
      Left            =   2670
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2085
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   5880
      TabIndex        =   22
      Top             =   165
      Width           =   6525
      Begin VB.CommandButton BtnNovo 
         Caption         =   "Novo"
         Height          =   585
         Left            =   270
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   195
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
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "Cancelar"
         Height          =   585
         Left            =   2325
         TabIndex        =   12
         Top             =   180
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produto"
      Height          =   3540
      Left            =   9225
      TabIndex        =   21
      Top             =   2970
      Width           =   7545
      Begin VB.Frame Frame3 
         Caption         =   "Estoque"
         Height          =   1635
         Left            =   5730
         TabIndex        =   41
         Top             =   945
         Width           =   1620
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
            Left            =   165
            TabIndex        =   47
            Top             =   570
            Width           =   1275
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
            Left            =   165
            TabIndex        =   46
            Top             =   1080
            Width           =   1275
         End
         Begin VB.Line LinhaSomaEstoque 
            X1              =   225
            X2              =   1455
            Y1              =   1005
            Y2              =   1005
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
            Left            =   165
            TabIndex        =   42
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.TextBox TxtValorItem 
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
      Begin VB.CommandButton BtnPesquisarProduto 
         Caption         =   "->"
         Height          =   405
         Left            =   2670
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   405
         Width           =   500
      End
      Begin VB.TextBox TxtValorFinalItem 
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1700
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
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   375
         Width           =   4095
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
      Begin VB.CommandButton BtnRemoverItem 
         Caption         =   "Remover"
         Height          =   630
         Left            =   5925
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   2775
         Width           =   1245
      End
      Begin VB.CommandButton BtnAdicionarItem 
         Caption         =   "Adicionar"
         Height          =   630
         Left            =   4575
         TabIndex        =   8
         Top             =   2790
         Width           =   1245
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
         Left            =   3990
         TabIndex        =   45
         Top             =   1935
         Width           =   285
      End
      Begin VB.Label Label14 
         Caption         =   "Valor"
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
         Left            =   390
         TabIndex        =   37
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label13 
         Caption         =   "Total"
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
         Left            =   1905
         TabIndex        =   35
         Top             =   2505
         Width           =   720
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
         TabIndex        =   28
         Top             =   1425
         Width           =   1245
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
         TabIndex        =   24
         Top             =   1965
         Width           =   1080
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
         TabIndex        =   23
         Top             =   465
         Width           =   1035
      End
   End
   Begin VB.TextBox TxtValorFinal 
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
      Left            =   13845
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2145
      Width           =   1700
   End
   Begin VB.TextBox TxtDescontoPedido 
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
      Left            =   13845
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1455
      Width           =   1035
   End
   Begin VB.TextBox TxtValor 
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
      Left            =   13845
      MaxLength       =   12
      TabIndex        =   9
      Top             =   750
      Width           =   1700
   End
   Begin VB.TextBox TxtCliente 
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
      Left            =   2670
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1305
      Width           =   1200
   End
   Begin VB.TextBox TxtNumPedido 
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
      Left            =   2685
      MaxLength       =   12
      TabIndex        =   1
      Top             =   435
      Width           =   1200
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
      Height          =   345
      Left            =   14985
      TabIndex        =   48
      Top             =   1500
      Width           =   285
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
      Left            =   4530
      TabIndex        =   27
      Top             =   2145
      Width           =   6645
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
      Left            =   3900
      TabIndex        =   26
      Top             =   2115
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
      Left            =   200
      TabIndex        =   25
      Top             =   2070
      Width           =   1635
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
      Left            =   3900
      TabIndex        =   20
      Top             =   1290
      Width           =   600
   End
   Begin VB.Label LblCliNome 
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
      Left            =   4530
      TabIndex        =   19
      Top             =   1320
      Width           =   6645
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Total:"
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
      Left            =   11730
      TabIndex        =   18
      Top             =   2130
      Width           =   2040
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
      Left            =   11805
      TabIndex        =   17
      Top             =   1425
      Width           =   1890
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor:"
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
      Left            =   12495
      TabIndex        =   16
      Top             =   765
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente"
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
      Left            =   195
      TabIndex        =   15
      Top             =   1290
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Pedido"
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
      Left            =   210
      TabIndex        =   13
      Top             =   420
      Width           =   1635
   End
End
Attribute VB_Name = "FormPedidos"
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
    Dim codPedido As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtNumPedido.Text & " " & _
          "SELECT TOP 1 IdPedidos FROM Pedido WHERE IdPedidos > @ID ORDER BY IdPedidos"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codPedido = rs("IdPedidos")
        Else
            codPedido = TxtNumPedido.Text
        End If
    rs.Close
    
    TxtNumPedido.Text = codPedido
    preenchePedido (codPedido)

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

Private Sub BtnCliente_Click()
    FormBuscaCliente.FORMULARIO = "FormPedidos"
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
    FormBuscaFormaPGTO.FORMULARIO = "FormPedidos"
    FormBuscaFormaPGTO.Show
    FormBuscaFormaPGTO.SetFocus
    MDIFormInicio.Enabled = False
End Sub

Private Sub BtnLiberarAlteracao_Click()
    
    MDIFormInicio.Enabled = False
    FormSolicitaAcesso.FORMULARIO = "FormPedidos"
    FormSolicitaAcesso.Show
    FormSolicitaAcesso.TxtLogin.SetFocus
    
End Sub

Private Sub BtnNovo_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim numeroPedido As Integer
    
    SQL = "SELECT MAX(IdPedidos)+1 AS Pedido FROM Pedido"
    
    rs.Open SQL, cn, adOpenStatic
        numeroPedido = rs("Pedido")
        TxtNumPedido.Text = rs("Pedido")
    rs.Close
    
    ' SETO O STATUS DA TELA PARA 0 -> PEDIDO NOVO
    STATUS = 0
    preenchePedido (numeroPedido)
    
End Sub

Private Sub BtnNovo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnPesquisarProduto_Click()

    FormBuscaProduto.FORMULARIO = "FormPedidos"
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
    Dim codPedido As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtNumPedido.Text & " " & _
          "SELECT TOP 1 IdPedidos FROM Pedido WHERE IdPedidos < @ID ORDER BY IdPedidos DESC"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codPedido = rs("IdPedidos")
        Else
            codPedido = TxtNumPedido.Text
        End If
    rs.Close
    
    TxtNumPedido.Text = codPedido
    preenchePedido (codPedido)

End Sub

Private Sub BtnVoltaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub Form_Load()
    
    Dim rs As New ADODB.Recordset
    Dim numeroPedido As Integer
    Dim SQL As String
    
    On Error GoTo TrataErro
    
    SQL = "DECLARE @maxPedido INT " & _
          "SELECT @maxPedido = MAX(IdPedidos) FROM Pedido " & _
          "SELECT @maxPedido AS Pedido"
             
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Pedido")) Then
            STATUS = 0
            numeroPedido = 1
            TxtNumPedido.Text = 1
        Else
            ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
            STATUS = 1
            numeroPedido = rs("Pedido")
        End If
    rs.Close
    ' LIMPA O RECORDSET
    Set rs = Nothing
    
    ' LIMPA OS CAMPOS DO ITEM
    limpaCamposItem
    
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preenchePedido (numeroPedido)
    
Exit Sub
TrataErro:
    MsgBox "Algum erro ocorreu ao carregar o Form - " & FormPedidos.Name & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub btnPedido_Click()
    FormBuscaPedido.Show
    FormBuscaPedido.SetFocus
    MDIFormInicio.Enabled = False
End Sub

Private Sub BtnAdicionarItem_Click()

    'On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim SQLVerificaPedido As String
    Dim SQLVerificarProduto As String
    Dim rsVerificarProduto As New ADODB.Recordset
    Dim descontoItem As String
    Dim VendaItemNegativo As Integer
    Dim EstoqueFinal As Long
    Dim ValorPedido As Double
    Dim valorItem As String
       
    SQL = "SELECT * FROM ConfiguracoesGerais"
    
    ' PEGA AS CONFIGURAÇÕES GERAIS
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            VendaItemNegativo = -1
        Else
            VendaItemNegativo = rs("VendaItemNegativo")
        End If
    rs.Close
    Set rs = Nothing
    
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
        valorItem = Replace(TxtValorItem.Text, ",", ".")
        
        SQLVerificaPedido = "SELECT IdPedidos FROM Pedido"
        
        rs.Open SQLVerificaPedido, cn, adOpenStatic
            If rs.EOF = True Then
                'ADICIONA O PRODUTO NA TABELA DE PRODUTOS PARA O PEDIDO 1 (PRIMEIRO PEDIDO)
                SQL = "DECLARE @ValorTotal decimal(18,2) " & _
                      "SELECT @ValorTotal = ((" & valorItem & " * " & TxtItemQuantidade.Text & ") - ((" & valorItem & " * " & TxtItemQuantidade.Text & ") * (" & descontoItem & " * 0.01))) " & _
                      "INSERT INTO ItensPedido (CodPedido, CodProduto, ValorProduto, QuantidadeProduto, DescontoProduto, ValorTotalProduto) " & _
                      "VALUES (1, " & TxtCodItem.Text & ", " & valorItem & ", " & TxtItemQuantidade.Text & ", " & descontoItem & ", @ValorTotal)"
            Else
                'ADICIONA O PRODUTO NA TABELA DE PRODUTOS PARA O PEDIDO (NORMALMENTE)
                SQL = "INSERT INTO ItensPedido (CodPedido, CodProduto, ValorProduto, QuantidadeProduto, DescontoProduto, ValorTotalProduto) " & _
                      "VALUES (" & TxtNumPedido.Text & ", " & TxtCodItem.Text & ", " & valorItem & ", " & TxtItemQuantidade.Text & ", " & descontoItem & ", ((" & valorItem & " * " & TxtItemQuantidade.Text & ") - ((" & valorItem & " * " & TxtItemQuantidade.Text & ") * (" & descontoItem & " * 0.01))))"

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
                EstoqueFinal = LblEstoqueItem.Caption - TxtItemQuantidade.Text
                
                ' VERIFICA SE O ESTOQUE É NEGATIVO E SE PODE SER VENDIDO NEGATIVO
                If EstoqueFinal >= 0 Then
                    'INSERE O PRODUTO NO BANCO
                    cn.Execute SQL
                    
                    'Mensagem de conclusão
                    MsgBox "Item incluído com sucesso"
                ElseIf EstoqueFinal < 0 Then
                    If VendaItemNegativo = 0 Then
                        'INSERE O PRODUTO NO BANCO
                        cn.Execute SQL

                        'Mensagem de conclusão
                        MsgBox "Item incluído com sucesso"
                        'Volta o foco para o codItem
                    Else
                        MsgBox ("Estoque insuficiente! Operação não cancelada...")
                    End If
                End If
            Else
                'ERRO CASO NÃO TENHA O PRODUTO CADASTRADO NO SISTEMA
                MsgBox "Produto não cadastrado", vbCritical
            End If
        rsVerificarProduto.Close
        
        'PEGA OS DADOS DA TABLE E ADICIONA NO GRID
        atualizaListaPedido
                
        'Zerando os valores ao inserir o produto
        limpaCamposItem
                    
        'SELECT para buscar a soma total dos valores do pedido, SE NÃO TIVER NENHUM VALOR RETORNA 0
        SQL = "DECLARE @ValorTotal DECIMAL(18,2) " & _
              "SELECT @ValorTotal = SUM(itens.ValorTotalProduto) " & _
              "FROM ItensPedido AS itens " & _
              "WHERE itens.CodPedido = " & TxtNumPedido.Text & " " & _
              "SELECT @ValorTotal = ISNULL(@ValorTotal, 0) " & _
              "SELECT @ValorTotal AS ValorTotal"
        'Adicionando o valorTotal no TextField
        
        rs.Open SQL, cn, adOpenStatic
            TxtValor.Text = Format(rs("ValorTotal"), "0.00")
        rs.Close
        
        'Verificando se o Desconto de Pedido está nulo, se sim coloca valor 0
        If TxtDescontoPedido = "" Then
            TxtDescontoPedido.Text = 0
        End If
        
        'Volta o foco para o codItem
        TxtCodItem.SetFocus
    End If
    
    'SEMPRE ATUALIZA O VALOR DO PEDIDO
    RecalcularValorPedido
Exit Sub

TrataErro:
    MsgBox " AÇÃO CANCELADA - Ocorreu um erro durante a tentativa de inserir o item: " & TxtCodItem & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub

Private Sub BtnRemoverItem_Click()
    
    On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim X As Integer
         
    If MsgBox("Confirmar a exclusão deste Item?", vbYesNo) = vbYes Then
        
        SQL = "DELETE FROM ItensPedido WHERE CodItensPedido = " & ListViewItensPedido.SelectedItem
        
        cn.Execute (SQL)
        
        'ATUALIZA A LISTA DE ITENS
        atualizaListaPedido
        'RECALCULA O VALOR DO PEDIDO
        RecalcularValorPedido
    End If
    
    RecalcularValorPedido
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
        SQL = "DELETE FROM ItensPedido WHERE CodPedido = " & TxtNumPedido.Text
        cn.Execute SQL
    End If
    
    ' HABILITA O MENU
    MDIFormInicio.Menu.Enabled = True
    MDIFormInicio.Relatorio.Enabled = True
    MDIFormInicio.Configurações.Enabled = True
    
End Sub

Private Sub LblValorProduto_Click()

End Sub





Private Sub ListViewItensPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtCliente_GotFocus()

    With TxtCliente
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Debug.Print KeyCode
    
    If KeyCode = 115 Then ' F4
        BtnCliente_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub




Private Sub TxtCliente_KeyPress(KeyAscii As Integer)

    On Error GoTo TrataErro
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER key.
        ' VERIFICA SE O CAMPO NÃO ESTÁ VAZIO
        If TxtCliente.Text = "" Then
            MsgBox ("Necessário informar um valor")
            KeyAscii = 0
        Else
            ' CHAMA A FUNÇÃO QUE ADICIONA O NOME DO CLIENTE NO LABEL
            PegaNomeCliente
            KeyAscii = 0
        End If
    End If
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro durante a tentativa de inserir o Cliente: " & TxtCodItem & vbCrLf & _
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

Private Sub TxtDescontoPedido_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub TxtFormaPgto_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(TxtFormaPgto.Text) Then
        TxtFormaPgto.Text = ""
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
        atualizaValorItem
    End If

End Sub

Private Sub TxtNomeItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtNumPedido_GotFocus()

    With TxtNumPedido
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtNumPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim validacao As Boolean

    If KeyCode = 115 Then ' F4
        btnPedido_Click
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
            MsgBox ("Necessário informar um valor")
            KeyAscii = 0
        Else
            ' CHAMA A FUNÇÃO QUE ADICIONA O NOME DA FORMA DE PAGAMENTO NO LABEL
            PegaNomeFormaPgto
            KeyAscii = 0
        End If
    End If
    
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro durante a tentativa de inserir a Forma de Pagamento: " & TxtCodItem & vbCrLf & _
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
    Dim EstoqueItem As Integer
    
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
                    
                    ' CHAMA A FUNÇÃO PARA MOSTRAR O ESTOQUE
                    EstoqueItem = SPsGlobais.VerificaEstoqueProduto(TxtCodItem.Text)
                    
                    ' ADICIONA O ESTOQUE NO LABEL
                    LblEstoqueItem.Caption = EstoqueItem
                    
                    ' ATIVA A LINHA DA SOMA E ADICIONA O TOTAL DO ESTOQUE
                    LinhaSomaEstoque.Visible = True
                    LblEstoqueFinal.Caption = EstoqueItem - 1
                    
                    ' ADICIONA O VALOR 1 NO TXT DA QUANTIDADE
                    TxtItemQuantidade.Text = 1
                    
                    'CHAMA A FUNÇÃO PARA SETAR O VALOR DO PRODUTO
                    ValorProduto
                    
                    'CHAMA A FUNÇÃO PARA ATUALIZAR O VALOR DO ITEM
                    atualizaValorItem
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

Private Sub TxtCodItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(TxtCodItem.Text) Then
        TxtCodItem.Text = ""
    End If
End Sub

Private Sub TxtNumPedido_KeyPress(KeyAscii As Integer)
        
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
        
    If KeyAscii = 13 Then ' The ENTER key.
        If TxtNumPedido.Text = "" Then
            MsgBox ("Necessário informar um valor")
            KeyAscii = 0
        Else
            If STATUS = 1 Then
                ' FUNÇÃO QUE VERIFICA SE EXISTE ESTE PEDIDO NO SQL
                buscaPedido (TxtNumPedido.Text)
                KeyAscii = 0
            Else
                SendKeys ("{tab}") ' MOVE PARA O PRÓXIMO CAMPO
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TxtValor_GotFocus()

    With TxtValor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtValor_LostFocus()
    
    Dim descontoPedido As String
               
    'Formata o número para apenas 2 casas decimais
    TxtValor.Text = Format(TxtValor.Text, "0.00")
            
    'CHAMA A FUNÇÃO PARA FAZER O CALCULO DO PEDIDO
    RecalcularValorPedido
End Sub

'Formatação <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub TxtValorFinal_Change()
    'Formata o número para apenas 2 casas decimais
    TxtValorFinal.Text = Format(TxtValorFinal.Text, "0.00")
End Sub
'Formatação >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



Private Sub TxtValor_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(TxtValor.Text) Then
        TxtValor.Text = ""
    End If
End Sub

Private Sub TxtDescontoPedido_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(TxtDescontoPedido.Text) Then
        TxtDescontoPedido.Text = ""
    End If
End Sub

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
                atualizaValorItem
                SendKeys "{tab}"   ' Set the focus to the next control.
                KeyAscii = 0       ' Ignore this key.
            End If
        End If
    End If
Exit Sub

TrataError:
    MsgBox " Ocorreu um erro ao informar a quantidade: " & TxtItemQuantidade & vbCrLf & _
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
    MsgBox " Ocorreu um erro ao informar a quantidade: " & TxtItemQuantidade & vbCrLf & _
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
        atualizaValorItem
    End If
    
End Sub

Private Sub TxtDescontoPedido_KeyPress(KeyAscii As Integer)
    
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
    MsgBox " Ocorreu um erro ao informar a o desconto do pedido: " & TxtDescontoPedido & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

'GotFocus para selecionar todo o valor do desconto Total
Private Sub TxtDescontoPedido_GotFocus()
    With TxtDescontoPedido
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtDescontoPedido_LostFocus()
    
    Dim descontoPedido As String
    Dim SQL As String
    Dim rs As New Recordset
    Dim ValorPedido As Double
    
    SQL = "SELECT * FROM ConfiguracoesGerais"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            descontoPedido = -1
        Else
            descontoPedido = Format(rs("MaxDescontoPedido"), "0.00")
        End If
    rs.Close
    
    If TxtValor.Text = "" Then
        MsgBox "Necessário o pedido ter um valor!"
        TxtDescontoPedido.Text = ""
        TxtValor.SetFocus
    Else
        ' SÓ VERIFICA EM NOVOS PEDIDOS
        If STATUS = 0 Then
            If TxtDescontoPedido.Text = "" Then
                TxtDescontoPedido.Text = 0
            End If
            
            ValorPedido = TxtDescontoPedido.Text
            
            If descontoPedido = -1 Then
                If ValorPedido = 100 Then
                    MsgBox "Não é permitido dar uma desconto de 100%"
                    TxtDescontoPedido.Text = 0
                    TxtDescontoPedido.SetFocus
                ElseIf ValorPedido > 99.994 Then
                    MsgBox "Valor de desconto inválido!"
                    TxtDescontoPedido.Text = 0
                    TxtDescontoPedido.SetFocus
                End If
            Else
                If ValorPedido > descontoPedido Then
                    MsgBox ("Desconto informado maior que " & descontoPedido & " permitido!")
                    TxtDescontoPedido.Text = 0
                    TxtDescontoPedido.SetFocus
                End If
            End If
        End If
    End If
    'Formata o número para apenas 2 casas decimais
    TxtDescontoPedido.Text = Format(TxtDescontoPedido.Text, "0.00")
                
    'CHAMA A FUNÇÃO QUE FAZ OS CALCULOS DO VALOR DO PEDIDO
    RecalcularValorPedido
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
    
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
    MsgBox " Ocorreu um erro ao informar a o desconto do pedido: " & TxtDescontoPedido & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub


Private Sub BtnFinalizar_Click()

    On Error GoTo TrataError

    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    Dim valorAntigo As Double
    Dim validar As Integer
    Dim resultado As Integer
    Dim verificarImpressao As Integer
    Dim tipoOperacao As Integer

    SQL = "SELECT * FROM ItensPedido WHERE CodPedido = " & TxtNumPedido.Text

    If STATUS = 0 Then ' PEDIDO NOVO (FINALIZAR)
            'Fazendo verificações
        rs.Open SQL, cn, adOpenStatic
            If TxtCliente.Text = "" Then
                MsgBox "Campo cliente não pode ser vazio"
                TxtCliente.SetFocus
            ElseIf TxtFormaPgto.Text = "" Then
                MsgBox "Campo Cond. Pagamento não pode ser vazio"
                TxtFormaPgto.SetFocus
            
            ElseIf rs.EOF = True Then
                MsgBox "Necessário informar ao menos um item no pedido!"
                TxtCodItem.SetFocus
            ElseIf TxtValor.Text = "" Then
                MsgBox "Campo Valor não pode ser vazio"
                TxtValor.SetFocus
            ElseIf TxtDescontoPedido = "" Then
                TxtDescontoPedido.Text = 0
            ElseIf TxtValorFinal.Text = "" Then
                MsgBox "Campo Valor Total não pode estar vazio"
                TxtValorFinal.SetFocus
            Else
            
                If PegaNomeCliente = 0 Then
                    Exit Sub
                End If
            
               'SELECT para buscar a soma dos valores do pedido
                SQL = "SELECT SUM(itens.ValorTotalProduto) " & _
                      "FROM ItensPedido AS itens " & _
                      "WHERE itens.CodPedido = " & TxtNumPedido.Text
                'Pega a soma dos itens
                valorAntigo = cn.Execute(SQL).GetString
                'Valida se a soma dos itens é diferente do valor informado
                If valorAntigo <> TxtValor.Text Then
                    validar = MsgBox("Valor do pedido alterado manualmente, prosseguir com a finalização?", vbYesNo)
                    If validar = vbYes Then
                        'Alterando a "," por "." para não dar erro no SQL ao adicionar o desconto
                        descontoPedido = Replace(TxtDescontoPedido.Text, ",", ".")
                        Valor = Replace(TxtValor.Text, ",", ".")
                        valorFinal = Replace(TxtValorFinal.Text, ",", ".")
                        
                        'Inserindo os valores na tabela PEDIDO
                        SQL = "INSERT INTO Pedido (PedidoIdCli, PedidoIdPgto, PedidoValor, PedidoDesconto, PedidoValorTotal) " & _
                              "VALUES (" & TxtCliente.Text & ", " & TxtFormaPgto.Text & ", " & Valor & ", " & descontoPedido & ", " & valorFinal & ")"
                        'Executando SQL
                        cn.Execute SQL
                        
                        verificarImpressao = MsgBox("Pedido Finalizado com sucesso! - Deseja imprimir o Pedido?", vbYesNo)
                        
                        If verificarImpressao = vbYes Then
                            imprimirPedido
                        End If
                            
                        ' FINALIZOU COM SUCESSO, SELECIONA O ULTIM PEDIDO NO BANCO
                        SQL = "DECLARE @maxPedido INT " & _
                              "SELECT @maxPedido = MAX(IdPedidos) FROM Pedido " & _
                              "SELECT @maxPedido AS Pedido"
                          
                        rsDados.Open SQL, cn, adOpenStatic
                            numeroPedido = rsDados("Pedido")
                        rsDados.Close
                        
                        ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
                        STATUS = 1
                            
                        ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
                        preenchePedido (numeroPedido)
                    Else
                        'Volta o foco para o txtValor
                        TxtValor.SetFocus
                    End If
                Else
                    resultado = finalizarPedido(TxtCliente.Text, TxtFormaPgto.Text, TxtValor.Text, TxtDescontoPedido.Text, TxtValorFinal.Text)
                        
                    If resultado = 0 Then
                        verificarImpressao = MsgBox("Pedido Finalizado com sucesso! - Deseja imprimir o Pedido?", vbYesNo)
                        
                        If verificarImpressao = vbYes Then
                            imprimirPedido
                        End If
                        
                        ' FINALIZOU COM SUCESSO, SELECIONA O ULTIM PEDIDO NO BANCO
                        SQL = "DECLARE @maxPedido INT " & _
                              "SELECT @maxPedido = MAX(IdPedidos) FROM Pedido " & _
                              "SELECT @maxPedido AS Pedido"
                          
                        rsDados.Open SQL, cn, adOpenStatic
                            numeroPedido = rsDados("Pedido")
                        rsDados.Close
                        
                        ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
                        STATUS = 1
                            
                        ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
                        preenchePedido (numeroPedido)
                    ElseIf resultado = 1 Then
                        Exit Sub
                    End If
                End If
            End If
        rs.Close
    ElseIf STATUS = 1 Then ' PEDIDO JÁ FINALIZADO (IMPRIMIR)
        ' CHAMA A FUNÇÃO DE IMPRIMIR PEDIDO
        imprimirPedido
    End If
Exit Sub

TrataError:
    MsgBox "Ocorreu um erro ao finalizar o pedido: " & TxtNumPedido & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub

Private Sub BtnCancelar_Click()
    
    On Error GoTo TrataErro
    
    ' CHAMA O FUNÇÃO QUE FINALIZA O FORM
    finalizaForm
    
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro ao tentar excluír o pedido N°: " & TxtNumPedido.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub BtnRecalcularValor_Click()
    
    On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim validar As Integer
    
    validar = MsgBox("Recalcular o valor do pedido?", vbYesNo)
    
    If validar = vbYes Then
        'SELECT para buscar a soma dos valores do pedido
        SQL = "SELECT SUM(itens.ValorTotalProduto) " & _
              "FROM ItensPedido AS itens " & _
              "WHERE itens.CodPedido = " & TxtNumPedido.Text
        
            rs.Open SQL, cn, adOpenStatic
                TxtValor.Text = rs.GetString
            rs.Close
            
            If TxtDescontoPedido.Text = "" Then
                descontoPedido = 0
            Else
                'Alterando a "," por "." para não dar erro no SQL
                descontoPedido = Replace(TxtDescontoPedido.Text, ",", ".")
            End If
            
            
            'SELECT para buscar o valot total do pedido (com o desconto)
            SQL = "SELECT SUM(itens.ValorTotalProduto) - (SUM(itens.ValorTotalProduto) * (" & descontoPedido & " * 0.01)) " & _
                  "FROM ItensPedido AS itens " & _
                  "WHERE itens.CodPedido = " & TxtNumPedido.Text
            'Adiconando o valor final do pedido
            rs.Open SQL, cn, adOpenStatic
                TxtValorFinal.Text = rs.GetString
            rs.Close
        'VOLTA O FOCO PARA O TXT COD ITEM
        TxtCodItem.SetFocus
    Else
        'VOLTA O FOCO PARA O TXT COD ITEM
        TxtCodItem.SetFocus
    End If
Exit Sub

TrataErro:
    MsgBox " Ocorreu um erro ao tentar recalcular o valor do pedido: " & TxtValor.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description

End Sub

Private Function finalizarPedido(PedidoIdCli As Integer, PedidoIdPgto As Integer, PedidoValor As String, PedidoDesconto As String, PedidoValorTotal As String)
        
    Dim CMD As New ADODB.Command
    Dim Parametros As New ADODB.parameter
    
    Dim Retorno As Integer
    
    With CMD
    Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "finalizarPedido"
    End With
    
        'Parametro 1
        nomeParametros = "RetornoOperacao"
        Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamReturnValue) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = -1 'Seta o valor do parametro (valor aleatório para teste)
        'Parametro 2
        nomeParametros = "OUTPUT"
        Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamOutput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = -1 'Seta o valor do parametro (valor aleatório para teste)
        'Parametro 3
        nomeParametros = "PedidoIdCli"
        Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = PedidoIdCli 'Seta o valor do parametro
        'Parametro 4
        nomeParametros = "PedidoIdPgto"
        Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = PedidoIdPgto 'Seta o valor do parametro
        'Parametro 5
        nomeParametros = "PedidoValor"
        Set Parametros = CMD.CreateParameter(nomeParametros, adNumeric, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            Parametros.Precision = 18 ' Tamanho
            Parametros.NumericScale = 2 ' Casas decimais
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = PedidoValor 'Valor 'Seta o valor do parametro
        'Parametro 6
        nomeParametros = "PedidoDesconto"
        Set Parametros = CMD.CreateParameter(nomeParametros, adNumeric, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            Parametros.Precision = 18 ' Tamanho
            Parametros.NumericScale = 2 ' Casas decimais
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = PedidoDesconto 'Seta o valor do parametro
        'Parametro 7
        nomeParametros = "PedidoValorTotal"
        Set Parametros = CMD.CreateParameter(nomeParametros, adNumeric, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            Parametros.Precision = 18 ' Tamanho
            Parametros.NumericScale = 2 ' Casas decimais
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = PedidoValorTotal 'Seta o valor do parametro
        'Executa o Command (SP)
        CMD.Execute
        'Adiciona o retorno da SP na variavel retorno
        Retorno = CMD.Parameters("RetornoOperacao").Value
        ' Se retorno = 0 deu certo | 1 = deu errado
    finalizarPedido = Retorno
End Function

Private Function RecalcularValorPedido()
    
    Dim Valor As Double
    Dim Desconto As Double
    Dim valorFinal As Double
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM ConfiguracoesGerais"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            descontoPedido = -1
        Else
            descontoPedido = Format(rs("MaxDescontoPedido"), "0.00")
        End If
    rs.Close
    
    If TxtValor.Text = "" Or TxtValor.Text = Null Then
        TxtValor.Text = 0
    End If
    
    Valor = TxtValor.Text
    
    TxtDescontoPedido.Text = Format(TxtDescontoPedido.Text, "0.00")
    
    If descontoPedido = -1 Then
        If ValorPedido = 100 Then
            MsgBox "Não é permitido dar uma desconto de 100%"
            TxtDescontoPedido.Text = 0
            TxtDescontoPedido.SetFocus
        ElseIf ValorPedido > 99.994 Then
            MsgBox "Valor de desconto inválido!"
            TxtDescontoPedido.Text = 0
            TxtDescontoPedido.SetFocus
        End If
    Else
        If ValorPedido > descontoPedido Then
            MsgBox ("Desconto informado maior que " & descontoPedido & " permitido!")
            TxtDescontoPedido.Text = 0
            TxtDescontoPedido.SetFocus
        Else
            If TxtDescontoPedido.Text = "" Then
                Desconto = 0
            Else
                Desconto = TxtDescontoPedido.Text
            End If
        End If
    End If
    
    'If TxtDescontoPedido.Text = "" Then
    '    Desconto = 0
    'ElseIf TxtDescontoPedido.Text > 99.994 Or TxtDescontoPedido.Text < 0 Then
    '    Desconto = 0
    'Else
    '    Desconto = TxtDescontoPedido.Text
    'End If

    'Faz o calculo do desconto
    valorFinal = (Valor - (Valor * (Desconto * 0.01)))
    
    'MsgBox Valor
    'MsgBox Desconto
    'MsgBox valorFinal
    'MOSTAR O VALOR NA TELA
    TxtValorFinal.Text = valorFinal
End Function

Private Function limpaCamposItem()

    TxtCodItem.Text = ""
    TxtNomeItem.Text = ""
    TxtValorItem.Text = ""
    TxtItemQuantidade.Text = ""
    TxtDescontoItem.Text = ""
    TxtValorFinalItem.Text = ""
    LblEstoqueItem.Caption = ""
    LblEstoqueFinal.Caption = ""
    LinhaSomaEstoque.Visible = False
    LblQuantiaEntrada.Caption = ""

End Function


Private Function preenchePedido(numeroPedido As Integer)
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim rsDados As New ADODB.Recordset
    Dim tipoOperacao As Integer
    
    ' SE STATUS = 0 ENTÃO PEDIDO NOVO
    If STATUS = 0 Then
        'HABILITO TUDO PARA QUE SEJA ALTERADO PARA DAR UMA NOVO PEDIDO
        TxtNumPedido.Enabled = False
        TxtCliente.Enabled = True
        TxtFormaPgto.Enabled = True
        ' SER FOR ADMIN PODE ALTERAR O VALOR DO PEDIDO MANUALMENTE
        If FormLogin.ADMIN = 1 Then
            TxtValor.Enabled = True
            TxtValorItem.Enabled = True
            ' SE FOR ADMIN NÃO PRECISA DO BOTÃO DE LIBERAÇÃO
            BtnLiberarAlteracao.Visible = False
        ElseIf FormLogin.ADMIN = 2 Then
            TxtValor.Enabled = False
            TxtValorItem.Enabled = False
            ' SE NÃO FOR ADMIN, PODE SOLICITAR LIBERAÇÃO
            BtnLiberarAlteracao.Visible = True
        End If
        TxtDescontoPedido.Enabled = True
        TxtValorFinal.Enabled = False
        BtnAdicionarItem.Enabled = True
        BtnCancelar.Enabled = True
        BtnNovo.Enabled = False
        BtnRecalcularValor.Enabled = True
        BtnRemoverItem.Enabled = True
        BtnAvancaRegistro.Enabled = False
        BtnVoltaRegistro.Enabled = False
        
        ' DESABILITA O MENU
        'MDIFormInicio.Menu.Enabled = False
        MDIFormInicio.Relatorio.Enabled = False
        MDIFormInicio.Configurações.Enabled = False
        
        ' LIMPA TODOS OS CAMPOS
        TxtCliente.Text = ""
        TxtFormaPgto.Text = ""
        TxtValor.Text = ""
        TxtDescontoPedido.Text = ""
        TxtValorFinal.Text = ""
        LblCliNome.Caption = ""
        LblFormaPgto.Caption = ""
        ListViewItensPedido.ListItems.Clear
        limpaCamposItem
        
        ' MUDA O NOME DO CAMPO IMPRIMIR PARA FINALIZAR
        BtnFinalizar.Caption = "Finalizar"
        
        ' DEFINO O FOCO PARA O TXT CLIENTE
        'SendKeys ("{tab}")
    ' SE STATUS = 1 ENTÃO PEDIDO JÁ FINALIZADO
    ElseIf STATUS = 1 Then
        SQL = "SELECT * " & _
              "From Pedido " & _
              "WHERE IdPedidos = " & numeroPedido
        
            rsDados.Open SQL, cn, adOpenStatic
                TxtNumPedido.Text = rsDados("IdPedidos")
                TxtCliente.Text = rsDados("PedidoIdCli")
                TxtFormaPgto.Text = rsDados("PedidoIdPgto")
                TxtValor.Text = rsDados("PedidoValor")
                TxtDescontoPedido.Text = rsDados("PedidoDesconto")
                TxtValorFinal.Text = rsDados("PedidoValorTotal")
            rsDados.Close
        Set rsDados = Nothing
        
        ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
        STATUS = 1
         
        'DESABILITO TUDO PARA QUE NÃO SEJA ALTERADO EM PEDIDOS NO MODO DE VISUALIZAÇÃO
        TxtNumPedido.Enabled = True
        TxtCliente.Enabled = False
        TxtFormaPgto.Enabled = False
        TxtValor.Enabled = False
        TxtDescontoPedido.Enabled = False
        TxtValorFinal.Enabled = False
        BtnAdicionarItem.Enabled = False
        BtnNovo.Enabled = True
        BtnCancelar.Enabled = False
        BtnRecalcularValor.Enabled = False
        BtnRemoverItem.Enabled = False
        BtnLiberarAlteracao.Visible = False
        BtnAvancaRegistro.Enabled = True
        BtnVoltaRegistro.Enabled = True
        
        ' HABILITA O MENU
        MDIFormInicio.Menu.Enabled = True
        MDIFormInicio.Relatorio.Enabled = True
        MDIFormInicio.Configurações.Enabled = True
        
        ' FORMATANDO OS CAMPOS
        TxtValor.Text = Format(TxtValor.Text, "0.00")
        TxtDescontoPedido.Text = Format(TxtDescontoPedido.Text, "0.00")
        
        ' MUDA O NOME DO CAMPO IMPRIMIR PARA FINALIZAR
        BtnFinalizar.Caption = "Imprimir"
        
        'Chama a função de atualizar a lista de itens
        atualizaListaPedido
    
        'CHAMA A FUNÇÃO DE ADICIONAR O NOME DO CLIENTE NA TELA
        PegaNomeCliente
        
        ' CHAMA A FUNÇÃO DE ADICIONAR O NOME DO CLIENTE NA TELA
        PegaNomeFormaPgto
    End If
End Function

Public Function imprimirPedido()

    Dim SQL As String
    Dim SQL2 As String
    Dim rs As New ADODB.Recordset
    
    ' SQL com o que vai buscas no banco
    SQL = "SELECT pe.IdPedidos AS Pedido, cli.CliNome AS Cliente, pgto.NomeFormaPgt AS FormaPGTO, pe.PedidoValor AS Valor, pe.PedidoDesconto AS Desconto, pe.PedidoValorTotal AS ValorTotal " & _
          "FROM Pedido as pe " & _
          "JOIN Clientes as cli on pe.PedidoIdCli = cli.IdCliente " & _
          "JOIN FormaPgto AS pgto on pe.PedidoIdPgto = pgto.IdFormaPgt " & _
          "WHERE pe.IdPedidos = " & TxtNumPedido.Text

    ' DEFINE O CABEÇALHO DO RELATÓRIO (IMPRSSÃO DE PEDIDO)
    rs.Open SQL, cn, adOpenStatic
        ImpressaoDePedido.Pedido = rs("Pedido")
        ImpressaoDePedido.Cliente = rs("Cliente")
        ImpressaoDePedido.FormaPGTO = rs("FormaPGTO")
        ImpressaoDePedido.Valor = rs("Valor")
        ImpressaoDePedido.Desconto = rs("Desconto")
        ImpressaoDePedido.ValorTotal = rs("ValorTotal")
    rs.Close

    ' DEFINE A CONEXÃO COM O BANCO
    ImpressaoDePedido.DataControlImpressaoPedido.ConnectionString = cn
    
    SQL2 = "SELECT itens.CodProduto AS Produto, produto.NomeProduto AS Nome, itens.ValorProduto AS ValorProduto, itens.QuantidadeProduto AS QuantidadeProduto, itens.DescontoProduto AS DescontoProduto, itens.ValorTotalProduto AS TotalProduto " & _
          "FROM ItensPedido AS itens " & _
          "JOIN Produtos AS produto ON itens.CodProduto = produto.CodProduto " & _
          "WHERE itens.CodPedido = " & TxtNumPedido.Text & " " & _
          "ORDER BY itens.CodItensPedido"
          
    ' PASSA A STRING QUE VAI SER EXECUTADA NO RELATÓRIO
    ImpressaoDePedido.DataControlImpressaoPedido.Source = SQL2
    
    ImpressaoDePedido.Show

End Function

Public Function PegaNomeCliente()
    
    'FAZ UMA VERIFICAÇÃO, SE FOR EM UM NOVO PEDIDO, VALIDA SE O CLIENTE ESTÁ INATIVO OU EXISTE
    '0 - NOVO PEDIDO
    '1 - PEDIDO JÁ FINALIZADO
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim Retorno As Integer
    
    Retorno = 1
    
    SQL = "SELECT CliNome, CliStatus FROM Clientes WHERE IdCliente = " & TxtCliente.Text
    
    rs.Open SQL, cn, adOpenStatic
        If STATUS = 0 Then
        
            ' FUNÇÃO QUE VERIFICA SE O CLIENTE EXISTE
            If verificaCliente(TxtCliente.Text) = 0 Then
                ' SE NÃO EXISTE (0), SAI DA FUNÇÃO
                Retorno = 0
                Exit Function
            End If
            
            ' FUNÇÃO QUE VERIFICA SE O TIPO DO CLIENTE É UM CLIENTE
            If verificaTipoCliente(TxtCliente.Text) = 0 Then
                ' SE CLIENTE NÃO TIVER O TIPO CLIENTE NO CADASTRO (0), SAI DA FUNÇÃO
                Retorno = 0
                Exit Function
            End If
            
            If rs("CliStatus") = False Then
                LblCliNome.Caption = rs("CliNome")
                SendKeys "{tab}" ' Set the focus to the next control.
            ElseIf rs("CliStatus") = True Then
                MsgBox "Cliente Inativo!"
                LblCliNome = ""
                TxtCliente = ""
                Retorno = 0
            End If
            
        ElseIf STATUS = 1 Then
            If rs.EOF = False Then
                LblCliNome.Caption = rs("CliNome")
            End If
        End If
    rs.Close
    
    Set rs = Nothing
    
    PegaNomeCliente = Retorno
    
End Function

Private Function verificaCliente(ID As Integer)
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim Retorno As Integer
    
    SQL = "SELECT CliNome, CliStatus FROM Clientes WHERE IdCliente = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            MsgBox "Cliente não cadastrado"
            LblCliNome = ""
            TxtCliente = ""
            TxtCliente.SetFocus
            
            Retorno = 0
        Else
            Retorno = 1
        End If
    rs.Close
    Set rs = Nothing
    
    verificaCliente = Retorno
    
End Function

Private Function verificaTipoCliente(ID As Integer)

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim Retorno As Integer
    
    SQL = "SELECT TipoCliente FROM TipoClientes WHERE IdCliente = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs("TipoCliente") = True Then
            Retorno = 1
        Else
            MsgBox ("Tipo cliente inválido para essa operação!")
            LblCliNome.Caption = ""
            TxtCliente.Text = ""
            TxtCliente.SetFocus
            
            Retorno = 0
        End If
    rs.Close
    Set rs = Nothing
    
    verificaTipoCliente = Retorno

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

Public Function buscaPedido(numeroPedido As Integer)
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    
    SQL = "SELECT * FROM Pedido WHERE IdPedidos = " & numeroPedido
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            
        Else
            MsgBox "Pedido " & numeroPedido & " não encontrado"
            
            SQL = "DECLARE @maxPedido INT " & _
                  "SELECT @maxPedido = MAX(IdPedidos) FROM Pedido " & _
                  "SELECT @maxPedido AS Pedido"
          
            rsDados.Open SQL, cn, adOpenStatic
                numeroPedido = rsDados("Pedido")
            rsDados.Close
        End If
    rs.Close
    
    ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
    STATUS = 1
    
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preenchePedido (numeroPedido)
    
    SendKeys "{tab}"   ' Set the focus to the next control.
    ' VOLTA O FOCO PARA O TXT PEDIDO
    SendKeys "+{tab}" ' SHIFT TAB
    
End Function

Private Function finalizaForm()

    Dim rs As New ADODB.Recordset
    Dim numeroPedido As Integer
    Dim SQL As String
    Dim validacao As Integer
    
    If STATUS = 0 Then ' PEDIDO NOVO
        validacao = MsgBox("Caso saia do pedido, os dados serão perdidos! Deseja mesmo sair?", vbYesNo)
        
        If validacao = vbYes Then
            'Limpa os itens do pedido
            SQL = "DELETE FROM ItensPedido WHERE CodPedido = " & TxtNumPedido.Text
            cn.Execute SQL
            
            ' VOLTA O STATUS PARA PEDIDO JÁ FINALIZADO
            STATUS = 1
            
            ' SQL PEGANDO O ULTIMO PEDIDO
            SQL = "DECLARE @maxPedido INT " & _
                  "SELECT @maxPedido = MAX(IdPedidos) FROM Pedido " & _
                  "SELECT @maxPedido AS Pedido"
            
            ' EXECUTA E ADICIONA O ULTIMO PEDIDO NA TELA
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("Pedido")) Then
                    STATUS = 0
                    numeroPedido = 1
                    TxtNumPedido.Text = 1
                Else
                    numeroPedido = rs("Pedido")
                    ' SETO O STATUS DA TELA PARA 1 -> PEDIDO JÁ FINALIZADO
                    STATUS = 1
                End If
            rs.Close
            ' LIMPA O RECORDSET
            Set rs = Nothing
            
            ' CHAMA A FUNÇÃO PARA ADICIONAR OS DADOS NA TELA
            preenchePedido (numeroPedido)
        ElseIf validacao = vbFalse Then
            TxtCliente.SetFocus
        End If
    Else ' PEDIDOS JÁ FINALIZADOS
        Unload Me
    End If
    
End Function

Public Function ValorProduto()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT ValorProduto FROM Produtos WHERE CodProduto = " & TxtCodItem.Text
    
    If TxtCodItem.Text = 0 Or TxtCodItem.Text = "" Then
        'DEIXA O CAMPO EM BRANCO
        TxtValorItem.Text = ""
    Else
        rs.Open SQL, cn, adOpenStatic
            'ADICIONA O VALOR DO PRODUTO NO CAMPO
            TxtValorItem.Text = Format(rs("ValorProduto"), "0.00")
        rs.Close
    End If
    
End Function

Private Function atualizaValorItem()
    
    Dim Valor As Double
    Dim Quantidade As Integer
    Dim Desconto As Double
    Dim ValorTotal As Double
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim descontoItem, EstoqueItem As Integer
       
    SQL = "SELECT * FROM ConfiguracoesGerais"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            descontoItem = -1
        Else
            descontoItem = Format(rs("MaxDescontoItemPedido"), "0.00")
        End If
    rs.Close
    
    ' PEGA O VALOR DO PRODUTO
    Valor = TxtValorItem.Text

    ' VERIFICA CAMPOS VAZIOS
    If TxtItemQuantidade.Text = "" Then
        TxtItemQuantidade.Text = 1
    Else
        Quantidade = TxtItemQuantidade.Text
    End If
    
    If TxtDescontoItem = "" Then
        TxtDescontoItem.Text = 0
    Else
        Desconto = TxtDescontoItem.Text
    End If

    ' SE DESCONTO ITEM = -1 UTILIZA O DESCONTO PADRAO DO SISTEMA MAX = 99,994
    If descontoItem = -1 Then
        If TxtDescontoItem.Text = 100 Then
            MsgBox ("Não é permitido o desconto de 100%")
            TxtDescontoItem.Text = ""
            Desconto = 0
            TxtDescontoItem.SetFocus
        ElseIf TxtDescontoItem.Text > 99.994 Then
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
    LblQuantiaEntrada.Caption = "-" & TxtItemQuantidade.Text
    
    ' ADICIONA O NOVO ESTOQUE
    LblEstoqueFinal.Caption = EstoqueItem - TxtItemQuantidade.Text
    
    ' EXECUTA
    cn.Execute (SQL)
    
    'CALCULA O VALOR COM BASE NOS PARAMETROS
    ValorTotal = ((Valor * Quantidade) - ((Valor * Quantidade) * (Desconto * 0.01)))
    
    'ADICIONA O VALOR CALCULADO NA TELA
    TxtValorFinalItem.Text = Format(ValorTotal, "0.00")
    
End Function

Private Function atualizaListaPedido()
    
    Dim rsDados As New ADODB.Recordset
    Dim SQL As String
    
    'SEMPRE LIMPA O LISTVIEW ANTES DE INSERIR NOVAMENTE
    ListViewItensPedido.ListItems.Clear
    
    'PEGA OS DADOS INICIAIS DA TABLE E ADICIONA NO GRID
    SQL = "SELECT itens.CodItensPedido AS CodItensPedido, itens.CodProduto AS Produto, produto.NomeProduto AS Nome, itens.ValorProduto AS Valor, itens.QuantidadeProduto AS Quantidade, itens.DescontoProduto AS Desconto, itens.ValorTotalProduto AS Total " & _
          "FROM ItensPedido AS itens " & _
          "JOIN Produtos AS produto ON itens.CodProduto = produto.CodProduto " & _
          "WHERE itens.CodPedido = " & TxtNumPedido.Text & " " & _
          "ORDER BY itens.CodItensPedido"
    rsDados.Open SQL, cn, adOpenDynamic
        Do While rsDados.EOF = False
            Set itens = ListViewItensPedido.ListItems.Add(, , rsDados("CodItensPedido"))
            itens.SubItems(1) = rsDados("Produto")
            itens.SubItems(2) = rsDados("Nome")
            itens.SubItems(3) = rsDados("Valor")
            itens.SubItems(4) = rsDados("Quantidade")
            itens.SubItems(5) = rsDados("Desconto")
            itens.SubItems(6) = rsDados("Total")
        
            'Move Para o próximo registro
            rsDados.MoveNext
        Loop
    rsDados.Close
End Function


Private Function atualizarListaPedidos()
    '-----------------------------------------------
    'EM PROCESSO DE DESONVOLVIMENTO, NÃO FUNCIONANDO
    '-----------------------------------------------
    Dim CMD As New ADODB.Command
    Dim Parametros As New ADODB.parameter
    Dim nomeParametros As String
    
    Dim Retorno As Variant
    
    With CMD
    Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "AtualizarListaItens"
    End With
    MsgBox "Teste 3"
    'Parametro 1
    nomeParametros = "RetornoOperacao"
    MsgBox "Teste 3.1"
    Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamReturnValue) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
        MsgBox "Teste 3.2"
        CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
        MsgBox "Teste 3.3"
        CMD.Parameters(nomeParametros).Value = -1 'Seta o valor do parametro (valor aleatório para teste)
    'Parametro 2
    MsgBox "Teste 4"
    nomeParametros = "OUTPUT"
    Set Parametros = CMD.CreateParameter(nomeParametros, adVarWChar, adParamOutput, 255) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
        CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
        CMD.Parameters(nomeParametros).Value = String(255, Chr(0))  'Seta o valor do parametro (valor aleatório para teste)
    'Parametro 3
    MsgBox "Teste 5"
    nomeParametros = "IdPedido"
    Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
        CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
        CMD.Parameters(nomeParametros).Value = TxtNumPedido.Text 'Seta o valor do parametro
    'Executa o Command (SP)
    MsgBox "Teste 6"
    CMD.Execute
    MsgBox "Teste 7"
    
End Function

Private Sub TxtValorFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtValorFinal_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtValorItem_GotFocus()

    With TxtValorItem
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtValorItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtValorItem_KeyPress(KeyAscii As Integer)

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

Private Sub TxtValorItem_LostFocus()

    'Formata o número para apenas 2 casas decimais
    TxtValorItem.Text = Format(TxtValorItem.Text, "0.00")

    If TxtCodItem.Text = "" Then
        ' NÃO ATUALIZA O VALOR DO ITEM
    Else
        ' CHAMA A FUNÇÃO QUE ATUALIZA O VALOR DO ITEM
        atualizaValorItem
    End If

End Sub
