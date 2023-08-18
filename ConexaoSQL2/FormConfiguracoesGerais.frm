VERSION 5.00
Begin VB.Form FormConfiguracoesGerais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações Gerais"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkAlterarValorItem 
      Caption         =   "Permitir alteração do valor do item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   270
      TabIndex        =   9
      Top             =   2415
      Width           =   2715
   End
   Begin VB.CheckBox ChkVendaItemNegativo 
      Caption         =   "Permitir vendas com estoque negativo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   270
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1755
      Width           =   2940
   End
   Begin VB.CommandButton BtnSalvar 
      Caption         =   "Salvar"
      Height          =   735
      Left            =   2460
      TabIndex        =   1
      Top             =   4275
      Width           =   1935
   End
   Begin VB.TextBox TxtMaxDescontoItensPedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   270
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1020
   End
   Begin VB.CheckBox ChkItens 
      Caption         =   "Valor 0 como Máximo de desconto"
      Height          =   435
      Left            =   6900
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1065
      Width           =   1875
   End
   Begin VB.CheckBox ChkPedido 
      Caption         =   "Valor 0 como Máximo de desconto"
      Height          =   345
      Left            =   6915
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   1950
   End
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   5355
      TabIndex        =   2
      Top             =   4275
      Width           =   1935
   End
   Begin VB.TextBox TxtMaxDescontoPedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   270
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   1020
   End
   Begin VB.Label Label2 
      Caption         =   "Maximo Desconto Permitido nos Itens (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1425
      TabIndex        =   4
      Top             =   1095
      Width           =   5250
   End
   Begin VB.Label Label1 
      Caption         =   "Maximo Desconto Permitido no Pedido (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1410
      TabIndex        =   3
      Top             =   315
      Width           =   5460
   End
End
Attribute VB_Name = "FormConfiguracoesGerais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnSair_Click()
    finalizaForm
End Sub

Private Sub BtnSair_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnSalvar_Click()
    
    ' CHAMA A FUNÇÃO QUE EXECUTA A SP E SALVA A REGRA
    salvaAlteracoes
    
End Sub

Private Sub BtnSalvar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkAlterarValorItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkItens_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkPedido_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkVendaItemNegativo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub Form_Load()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM ConfiguracoesGerais"
    
    ' BUSCA AS INFORMAÇÕES DA TABELA
    rs.Open SQL, cn, adOpenStatic
        ' SE NÃO TIVER NADA NA TABELA DEFINE OS VALORES COMO "0"
        If rs.EOF = True Then
            TxtMaxDescontoItensPedido.Text = 0
            TxtMaxDescontoPedido.Text = 0
            ChkVendaItemNegativo.Value = Unchecked
        Else
            ' SE O VALOR FOR -1, ENTÃO O TXT É 0 E O CHECK BOX MARCADO
            If rs("MaxDescontoItemPedido") = -1 Then
                TxtMaxDescontoItensPedido.Text = 0
            ' SE O VALOR FOR 0 ENTÃO O CHECK BOX É MARCADO
            ElseIf rs("MaxDescontoItemPedido") = 0 Then
                TxtMaxDescontoItensPedido.Text = rs("MaxDescontoItemPedido")
                ChkItens.Value = Checked
            ' SENÃO O TXT RECEBE O VALOR BUSCADO
            Else
                TxtMaxDescontoItensPedido.Text = rs("MaxDescontoItemPedido")
            End If
            
            ' SE O VALOR FOR -1, ENTÃO O TXT É 0
            If rs("MaxDescontoPedido") = -1 Then
                TxtMaxDescontoPedido.Text = 0
            ' SE O VALOR FOR 0 ENTÃO O CHECK BOX É MARCADO
            ElseIf rs("MaxDescontoPedido") = 0 Then
                TxtMaxDescontoPedido.Text = rs("MaxDescontoPedido")
                ChkPedido.Value = Checked
            ' SENÃO O TXT RECEBE O VALOR BUSCADO
            Else
                TxtMaxDescontoPedido.Text = rs("MaxDescontoPedido")
            End If
            
            ' VERIFICA O VendaItemNegativo
            If rs("VendaItemNegativo") = 0 Then
                ChkVendaItemNegativo.Value = Checked
            Else
                ChkVendaItemNegativo.Value = Unchecked
            End If
            
            ' VERIFICA O AlterarValorItem
            If rs("AlterarValorItem") = 0 Then
                ChkAlterarValorItem.Value = Checked
            Else
                ChkAlterarValorItem.Value = Unchecked
            End If
        End If
    rs.Close
End Sub

Private Function finalizaForm()
    
    Unload Me

End Function

Private Sub Form_Unload(Cancel As Integer)
    ' ATIVA NOVAMENTE O MDI
    MDIFormInicio.Enabled = True
End Sub

Private Sub TxtMaxDescontoItensPedido_Change()
    
    ' CHAMA A FUNÇÃO DE VERIFICAR OS CHECK BOX
    VerificaChkBox
    
    If TxtMaxDescontoItensPedido.Text = "" Then
        TxtMaxDescontoItensPedido.Text = 0
    ElseIf TxtMaxDescontoItensPedido.Text < 0 Then
        TxtMaxDescontoItensPedido.Text = 0
    ElseIf TxtMaxDescontoItensPedido.Text > 100 Then
        TxtMaxDescontoItensPedido.Text = 100
    End If
    
End Sub

Private Sub TxtMaxDescontoItensPedido_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtMaxDescontoItensPedido_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtMaxDescontoItensPedido_LostFocus()

    TxtMaxDescontoItensPedido.Text = Format(TxtMaxDescontoItensPedido.Text, "0.00")

End Sub

Private Sub TxtMaxDescontoPedido_Change()
    
    ' CHAMA A FUNÇÃO DE VERIFICAR OS CHECK BOX
    VerificaChkBox
    
    If TxtMaxDescontoPedido.Text = "" Then
        TxtMaxDescontoPedido.Text = 0
    ElseIf TxtMaxDescontoPedido.Text > 100 Then
        TxtMaxDescontoPedido.Text = 100
    ElseIf TxtMaxDescontoPedido.Text < 0 Then
        TxtMaxDescontoPedido.Text = 0
    End If
        
End Sub

Private Function VerificaChkBox()

    ' VERIFICA SE O TXT DE DESCONTO PEDIDO IRÁ MOSTRAR OU NÃO O CHECK BOX -> MOSTRA SE FOR 0
    If TxtMaxDescontoPedido.Text = "" Then
        ChkPedido.Visible = False
    ElseIf TxtMaxDescontoPedido.Text = 0 Then
        ChkPedido.Visible = True
    Else
        ChkPedido.Visible = False
    End If
    
    ' VERIFICA SE O TXT DE DESCONTO ITEM IRÁ MOSTRAR OU NÃO O CHECK BOX -> MOSTRA SE FOR 0
    If TxtMaxDescontoItensPedido.Text = "" Then
        ChkItens.Visible = False
    ElseIf TxtMaxDescontoItensPedido.Text = 0 Then
        ChkItens.Visible = True
    Else
        ChkItens.Visible = False
    End If
    
End Function

Private Function salvaAlteracoes()

    Dim CMD As New ADODB.Command
    Dim ItemPedio As Double
    Dim Pedio As Double
    Dim validacao As Integer
    Dim VendaItemNegativo As Integer
    Dim AlterarValorItem As Integer
    
    ' VERIFICA SE O VALOR DO TXT É 0 E SE O CHECK BOX ESTÁ MARCADO
    If TxtMaxDescontoItensPedido.Text = 0 Then
        If ChkItens.Value = Checked Then
            ItemPedido = 0
        Else
            ItemPedido = -1
        End If
    Else
        ItemPedido = Replace(TxtMaxDescontoItensPedido.Text, ",", ".")
    End If
    
    ' VERIFICA SE O VALOR DO TXT É 0 E SE O CHECK BOX ESTÁ MARCADO
    If TxtMaxDescontoPedido.Text = 0 Then
        If ChkPedido.Value = Checked Then
            Pedido = 0
        Else
            Pedido = -1
        End If
    Else
        Pedido = Replace(TxtMaxDescontoPedido.Text, ",", ".")
    End If
    
    ' VERIFICA O CHECKBOX VendaItemNegativo => 0 -> PERMITE VENDA | 1 -> NÃO PERMITE
    If ChkVendaItemNegativo.Value = Checked Then
        VendaItemNegativo = 0
    Else
        VendaItemNegativo = 1
    End If
    
    ' VERIFICA O CHECKBOX AlterarValorItem => 0 -> PERMITE ALTERAÇÕES | 1 -> NÃO PERMITE ALTERAÇÕES
    If ChkAlterarValorItem.Value = Checked Then
        AlterarValorItem = 0
    Else
        AlterarValorItem = 1
    End If
    
    ' CONFIGURA O COMMAND
    CMD.ActiveConnection = cn
    CMD.CommandText = "AtualizaConfiguracoesGerais"
    CMD.CommandType = adCmdStoredProc
    
    ' PARAMETROS PARA A SP
    CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adSmallInt, adParamReturnValue, , 99)
    CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adSmallInt, adParamOutput, , 99)
    CMD.Parameters.Append CMD.CreateParameter("MaxDescontoItemPedido", adVarChar, adParamInput, 10, ItemPedido)
    CMD.Parameters.Append CMD.CreateParameter("MaxDescontoPedido", adVarChar, adParamInput, 10, Pedido)
    CMD.Parameters.Append CMD.CreateParameter("VendaItemNegativo", adSmallInt, adParamInput, 1, VendaItemNegativo)
    CMD.Parameters.Append CMD.CreateParameter("AlterarValorItem", adSmallInt, adParamInput, 1, AlterarValorItem)
    CMD.Parameters.Append CMD.CreateParameter("MaxDescontoItemEntrada", adSmallInt, adParamInput, 1, 10)
    CMD.Parameters.Append CMD.CreateParameter("MaxDescontoEntrada", adSmallInt, adParamInput, 1, 10)
    
    CMD.Execute
    
    validacao = CMD.Parameters("RetornoOperacao").Value
    ' 1 -> CRIOU UMA NOVA REGRA | 0 -> ATUALIZOU A REGRA | -1 -> ERRO
    
    If validacao = -1 Then
        MsgBox ("Ocorreu um erro ao tentar salvar as configurações!")
    ElseIf validacao = 1 Then
        MsgBox ("Nova Regra criada com sucesso!")
    ElseIf validacao = 0 Then
        MsgBox ("Regra atualizada com sucesso!")
        finalizaForm
    End If
    
End Function

Private Sub TxtMaxDescontoPedido_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtMaxDescontoPedido_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtMaxDescontoPedido_LostFocus()

    TxtMaxDescontoPedido.Text = Format(TxtMaxDescontoPedido.Text, "0.00")

End Sub
