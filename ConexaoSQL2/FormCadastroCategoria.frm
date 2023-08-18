VERSION 5.00
Begin VB.Form FormCadastroCategoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Categorias"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   480
      TabIndex        =   9
      Top             =   2235
      Width           =   9645
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
         Height          =   660
         Left            =   7500
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   345
         Width           =   1755
      End
      Begin VB.CommandButton BtnConfirmar 
         Caption         =   "Confirmar"
         Height          =   660
         Left            =   375
         TabIndex        =   3
         Top             =   330
         Width           =   1755
      End
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "Cancelar"
         Height          =   660
         Left            =   2760
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   330
         Width           =   1755
      End
      Begin VB.CommandButton BtnNovo 
         Caption         =   "Novo"
         Height          =   660
         Left            =   5160
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   330
         Width           =   1755
      End
   End
   Begin VB.TextBox TxtCodigoCategoria 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2685
      TabIndex        =   1
      Top             =   390
      Width           =   1260
   End
   Begin VB.TextBox TxtNomeCategoria 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1830
      TabIndex        =   2
      Top             =   1305
      Width           =   8310
   End
   Begin VB.CheckBox ChkInativo 
      Caption         =   "Inativo"
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
      Height          =   525
      Left            =   7245
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   405
      Width           =   1530
   End
   Begin VB.CommandButton BtnPesquisarCategoria 
      Caption         =   "->"
      Height          =   450
      Left            =   1815
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   465
      Width           =   690
   End
   Begin VB.CommandButton BtnVoltaItem 
      Caption         =   "<"
      Height          =   555
      Left            =   4620
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   375
      Width           =   675
   End
   Begin VB.CommandButton BtnAvancaItem 
      Caption         =   ">"
      Height          =   555
      Left            =   5550
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   375
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   30
      TabIndex        =   8
      Top             =   480
      Width           =   1605
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   30
      TabIndex        =   7
      Top             =   1320
      Width           =   1605
   End
End
Attribute VB_Name = "FormCadastroCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STATUS As Integer

Private Sub BtnAlterar_Click()

    ' ALTERA PARA O STATUS 2 -> ALTERAÇÃO DE CATEGORIA JÁ CADASTRADA
    STATUS = 2
    ' CHAMA A FUNÇÃO QUE PREENCHE OS DADOS
    preencheCategoria (TxtCodigoCategoria.Text)
    
End Sub

Private Sub BtnAlterar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnAvancaItem_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim codCategoria As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtCodigoCategoria.Text & " " & _
          "SELECT TOP 1 IdCategoria FROM Categoria WHERE IdCategoria > @ID ORDER BY IdCategoria"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codCategoria = rs("IdCategoria")
        Else
            codCategoria = TxtCodigoCategoria.Text
        End If
    rs.Close
    
    TxtCodigoCategoria.Text = codCategoria
    preencheCategoria (codCategoria)

End Sub

Private Sub BtnAvancaItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnCancelar_Click()

    On Error GoTo TrataErro
    
    ' CHAMA O FUNÇÃO QUE FINALIZA O FORM
    finalizaForm
    
Exit Sub
TrataErro:
    MsgBox " Ocorreu um erro ao tentar cancelar o cadastro N°: " & TxtCodigoCategoria.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub BtnCancelar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnConfirmar_Click()

    On Error GoTo TrataError

    Dim rs As New ADODB.Recordset
    Dim statusCategoria As Integer
    Dim Retorno As Integer
    Dim numeroCategoria As Integer
    Dim validacao As Integer
    
    ' VERIFICA SE O NOME JÁ EXITE CADASTRADO
    SQL = "SELECT NomeCategoria " & _
          "FROM Categoria " & _
          "WHERE NomeCategoria LIKE '" & TxtNomeCategoria.Text & "' AND IdCategoria != " & TxtCodigoCategoria.Text
    
    rs.Open SQL, cn, adOpenStatic
        ' NÃO EXISTE NENHUM NOME IGUAL CADASTRADO
        If rs.EOF = True Then
            validacao = 0
        ' EXISTE UM NOME IGUAL CADASTRADO
        Else
            validacao = 1
        End If
    rs.Close
    
    If TxtNomeCategoria.Text = "" Then
        MsgBox ("Necessário informar o nome da Categoria!")
        TxtNomeCategoria.SetFocus
    ElseIf validacao = 1 Then
        MsgBox ("Nome de Categoria já cadastrado, informe um novo nome!")
        TxtNomeCategoria.SetFocus
    Else
        If ChkInativo.Value = Checked Then
            statusCategoria = 1
        ElseIf ChkInativo.Value = Unchecked Then
            statusCategoria = 0
        End If
        
        If STATUS = 0 Then
            Retorno = CadastroCategoria(-1, TxtNomeCategoria.Text, statusCategoria)
        ElseIf STATUS = 2 Then
            Retorno = CadastroCategoria(TxtCodigoCategoria.Text, TxtNomeCategoria.Text, statusCategoria)
        End If
        
        If Retorno = 0 Then
            MsgBox ("Categoria atualizada com sucesso!")
        ElseIf Retorno = 1 Then
            MsgBox ("Categoria cadastrada com sucesso!")
        ElseIf Retorno = 2 Then
            If STATUS = 0 Then
                MsgBox ("Ocorreu um erro ao tentar cadastrar a Categoria!")
                Exit Sub
            ElseIf STATUS = 2 Then
                MsgBox ("Ocorreu um erro ao tentar atualizar a Categoria!")
                Exit Sub
            End If
        End If
        
        If STATUS = 0 Then
            SQL = "DECLARE @MaxID INT " & _
                  "SELECT @MaxID = MAX(IdCategoria) FROM Categoria " & _
                  "SELECT @MaxID AS Categoria"
        
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("Categoria")) Then
                    STATUS = 0
                    numeroCategoria = 1
                    TxtCodigoCategoria.Text = 1
                Else
                    ' SETO O STATUS DA TELA PARA 1 -> CATEGORIA JÁ CADASTRADA
                    STATUS = 1
                    numeroCategoria = rs("Categoria")
                End If
            rs.Close
            
            ' LIMPA O RECORDSET
            Set rs = Nothing
        ElseIf STATUS = 2 Then
            STATUS = 1
            numeroCategoria = TxtCodigoCategoria.Text
        End If
        
        ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
        preencheCategoria (numeroCategoria)
    End If
    
Exit Sub
TrataError:
    MsgBox "Ocorreu um erro ao finalizar o cadastro: " & TxtCodigoCategoria.Text & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub

Private Sub BtnConfirmar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Public Sub BtnNovo_Click()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim numeroCategoria As Integer
    
    SQL = "SELECT MAX(IdCategoria)+1 AS Categoria FROM Categoria"
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Categoria")) = True Then
            numeroCategoria = 1
            TxtCodigoCategoria.Text = 1
        Else
            numeroCategoria = rs("Categoria")
            TxtCodigoCategoria.Text = rs("Categoria")
        End If
    rs.Close
    
    ' SETO O STATUS DA TELA PARA 0 -> PEDIDO NOVO
    STATUS = 0
    preencheCategoria (numeroCategoria)

End Sub

Private Sub BtnNovo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnPesquisarCategoria_Click()

    MDIFormInicio.Enabled = False
    FormBuscaCategoria.FORMULARIO = "FormCadastroCategoria"
    FormBuscaCategoria.Show

End Sub

Private Sub BtnPesquisarCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnVoltaItem_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim codCategoria As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtCodigoCategoria.Text & " " & _
          "SELECT TOP 1 IdCategoria FROM Categoria WHERE IdCategoria < @ID ORDER BY IdCategoria DESC"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codCategoria = rs("IdCategoria")
        Else
            codCategoria = TxtCodigoCategoria.Text
        End If
    rs.Close
    
    TxtCodigoCategoria.Text = codCategoria
    preencheCategoria (codCategoria)

End Sub

Private Sub BtnVoltaItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkInativo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo TrataErro
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim numeroCategoria As Integer
    
    SQL = "DECLARE @MaxID INT " & _
          "SELECT @MaxID = MAX(IdCategoria) FROM Categoria " & _
          "SELECT @MaxID AS Categoria"
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Categoria")) Then
            STATUS = 0
            numeroCategoria = 1
            TxtCodigoCategoria.Text = 1
        Else
            ' SETO O STATUS DA TELA PARA 1 -> CATEGORIA JÁ CADASTRADA
            STATUS = 1
            numeroCategoria = rs("Categoria")
        End If
    rs.Close
    
    ' LIMPA O RECORDSET
    Set rs = Nothing
    
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preencheCategoria (numeroCategoria)
    
Exit Sub
TrataErro:
    MsgBox "Algum erro ocorreu ao carregar o Form - " & FormCadastroCategoria.Name & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Function preencheCategoria(numeroCategoria As Integer)

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    ' NOVA CATEGORIA
    If STATUS = 0 Then
        
        ' LIBERAEÇÕES E BLOQUEIO DE CAMPOS
        TxtCodigoCategoria.Enabled = False
        BtnAvancaItem.Enabled = False
        BtnVoltaItem.Enabled = False
        TxtNomeCategoria.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        ElseIf FormLogin.ADMIN = 2 Then
            ChkInativo.Enabled = False
        End If
        BtnNovo.Enabled = False
        BtnCancelar.Enabled = True
        BtnConfirmar.Enabled = True
        BtnAlterar.Enabled = False
        
        ' LIMPAR CAMPOS
        TxtNomeCategoria.Text = ""
        ChkInativo.Value = Unchecked
        
        ' DEFINO O FOCO PARA O TXT NOME
        'SendKeys ("{tab}")
        
    ' CATEGORIA JÁ CADASTRADA
    ElseIf STATUS = 1 Then
    
        SQL = "SELECT * FROM Categoria " & _
              "WHERE IdCategoria = " & numeroCategoria

        rs.Open SQL, cn, adOpenStatic
            TxtCodigoCategoria.Text = rs("IdCategoria")
            TxtNomeCategoria.Text = rs("NomeCategoria")
            If rs("StatusCategoria") = False Then
                ChkInativo.Value = Unchecked
            ElseIf rs("StatusCategoria") = True Then
                ChkInativo.Value = Checked
            End If
        rs.Close
        Set rs = Nothing
        
        STATUS = 1
        
        ' LIBERAEÇÕES E BLOQUEIO DE CAMPOS
        TxtCodigoCategoria.Enabled = True
        BtnAvancaItem.Enabled = True
        BtnVoltaItem.Enabled = True
        TxtNomeCategoria.Enabled = False
        ChkInativo.Enabled = False
        BtnNovo.Enabled = True
        BtnCancelar.Enabled = False
        BtnConfirmar.Enabled = False
        BtnAlterar.Enabled = True
    
    ElseIf STATUS = 2 Then
        ' LIBERAEÇÕES E BLOQUEIO DE CAMPOS
        TxtCodigoCategoria.Enabled = False
        BtnAvancaItem.Enabled = False
        BtnVoltaItem.Enabled = False
        TxtNomeCategoria.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        ElseIf FormLogin.ADMIN = 2 Then
            ChkInativo.Enabled = False
        End If
        BtnNovo.Enabled = False
        BtnCancelar.Enabled = True
        BtnConfirmar.Enabled = True
        BtnAlterar.Enabled = False
    End If
    
End Function

Private Function finalizaForm()

    Dim rs As New ADODB.Recordset
    Dim numeroCategoria As Integer
    Dim SQL As String
    Dim validacao As Integer
    
    ' NOVO CADASTRO
    If STATUS = 0 Then
        validacao = MsgBox("Deseja cancelar o cadastro? Todos os dados serão perdidos!", vbYesNo)
        
        If validacao = vbYes Then
            ' VOLTA O STATUS PARA CATEGORIA JÁ CADASTRADA
            STATUS = 1
            
            SQL = "DECLARE @MaxID INT " & _
                  "SELECT @MaxID = MAX(IdCategoria) FROM Categoria " & _
                  "SELECT @MaxID AS Categoria"
            
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("Categoria")) Then
                    STATUS = 0
                    numeroCategoria = 1
                    TxtCodigoCategoria.Text = 1
                Else
                    ' SETO O STATUS DA TELA PARA 1 -> CATEGORIA JÁ CADASTRADA
                    STATUS = 1
                    numeroCategoria = rs("Categoria")
                End If
            rs.Close
            
            ' LIMPA O RECORDSET
            Set rs = Nothing
            
            ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
            preencheCategoria (numeroCategoria)
            
        Else
            ' NÃO CANCELOU O PEDIDO
            TxtNomeCategoria.SetFocus
        End If
    
    ' CATEGORIA JÁ CADATRADA
    ElseIf STATUS = 2 Then
        validacao = MsgBox("Deseja cancelar a alteração? Todos as mudanças serão perdidas!", vbYesNo)
        
        If validacao = vbYes Then
            ' VOLTA O STATUS PARA CATEGORIA JÁ CADASTRADA
            STATUS = 1
            
            numeroCategoria = TxtCodigoCategoria.Text
            
            ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
            preencheCategoria (numeroCategoria)
            
        Else
            ' NÃO CANCELOU O PEDIDO
            TxtCodigoCategoria.SetFocus
        End If
    Else
        Unload Me
    End If
End Function

Private Function CadastroCategoria(IdCategoria As Integer, nomeCategoria As String, statusCategoria As Integer)

    Dim CMD As New ADODB.Command
    Dim validacao As Integer
    
    ' CONFIGURA O COMMAND
    CMD.ActiveConnection = cn
    CMD.CommandText = "cadastroCategoria"
    CMD.CommandType = adCmdStoredProc
    
    ' PARAMETROS PARA A SP
    CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adSmallInt, adParamReturnValue, , 99)
    CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adSmallInt, adParamOutput, , 99)
    CMD.Parameters.Append CMD.CreateParameter("IdCategoria", adInteger, adParamInput, 12, IdCategoria)
    CMD.Parameters.Append CMD.CreateParameter("NomeCategoria", adVarChar, adParamInput, 50, nomeCategoria)
    CMD.Parameters.Append CMD.CreateParameter("StatusCategoria", adSmallInt, adParamInput, 1, statusCategoria)
    
    CMD.Execute
    
    validacao = CMD.Parameters("RetornoOperacao").Value
    
    CadastroCategoria = validacao
End Function

Private Sub TxtCodigoCategoria_GotFocus()

    With TxtCodigoCategoria
        .SelStart = 0
        .SelLength = Len(.Text)
    End With


End Sub

Private Sub TxtCodigoCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 115 Then ' F4
        BtnPesquisarCategoria_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtCodigoCategoria_KeyPress(KeyAscii As Integer)

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If TxtCodigoCategoria.Text = "" Then
            MsgBox ("Necessário informar uma Categoria!")
            KeyAscii = 0
        Else
            If STATUS = 1 Then
                ' FUNÇÃO QUE VERIFICA SE EXISTE ESTE PEDIDO NO SQL
                buscaCategoria (TxtCodigoCategoria.Text)
                KeyAscii = 0
            End If
        End If
    End If

End Sub

Public Function buscaCategoria(numeroCategoria As Integer)

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    
    SQL = "SELECT IdCategoria " & _
          "FROM Categoria " & _
          "WHERE IdCategoria = " & numeroCategoria
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            
        Else
            MsgBox "Categoria " & numeroCategoria & " não encontrado"
            
            SQL = "DECLARE @maxID INT " & _
                  "SELECT @maxID = MAX(IdCategoria) FROM Categoria " & _
                  "SELECT @maxID AS Categoria"
          
            rsDados.Open SQL, cn, adOpenStatic
                numeroCategoria = rsDados("Categoria")
            rsDados.Close
        End If
    rs.Close
    
    STATUS = 1
    preencheCategoria (numeroCategoria)
    
    SendKeys "{tab}"   ' Set the focus to the next control.
    ' VOLTA O FOCO PARA O ID CATEGORIA
    SendKeys "+{tab}" ' SHIFT TAB

End Function

Private Sub TxtNomeCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtNomeCategoria_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ' ENTER
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If

End Sub
