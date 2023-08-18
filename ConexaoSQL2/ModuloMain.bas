Attribute VB_Name = "ModuloMain"
Public cn As New ADODB.Connection

Sub main()
    
    If conectaBanco = False Then
        MsgBox "Não foi possível conctar ao Banco de Dados"
        End
    Else
        FormLogin.Show 'Chama a tela de Login
    End If
End Sub

Private Function conectaBanco() As Boolean
    
    On Error GoTo TrataErro
    
    cn.Open "Provider=SQLOLEDB;" & _
                "Initial Catalog=LEONARDODB;" & _
                "Data Source=LEO\LEONARDODB;" & _
                "integrated security=SSPI; persist security info=True;"
            conectaBanco = True 'Retorna TRUE se conectou
Exit Function

TrataErro:
    conectaBanco = False 'Retorna FALSO se não conctou
End Function
