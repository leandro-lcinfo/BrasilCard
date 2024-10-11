Attribute VB_Name = "Geral"
Option Explicit
'Declaro as Variaveis***************************************************
Global vServidor
Global vBancoDados
Global vSQL As String
Global vConn
Global vRS As Recordset

Global vRelatorio As String
Global vFormula As String

Sub IniciaConexaoDB()

'Faz a conexão com o banco de dados local*******************************
vServidor = "leandro-vaio\sqlexpress"
vBancoDados = "CartaoDeCredito"

Set vConn = CreateObject("ADODB.Connection")
Set vRS = New ADODB.Recordset

'No meu caso estou usando o SQL SERVER com Autenticação do Windows
vConn.Open = "Provider=SQLOLEDB; " & _
            "Initial Catalog=" & vBancoDados & "; " & _
            "Data Source=" & vServidor & "; " & _
            "integrated security=SSPI; persist security info=True;"
'**********************************************************************

End Sub
