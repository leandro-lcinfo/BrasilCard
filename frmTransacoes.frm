VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransacoes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Transações de Cartão de Crédito"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totais"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   7080
      Width           =   13695
      Begin VB.Label totalBaixaLBL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Baixa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label totalMediaLBL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Média"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label totalAltaLBL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alta"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Valor Total R$"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label TotalValorLBL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label totalRegistrosLBL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Registros"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Controles"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   7920
      TabIndex        =   16
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command7 
         Height          =   855
         Left            =   120
         Picture         =   "frmTransacoes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cadastrar Clientes"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Height          =   855
         Left            =   1080
         Picture         =   "frmTransacoes.frx":3042
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cadastrar Transação"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Height          =   855
         Left            =   2040
         Picture         =   "frmTransacoes.frx":3F14
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Visualizar Relatório"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Height          =   855
         Left            =   3000
         Picture         =   "frmTransacoes.frx":7521
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exportar para Excel"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Height          =   855
         Left            =   3960
         Picture         =   "frmTransacoes.frx":AD47
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar Campos"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Height          =   855
         Left            =   4920
         Picture         =   "frmTransacoes.frx":E381
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar Janela"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Registros"
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   13695
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5055
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   8916
         _Version        =   393216
         BackColorBkg    =   16777215
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Filtros"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1440
      TabIndex        =   12
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton Command2 
         Height          =   855
         Left            =   5400
         Picture         =   "frmTransacoes.frx":119F3
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Pesquisar"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Valor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Descrição"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Número Cartão"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTIni 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130416641
         CurrentDate     =   45575
      End
      Begin MSComCtl2.DTPicker DTFim 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130416641
         CurrentDate     =   45575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Data Fim"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Data Início"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Image logoIMG 
      Height          =   1215
      Left            =   120
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vValor, vData, vCartao, vNome As String
Dim vValorTotal As Double
Dim vContador, vContadorA, vContadorM, vContadorB As Integer

Sub MontaGrid()
Grid1.FixedCols = 0
Grid1.Rows = 1
Grid1.Cols = 7

Grid1.ColWidth(0) = 0
Grid1.ColWidth(1) = 1200
Grid1.ColWidth(2) = 3000
Grid1.ColWidth(3) = 3000
Grid1.ColWidth(4) = 1200
Grid1.ColWidth(5) = 3500
Grid1.ColWidth(6) = 1200

Grid1.TextMatrix(0, 0) = "id"
Grid1.TextMatrix(0, 1) = "Data"
Grid1.TextMatrix(0, 2) = "Nome"
Grid1.TextMatrix(0, 3) = "Número do Cartão"
Grid1.TextMatrix(0, 4) = "Valor R$"
Grid1.TextMatrix(0, 5) = "Descrição"
Grid1.TextMatrix(0, 6) = "Categoria"

End Sub
Sub PesquisaRegistros()

If CDate(DTIni.Value) > CDate(DTFim.Value) Then
    MsgBox "A Data Inicial não pode ser maior que a Data Final!", vbInformation, "ATENÇÃO"
    Exit Sub
End If

Screen.MousePointer = 11

'Preparo a data para pesquisar*****************************************
vDataI = Format(Trim(DTIni.Value), "yyyy-mm-dd")
vDataF = Format(Trim(DTFim.Value), "yyyy-mm-dd")
'**********************************************************************

''Preparo o Select dos registros***************************************
Grid1.Rows = 1
vContador = 0
vContadorA = 0
vContadorM = 0
vContadorB = 0
vValorTotal = 0

vSQL = "select * from tblTransacao where "
vSQL = vSQL & "dataTransacao between '" & vDataI & "' and '" & vDataF & "'"

If Text1.Text <> "" Then
    If Option1.Value = True Then        'Numero do Cartão
        vSQL = vSQL & " and nuCartao like '%" & Text1.Text & "%'"
    ElseIf Option2.Value = True Then    'Descrição
        vSQL = vSQL & " and descricao like '%" & Text1.Text & "%'"
    ElseIf Option3.Value = True Then    'Valor
        vValor = Replace(Text1, ".", "")
        vValor = Replace(vValor, ",", ".")
        vSQL = vSQL & " and valorTransacao like '" & vValor & "%'"
    End If
End If

vSQL = vSQL & " order by dataTransacao asc"
Set vRS = vConn.Execute(vSQL)
Do While Not vRS.EOF
    
    'Acerto os valores para exbição
    vCartao = " " & vRS("nuCartao")
    vValor = FormatNumber(vRS("valorTransacao"), 2, vbTrue)
    vData = Format(Trim(vRS("dataTransacao")), "dd/mm/yyyy")
    
    'Seleciono o nome do cliente
    vNome = ""
    vSQL = "select * from tblClientes where nuCartao='" & Trim(vCartao) & "' "
    Set vRS1 = vConn.Execute(vSQL)
    If Not vRS1.EOF Then
        vNome = vRS1!nome
    End If
    
    'Verifico a Categoria
    If vRS("valorTransacao") > 1000 Then
        vCategoria = "ALTA"
        vContadorA = vContadorA + 1
    ElseIf vRS("valorTransacao") > 500 And vRS("valorTransacao") < 1000 Then
        vCategoria = "MÉDIA"
        vContadorM = vContadorM + 1
    ElseIf vRS("valorTransacao") < 500 Then
         vCategoria = "BAIXA"
         vContadorB = vContadorB + 1
    End If
    
    'Adiciono o registro na grid
    Grid1.AddItem vRS("idTransacao") & Chr(9) & vData _
    & Chr(9) & vNome & Chr(9) & vCartao & Chr(9) & vValor _
    & Chr(9) & vRS("descricao") & Chr(9) & vCategoria
    
    vValorTotal = CDbl(vValorTotal) + CDbl(vRS("valorTransacao"))
    vContador = vContador + 1
    
    vRS.MoveNext
Loop
'**********************************************************************

'Totais****************************************************************
totalRegistrosLBL = vContador
TotalValorLBL = FormatNumber(vValorTotal, 2, vbTrue)
totalAltaLBL = vContadorA
totalMediaLBL = vContadorM
totalBaixaLBL = vContadorB

vRS.Close

Screen.MousePointer = 0


End Sub




Private Sub Command1_Click()

'Fecha a conexão*******************************************************
vConn.Close
'**********************************************************************

'Fecha janela**********************************************************
End
'**********************************************************************

End Sub
Sub AjsutaDatas()

'Ajusta as datas de início e fim******************
Dim vDia, vMes, vAno, vData As String

vDia = "01"
vMes = Month(Date)
vAno = Year(Date)
vData = vDia & "/" & vMes & "/" & vAno

DTIni.Value = vData
DTFim.Value = DateAdd("d", -1, CDate("01/" & Month(DateAdd("m", 1, DTIni)) & "/" & Year(DateAdd("m", 1, DTIni))))
'*************************************************

End Sub

Private Sub Command2_Click()

PesquisaRegistros

End Sub

Private Sub Command3_Click()

AjsutaDatas
Option1.Value = True
Text1.Text = ""

PesquisaRegistros

End Sub


Private Sub Command4_Click()

Dim fApp As Excel.Application
Dim fBook As Excel.Workbook
Dim fSheet As Excel.Worksheet

If Grid1.Rows = 1 Then
    Exit Sub
End If

Screen.MousePointer = 11

   'Carregar o Excel:
   Set fApp = CreateObject("Excel.Application")
   'Crie um WorkBook:
   Set fBook = fApp.Workbooks.Add

   'Defina Uma nova Planilha
   Set fSheet = fApp.ActiveWorkbook.Sheets.Add

    fSheet.Name = "trasacoes.xlsx"

    fApp.Visible = True
    fSheet.Visible = True
    
   'Definir o conteúdo das células:
    vContador = 2
    
    vDataI = Format(Trim(DTIni.Value), "yyyy-mm-dd")
    vDataF = Format(Trim(DTFim.Value), "yyyy-mm-dd")
    vSQL = "select * from tblTransacao where "
    vSQL = vSQL & "dataTransacao between '" & vDataI & "' and '" & vDataF & "'"
    If Text1.Text <> "" Then
        If Option1.Value = True Then        'Numero do Cartão
            vSQL = vSQL & " and nuCartao like '%" & Text1.Text & "%'"
        ElseIf Option2.Value = True Then    'Descrição
            vSQL = vSQL & " and descricao like '%" & Text1.Text & "%'"
        ElseIf Option3.Value = True Then    'Valor
            vValor = Replace(Text1, ".", "")
            vValor = Replace(vValor, ",", ".")
            vSQL = vSQL & " and valorTransacao like '" & vValor & "%'"
        End If
    End If

    vSQL = vSQL & " order by dataTransacao asc"
    Set vRS = vConn.Execute(vSQL)
    If Not vRS.EOF Then
        With fSheet
        
            'Monto o cabeçalho
            .Cells(1, 1).Value = "Data"
            .Cells(1, 2).Value = "Nome"
            .Cells(1, 3).Value = "Cartão"
            .Cells(1, 4).Value = "Valor R$"
            .Cells(1, 5).Value = "Descrição"
            .Cells(1, 6).Value = "Categoria"
           
            Do While Not vRS.EOF
                
                'Variaveis da tabel Transacao
                vData = vRS!dataTransacao
                vCartao = vRS!nuCartao
                vValor = vRS!valorTransacao
                vDescricao = vRS!descricao
                
                'Seleciono o nome do cliente
                vNome = ""
                vSQL = "select * from tblClientes where nuCartao='" & Trim(vCartao) & "' "
                Set vRS1 = vConn.Execute(vSQL)
                If Not vRS1.EOF Then
                    vNome = vRS1!nome
                End If
    
                'Verifico a Categoria
                If vRS("valorTransacao") > 1000 Then
                    vCategoria = "ALTA"
                ElseIf vRS("valorTransacao") > 500 And vRS("valorTransacao") < 1000 Then
                    vCategoria = "MÉDIA"
                ElseIf vRS("valorTransacao") < 500 Then
                     vCategoria = "BAIXA"
                End If
                
                'Data***************************************
                .Cells(vContador, 1).Value = vData
                
                'Nome************************************
                .Cells(vContador, 2).Value = vNome
                
                'Cartao***************************
                .Cells(vContador, 3).Value = vCartao
                
                'Valor*****************************
                .Cells(vContador, 4).Value = vValor
                
                'Descricao*******************************
                .Cells(vContador, 5).Value = vDescricao
                
                'Categoria****************************
                .Cells(vContador, 6).Value = vCategoria
                
                vContador = vContador + 1
                vRS.MoveNext
            Loop
            
        End With
        vContador = vContador - 2
        MsgBox vContador & " produtos foram exportados para planilha!", vbInformation, "ATENÇÃO"
    End If

   'Limpe as variáveis de Objeto:
   Set fSheet = Nothing
   Set fBook = Nothing
   Set fApp = Nothing

Screen.MousePointer = 0

End Sub

Private Sub Command5_Click()
'
'vSQL = "select * from tblTransacao where "
'vSQL = vSQL & "dataTransacao between '" & vDataI & "' and '" & vDataF & "'"
'
'If Text1.Text <> "" Then
'    If Option1.Value = True Then        'Numero do Cartão
'        vSQL = vSQL & " and nuCartao like '%" & Text1.Text & "%'"
'    ElseIf Option2.Value = True Then    'Descrição
'        vSQL = vSQL & " and descricao like '%" & Text1.Text & "%'"
'    ElseIf Option3.Value = True Then    'Valor
'        vValor = Replace(Text1, ".", "")
'        vValor = Replace(vValor, ",", ".")
'        vSQL = vSQL & " and valorTransacao like '" & vValor & "%'"
'    End If
'End If
'
'vSQL = vSQL & " order by dataTransacao asc"


Unload ReportView

If Grid1.Rows = 1 Then
    MsgBox "Não há registros para exibição!", vbInformation, "ATENÇÃO"
    Exit Sub
End If

vRelatorio = "transacoes.rpt"

Dim vDiaI, vMesI, vAnoI, vDiaF, vMesF, vAnoF As String

vAnoI = Year(DTIni.Value)
vMesI = Month(DTIni.Value)
vDiaI = Day(DTIni.Value)

vAnoF = Year(DTFim.Value)
vMesF = Month(DTFim.Value)
vDiaF = Day(DTFim.Value)


vFormula = "date({tblTransacao.dataTransacao}) >= Date(" & vAnoI & "," & vMesI & "," & vDiaI & ") and " & _
           "date({tblTransacao.dataTransacao}) <= Date(" & vAnoF & "," & vMesF & "," & vDiaF & ")"
           
If Text1.Text <> "" Then
    If Option1.Value = True Then
        vFormula = vFormula & " and {tblTransacao.nuCartao} like '*" & Text1.Text & "*'"
    ElseIf Option2.Value = True Then
        vFormula = vFormula & " and {tblTransacao.descricao} like '" & Text1.Text & "*'"
    ElseIf Option3.Value = True Then
        vFormula = vFormula & " and {tblTransacao.valorTransacao} like '" & Text1.Text & "*'"
    End If
End If
               
           
           
           
           

ReportView.Show

End Sub

Private Sub Command6_Click()

frmCadastro.Show 1

End Sub

Private Sub Command7_Click()

frmClientes.Show 1

End Sub

Private Sub Form_Load()

'Ajusta a posição inicial do form**************************************
Me.Top = 0
Me.Left = 0
'**********************************************************************

'Carrega o logotipo da empresa*****************************************
logoIMG.Picture = LoadPicture(App.Path & "\miniLogo.jpg")
'**********************************************************************

'Ajusta as datas da pesquisa*******************************************
AjsutaDatas
'**********************************************************************

'Faz a conexão com o banco de dados ***********************************
IniciaConexaoDB
'**********************************************************************

'Pesquisa Registros para exibição**************************************
MontaGrid
PesquisaRegistros
'**********************************************************************

End Sub

Private Sub Grid1_Click()
Unload frmCadastro
frmCadastro.Show

Grid1.Col = 0
frmCadastro.idTXT = Grid1.Text

Grid1.Col = 1
frmCadastro.DT1.Value = Grid1.Text

Grid1.Col = 3
frmCadastro.nuCartaoTXT = Trim(Grid1.Text)

Grid1.Col = 4
frmCadastro.valorTXT = Grid1.Text

Grid1.Col = 5
frmCadastro.DescricaoTXT = Grid1.Text


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    PesquisaRegistros
End If

End Sub
