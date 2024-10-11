VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCadastro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cadastrar Transação"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Informações do Registro"
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      Begin VB.TextBox idTXT 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox DescricaoTXT 
         Appearance      =   0  'Flat
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2280
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1440
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
         Format          =   156827649
         CurrentDate     =   45575
      End
      Begin VB.TextBox valorTXT 
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
         Left            =   1560
         TabIndex        =   2
         Text            =   "0"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox nuCartaoTXT 
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
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Descrição:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Data:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Valor R$:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Número do Cartão:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      Picture         =   "frmCadastro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Adicionar Registro"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      Picture         =   "frmCadastro.frx":361F
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Excluir Registro"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton LimparBTO 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      Picture         =   "frmCadastro.frx":6C13
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Limpar Campos"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   3120
      Picture         =   "frmCadastro.frx":A24D
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Fechar Janela"
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vDia, vMes, vAno, vData As String
Dim vCodigo As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

If MsgBox("Deseja excluir o registro?", vbYesNo, "ATENÇÃO") = vbNo Then
    Exit Sub
Else
    vSQL = "Delete from tblTransacao where idTransacao=" & idTXT & ""
    vConn.Execute (vSQL)
End If

'Limpa os campos
LimparBTO_Click

End Sub
Sub VerificaCartao()

'Verifico se o Cartão de Credito foi cadastrado**************************
If nuCartaoTXT.Text <> "" Then
    vSQL = "select * from tblclientes where nucartao='" & nuCartaoTXT.Text & "'"
    Set vRS = vConn.Execute(vSQL)
    If vRS.EOF Then
        MsgBox "O número de cartão informado não foi cadastrado para nenhum cliente!", vbInformation, "ATENÇÃO"
        Exit Sub
    End If
End If
'************************************************************************

End Sub
Private Sub Command4_Click()

'Verifico se os campos estão todos preenchidos***************************
If nuCartaoTXT.Text = "" Then
    MsgBox "O número do Cartão é obrigatório!", vbInformation, "ATENÇÃO"
    nuCartaoTXT.SetFocus
End If

If IsNumeric(valorTXT.Text) = False Then
    MsgBox "Digite um valor válido!", vbInformation, "ATENÇÃO"
    valorTXT.SetFocus
End If

If DescricaoTXT.Text = "" Then
    MsgBox "A descrição é obrigatória!", vbInformation, "ATENÇÃO"
    DescricaoTXT.SetFocus
End If
'************************************************************************

VerificaCartao


'Acerta o formato da data************************************************
vDia = Day(DT1.Value)
vMes = Month(DT1.Value)
vAno = Year(DT1.Value)
vData = vAno & "-" & vMes & "-" & vDia
'************************************************************************

'Acerta o valor do registro**********************************************
vValor = Replace(valorTXT, ".", "")
vValor = Replace(vValor, ",", ".")
'************************************************************************

'Verifico se existe um registro com o código de id******************
vCodigo = idTXT.Text
vSQL = "select * from tblTransacao where idTransacao=" & vCodigo & ""
Set vRS = vConn.Execute(vSQL)
If vRS.EOF Then
    'Se não houver faço a inclusão
    vSQL = "INSERT INTO  tblTransacao"
    vSQL = vSQL & "(nuCartao,      ValorTransacao,"
    vSQL = vSQL & "dataTransacao,  Descricao)"
    vSQL = vSQL & " VALUES "
    vSQL = vSQL & "('" & nuCartaoTXT & "',   '" & vValor & "',"
    vSQL = vSQL & "'" & DT1.Value & "',      '" & DescricaoTXT & "')"
    vConn.Execute (vSQL)
    
    vSQL = "select max(idTransacao)as codigo from tblTransacao"
    Set vRS = vConn.Execute(vSQL)
    If IsNull(vRS!codigo) Then
        idTXT.Text = 1
    Else
        idTXT.Text = vRS!codigo
    End If
Else
    'Se houver gravo o registro novamente
    vSQL = "update tblTransacao set "
    vSQL = vSQL & "tblTransacao.nuCartao          ='" & Trim(nuCartaoTXT) & "',"
    vSQL = vSQL & "tblTransacao.ValorTransacao    ='" & vValor & "',"
    vSQL = vSQL & "tblTransacao.dataTransacao     ='" & vData & "',"
    vSQL = vSQL & "tblTransacao.Descricao         ='" & DescricaoTXT & "' "
    vSQL = vSQL & "where tblTransacao.IdTransacao = " & vCodigo & ""
    vConn.Execute (vSQL)
End If
'**************************************************************************

MsgBox "Registro Gravado com sucesso!", vbInformation, "ATENÇÃO"

LimparBTO_Click

End Sub

Private Sub Form_Load()

'Ajusta a posição inicial do form**************************************
Me.Top = 0
Me.Left = 0
'**********************************************************************

End Sub



Private Sub LimparBTO_Click()
idTXT = 0
nuCartaoTXT = ""
DT1 = Date
valorTXT = 0
DescricaoTXT = ""

nuCartaoTXT.SetFocus

frmTransacoes.PesquisaRegistros

End Sub

Private Sub nuCartaoTXT_LostFocus()

VerificaCartao

End Sub

Private Sub valorTXT_GotFocus()

valorTXT.SelStart = 0
valorTXT.SelLength = Len(valorTXT.Text)

End Sub

Private Sub valorTXT_LostFocus()

If IsNumeric(valorTXT.Text) = False Then
    valorTXT.Text = 0
Else
    valorTXT.Text = FormatNumber(valorTXT.Text, 2, vbTrue)
End If

End Sub
