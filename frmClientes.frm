VERSION 5.00
Begin VB.Form frmClientes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cadastrar Cleintes"
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Registros"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   2895
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1785
         ItemData        =   "frmClientes.frx":0000
         Left            =   120
         List            =   "frmClientes.frx":0002
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Informações do Registro"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2895
      Begin VB.TextBox nomeTXT 
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
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox idTXT 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   1
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nome:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Número do Cartão:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      Picture         =   "frmClientes.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Adicionar Registro"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      Picture         =   "frmClientes.frx":3623
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Excluir Registro"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton LimparBTO 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      Picture         =   "frmClientes.frx":6C17
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Limpar Campos"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   3120
      Picture         =   "frmClientes.frx":A251
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Fechar Janela"
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "frmClientes"
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
    vSQL = "Delete from tblCleintes where idcliente=" & idTXT & ""
    vConn.Execute (vSQL)
End If

'Limpa os campos
LimparBTO_Click

End Sub

Private Sub Command4_Click()

'Verifico se os campos estão todos preenchidos***************************
If nomeTXT.Text = "" Then
    MsgBox "O Nome do Cliente é obrigatório!", vbInformation, "ATENÇÃO"
    nomeTXT.SetFocus
End If

If nuCartaoTXT.Text = "" Then
    MsgBox "O número do Cartão é obrigatório!", vbInformation, "ATENÇÃO"
    nuCartaoTXT.SetFocus
End If
'************************************************************************

'Verifico se existe um registro com o código de id******************
vCodigo = idTXT.Text
vSQL = "select * from tblClientes where idCliente=" & vCodigo & ""
Set vRS = vConn.Execute(vSQL)
If vRS.EOF Then
    'Se não houver faço a inclusão
    vSQL = "INSERT INTO  tblClientes"
    vSQL = vSQL & "(nome,nuCartao)"
    vSQL = vSQL & " VALUES "
    vSQL = vSQL & "('" & nomeTXT & "','" & nuCartaoTXT & "')"
    vConn.Execute (vSQL)
    
Else
    'Se houver gravo o registro novamente
    vSQL = "update tblClientes set "
    vSQL = vSQL & "tblClientes.nome             ='" & nomeTXT & "',"
    vSQL = vSQL & "tblClientes.nuCartao        ='" & nuCartaoTXT & "' "
    vSQL = vSQL & "where tblClientes.Idcliente = " & vCodigo & ""
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

'Carrega os registros de clientes *************************************
CarregaNome
'**********************************************************************

End Sub

Sub CarregaNome()

List1.Clear
vSQL = "select * from tblClientes order by nome asc"
Set vRS = vConn.Execute(vSQL)
Do While Not vRS.EOF
    List1.AddItem vRS("nome")
    vRS.MoveNext
Loop

End Sub

Private Sub LimparBTO_Click()
idTXT = 0
nomeTXT.Text = ""
nuCartaoTXT = ""
nomeTXT.SetFocus

CarregaNome

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

Private Sub List1_Click()

vNome = List1.Text
vSQL = "select * from tblClientes where nome='" & vNome & "'"
Set vRS = vConn.Execute(vSQL)
If Not vRS.EOF Then
    idTXT.Text = vRS!idCliente
    nomeTXT.Text = vRS!nome
    nuCartaoTXT.Text = vRS!nucartao
End If



End Sub
