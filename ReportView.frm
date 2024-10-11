VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form ReportView 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizador de Relatórios"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10785
   Icon            =   "ReportView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   10785
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CrViewer 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _cx             =   19076
      _cy             =   16536
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1046
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "ReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public apl As New CRAXDRT.Application
Public rpt As New CRAXDRT.Report
Public WithEvents sect As CRAXDRT.Section
Attribute sect.VB_VarHelpID = -1


Private Sub Form_Load()

Dim RET As Boolean

Me.Top = 0
Me.Left = 0
Me.Height = 9840
Me.Width = 10875

Set rpt = apl.OpenReport(App.Path & "\" & vRelatorio)

rpt.RecordSelectionFormula = vFormula

CrViewer.ReportSource = rpt
CrViewer.ViewReport

CrViewer.Zoom 1

End Sub

