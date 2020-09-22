VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FRMMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PDF Compact"
   ClientHeight    =   1770
   ClientLeft      =   3525
   ClientTop       =   3105
   ClientWidth     =   6585
   Icon            =   "FRMMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compact a PDF"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Status"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.Label LBLStatus 
         Alignment       =   2  'Center
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6255
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog ComBox 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FRMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Free 4 All 2 use... :)

Dim OnTop As New clsOnTop

'Define Objects
Dim AVApp As Acrobat.CAcroApp
Dim AVDoc As Acrobat.CAcroAVDoc
Dim PDDoc As Acrobat.CAcroPDDoc

Dim OriginalFileSize As Long
Dim NewFileSize As Long

Private Sub Command1_Click()
On Error GoTo errorfound
Dim Ans As VbMsgBoxResult
Dim PDSave As Acrobat.PDSaveFlags

OnTop.MakeTopMost Me.hWnd

ComBox.ShowOpen                                             'Show Common Dialog Box
If ComBox.FileName Like "*.pdf" Then
    Ans = MsgBox("You have selected " & ComBox.FileName & " to compact.  Is the correct?", vbYesNo, "Confirm PDF Compact???")
    If Ans = vbYes Then
        'Create Acrobat Objects
        Set AVApp = CreateObject(Class:="AcroExch.App")
        Set AVDoc = CreateObject(Class:="AcroExch.AVDoc")
        OriginalFileSize = FileLen(ComBox.FileName)         'Get Filesize
        AVApp.Show                                          'Show Acrobat App
        PDSave = PDSaveFull                                 'Set Save Flag
        AVDoc.Open ComBox.FileName, ComBox.FileName         'Open PDF Document
        
        Set PDDoc = AVDoc.GetPDDoc                          'Create Acrobat Object
        
        PDDoc.Save PDSave, ComBox.FileName                  'Save PDF Document
        PDDoc.Close                                         'Close Layer
        AVDoc.Close False                                   'Close layer
        AVApp.Exit                                          'Exit Program
        NewFileSize = FileLen(ComBox.FileName)              'Get Filesize
        FRMMain.LBLStatus.Caption = "[ Original  Size ] : " & OriginalFileSize & " Bytes" & vbNewLine & _
                                    "[ Compacted Size ] : " & NewFileSize & " Bytes"    'Update Interface
    End If
Else
    MsgBox "You have specified a file that isn't a pdf file."   'Error
End If
If NewFileSize = OriginalFileSize Then
    If NewFileSize = 0 Then
        FRMMain.LBLStatus.Caption = "No PDF File Selected..."   'Error
    Else
        MsgBox "No change in file size.  Is compressed to the max.", vbCritical, "No Change in File Size" 'Compacted Already
    End If
End If
errorfound:
End Sub

Private Sub Command2_Click()
End                         'End Program
End Sub

Private Sub Form_Load()
On Error GoTo errorfound
OnTop.MakeTopMost Me.hWnd
errorfound:
End Sub

Private Sub Timer1_Timer()
Static TCount As Long
If FRMMain.LBLStatus.Caption = "No PDF File Selected..." Then
    TCount = TCount + 1
    If TCount > 3 Then
        FRMMain.LBLStatus.Caption = ""  'Reset Label
        FRMMain.Refresh
    End If
End If
On Error GoTo errorfound
DoEvents
OnTop.MakeTopMost Me.hWnd
FRMMain.Caption = "PDF Compact " & Now  'Update Interface
errorfound:
End Sub


Public Function StatusUpdate(TXTString As String) As Boolean
'Update Interface Routine
On Error GoTo errorfound
FRMMain.LBLStatus.Caption = TXTString
FRMMain.LBLStatus.Refresh
StatusUpdate = True
Exit Function
errorfound:
StatusUpdate = False
End Function
