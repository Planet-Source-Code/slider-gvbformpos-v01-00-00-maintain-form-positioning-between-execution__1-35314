VERSION 5.00
Object = "{28889366-9C43-47DB-8AA5-BC05018A0F98}#3.0#0"; "gvbFormPos.ocx"
Begin VB.MDIForm fMDI 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   2625
   ClientTop       =   1590
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Begin gvbFormPostion.gvbFormPos gvbFormPos1 
      Left            =   1260
      Top             =   1155
      _ExtentX        =   1323
      _ExtentY        =   1138
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.CommandButton cmdDialog 
         Caption         =   "Centre Me!"
         Height          =   435
         Left            =   1680
         TabIndex        =   1
         Top             =   105
         Width           =   1485
      End
   End
End
Attribute VB_Name = "fMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:   fMDI
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        31/05/2002
' Version:     01.00.00
' Description: Example of using the gvbFormPos control with MDI child/parent
'              forms
' History:     01.00.00 Initial Release
' Notes:       If the StartUpPosition property of the form is *not*
'              '0 - Manual', then the gvbFormPos control will raise an
'              error. If any other StartUpPosition property setting, then
'              the desired effect won't be achieved.
'
'===========================================================================

Option Explicit

Private mbAbort As Boolean

Private Sub cmdDialog_Click()
    gvbFormPos1.CenterForm
End Sub

Private Sub MDIForm_Activate()
    '## error during load operation
    If mbAbort Then Unload Me
End Sub

Private Sub MDIForm_Load()

    On Error GoTo ErrorHandler
    gvbFormPos1.Hook Me, App.Title, "MDI Form"
    fChild.Show

    Exit Sub
ErrorHandler:
    MsgBox "ERROR Num = " + CStr(Err.Number) + "  Error Msg = " + Err.Description, vbCritical + vbApplicationModal, "Critical Error!"
    mbAbort = True
End Sub

Public Function Host()

End Function
