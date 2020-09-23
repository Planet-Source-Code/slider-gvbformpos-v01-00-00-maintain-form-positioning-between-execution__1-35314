VERSION 5.00
Object = "{28889366-9C43-47DB-8AA5-BC05018A0F98}#3.0#0"; "GVBFOR~2.OCX"
Begin VB.Form fTestForm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin gvbFormPostion.gvbFormPos gvbFormPos1 
      Left            =   210
      Top             =   210
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Centre Me!"
      Height          =   540
      Left            =   1575
      TabIndex        =   0
      Top             =   1155
      Width           =   1485
   End
End
Attribute VB_Name = "fTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:   gvbFormPos
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        31/05/2002
' Version:     01.00.00
' Description: Example of using the control with SDI forms
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

Private Sub Form_Activate()
    '## error during load operation
    If mbAbort Then Unload Me
End Sub

Private Sub Form_Initialize()

    On Error GoTo ErrorHandler
    gvbFormPos1.Hook Me, App.Title, gFormName
    Me.Caption = gFormName

    Exit Sub
ErrorHandler:
    MsgBox "ERROR Num = " + CStr(Err.Number) + "  Error Msg = " + Err.Description, vbCritical + vbApplicationModal, "Critical Error!"
    mbAbort = True
End Sub
