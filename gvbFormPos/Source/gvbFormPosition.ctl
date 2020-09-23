VERSION 5.00
Begin VB.UserControl gvbFormPos 
   CanGetFocus     =   0   'False
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ForwardFocus    =   -1  'True
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "gvbFormPosition.ctx":0000
   ScaleHeight     =   2220
   ScaleWidth      =   3495
   ToolboxBitmap   =   "gvbFormPosition.ctx":030A
End
Attribute VB_Name = "gvbFormPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Control Name: gvbFormPos
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         31/05/2002
' Version:      01.00.00
' Description:  Form Handler
' History:      01.00.00 Initial Release
' Notes:        This class wraps the saving & loading to/from the registry
'               plus centering of a Form/MDIForm.
'
'               The original *raw* code author for saving & loading was Elias
'               Barbosa (made public via Planet Source Code 30/05/2002). Link:-
'               www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=35296&lngWId=1
'               This code unforetunately did not handle the moving of a form
'               which is a main requirement for myself and probably many
'               other developers. This Custom Control eliminates this problem
'               through the use of subclassing (The original subclassing code
'               (now modified to work within a usercontrol) came from:-
'               www.vbaccelerator.com) the Window's PositionChanged message.
'
'===========================================================================

Option Explicit

'===========================================================================
' Debugging... Saves adding the debug statements to the form events
'
#Const DEBUGMODE = 1                    '## 0=No debug
                                        '   1=debug
#If DEBUGMODE = 1 Then
    Private dbgCtrlName As String
#End If

'===========================================================================
'## Private: Windows 32-bit API Declarations
'
Private Const WM_WINDOWPOSCHANGED As Long = &H47

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'===========================================================================
' Private: Declarations
'
Private Enum eMode
    [Resize] = 0
    [Unload] = 1
End Enum

Private WithEvents oMdiForm   As VB.MDIForm
Attribute oMdiForm.VB_VarHelpID = -1
Private WithEvents oForm      As VB.Form
Attribute oForm.VB_VarHelpID = -1

Private moHookedForm          As Object
Private msAppName             As String
Private msSection             As String
Private mbIsBusy              As Boolean

Private m_emr                 As EMsgResponse
Private moSubclass            As GSubclass
Attribute moSubclass.VB_VarHelpID = -1

'===========================================================================
' Private: cTreeView Internal Error Declarations
'
Private Const csSOURCE_ERR    As String = "cForm"
Private Const clFORMHOOK_ERR1 As Long = vbObjectError + 101
Private Const csFORMHOOK_ERR1 As String = "Invalid object! Must be a VB Form or MDIForm."
Private Const clFORMHOOK_ERR2 As Long = vbObjectError + 102
Private Const csFORMHOOK_ERR2 As String = "Invalid StartUpPosition! Must be set to '0 - Manual'."
Private Const clFORMHOOK_ERR3 As Long = vbObjectError + 103
Private Const csFORMHOOK_ERR3 As String = "No Form hooked!"

'===========================================================================
' Public subroutines and functions
'
Public Sub CenterForm()

    Dim tCRect   As RECT        '## Holds the area that the form is to be centered in
    Dim tTBRect  As RECT        '## Holds the TaskBar area
    Dim X        As Single      '## temp LeftPosition
    Dim Y        As Single      '## temp TopPosition
    Dim bIsChild As Boolean
    Dim hTray    As Long
   
    If (moHookedForm Is Nothing) Then
        Err.Raise clFORMHOOK_ERR3, csSOURCE_ERR, csFORMHOOK_ERR3
        Exit Sub
    End If
    '## Just incase we have *the* MDIForm
    '
    On Error Resume Next
    bIsChild = moHookedForm.MDIChild

    If bIsChild Then                                            '## Check if the form is a MDIChild.
        GetClientRect GetParent(moHookedForm.hwnd), tCRect      '## Center it in the MDIParent.
    Else
        '## Center it in the available desktop area.
        '
        GetClientRect GetDesktopWindow(), tCRect                '## Get the Desktop area
        hTray = FindWindow("Shell_TrayWnd", vbNullString)       '## Check for the Task Bar.
        If hTray Then                                           '## Is there a taskbar?
            GetWindowRect hTray, tTBRect                        '## Get Taskbar area
            If (tTBRect.Right - tTBRect.Left) > (tTBRect.Bottom - tTBRect.Top) Then
                If tTBRect.Top <= 0 Then
                    tCRect.Top = tCRect.Top + tTBRect.Bottom                        '## Top of Screen.
                Else
                    tCRect.Bottom = tCRect.Bottom - (tTBRect.Bottom - tTBRect.Top)  '## Bottom of Screen.
                End If
            Else
                If tTBRect.Left <= 0 Then
                    tCRect.Left = tCRect.Left + tTBRect.Right                       '## Left of the Screen.
                Else
                    tCRect.Right = tCRect.Right - (tTBRect.Right - tTBRect.Left)    '## Right of the Screen.
                End If
            End If
        End If
    End If

    '## Center the Form
    '
    With moHookedForm
        X = (((tCRect.Right - tCRect.Left) * Screen.TwipsPerPixelX) - .Width) / 2
        Y = (((tCRect.Bottom - tCRect.Top) * Screen.TwipsPerPixelY) - .Height) / 2
        .Move X, Y
    End With

End Sub

Public Sub Hook(oNewForm As Object, AppName As String, RegSectionName As String)

    If (TypeOf oNewForm Is VB.MDIForm) Then
        Set oMdiForm = oNewForm
    ElseIf (TypeOf oNewForm Is VB.Form) Then
        Set oForm = oNewForm
    Else
        Err.Raise clFORMHOOK_ERR1, csSOURCE_ERR, csFORMHOOK_ERR1
        Exit Sub
    End If
    
    mbIsBusy = True
    If Not (oNewForm.StartUpPosition = vbStartUpManual) Then
        Err.Raise clFORMHOOK_ERR2, csSOURCE_ERR, csFORMHOOK_ERR2
        Exit Sub
    End If

    If Len(RegSectionName) = 0 Then
        msSection = oNewForm.Caption
    Else
        msSection = RegSectionName
    End If
    If Len(AppName) = 0 Then
        msAppName = msSection
    End If
    msAppName = AppName

    Set moHookedForm = oNewForm

    #If DEBUGMODE = 1 Then
        With moHookedForm
            dbgCtrlName = .Name + "(" + UserControl.Name + ")"
        End With
    #End If

    Call GetSettings                '## Load Form information
    Set moSubclass = New GSubclass

    moSubclass.AttachMessage Me, moHookedForm.hwnd, WM_WINDOWPOSCHANGED

End Sub

'===========================================================================
' Friend: Subclassing
'
Friend Property Let MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Friend Property Get MsgResponse() As EMsgResponse
    'Debug.Print CurrentMessage
    MsgResponse = m_emr
End Property

Friend Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If iMsg = WM_WINDOWPOSCHANGED Then
        #If DEBUGMODE = 1 Then
            Debug.Print dbgCtrlName + "::Subclass (Move Event)"
        #End If
        Call SaveSettings([Resize])
    End If
    m_emr = emrPostProcess
End Function

'===========================================================================
' Private: Usercontrol Events
'
Private Sub UserControl_Resize()
    Size 500, 500
End Sub

'===========================================================================
' Private: Form/MDIForm Events
'
Private Sub oForm_Resize()
    Call SaveSettings([Resize])
End Sub

Private Sub oForm_Unload(Cancel As Integer)
    Call SaveSettings([Unload])
End Sub

Private Sub oMdiForm_Resize()
    Call SaveSettings([Resize])
End Sub

Private Sub oMdiForm_Unload(Cancel As Integer)
    Call SaveSettings([Unload])
End Sub

'===========================================================================
' Private: Internal Methods
'
Private Sub GetSettings()

    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName + "::GetSettings (Load Event)"
    #End If

    '## Get the size that the window is supposed to have if it is on Normal state.
    '
    With moHookedForm
        mbIsBusy = True

        .Left = GetSetting(msAppName, msSection, "Left", .Left)
        .Top = GetSetting(msAppName, msSection, "Top", .Top)
        .Width = GetSetting(msAppName, msSection, "Width", .Width)
        .Height = GetSetting(msAppName, msSection, "Height", .Height)
        '
        '## Now, set the WindowState.
        '
        .WindowState = GetSetting(msAppName, "Settings", "MainWindowState", vbNormal)
        mbIsBusy = False
    End With

End Sub

Private Sub SaveSettings(Mode As eMode)

    With moHookedForm
        If mbIsBusy Then Exit Sub

        Select Case Mode
            Case [Resize]

                #If DEBUGMODE = 1 Then
                    Debug.Print dbgCtrlName + "::SaveSettings:Window_Size (Resize Event)"
                #End If

                '## Only save the window size if it is on Normal WindowState. Check,
                '   also, the Visible property to avoid saving the window size while
                '   the Form is loading. This will prevent conflicts.
                '
                If (.WindowState = vbNormal) And (.Visible) Then
                    SaveSetting msAppName, msSection, "Left", .Left
                    SaveSetting msAppName, msSection, "Top", .Top
                    SaveSetting msAppName, msSection, "Width", .Width
                    SaveSetting msAppName, msSection, "Height", .Height
                End If

            Case [Unload]

                #If DEBUGMODE = 1 Then
                    Debug.Print dbgCtrlName + "::SaveSettings:Window_State (Unload Event)"
                #End If

                '## Save WindowState only if the window is not on Minimized state.
                '
                If (.WindowState <> vbMinimized) Then
                    SaveSetting msAppName, msSection, "State", .WindowState
                End If

                moSubclass.DetachMessage Me, moHookedForm.hwnd, WM_WINDOWPOSCHANGED
                Set moSubclass = Nothing

        End Select
    End With

End Sub
