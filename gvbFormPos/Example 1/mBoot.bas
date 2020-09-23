Attribute VB_Name = "mBoot"
'===========================================================================
'
' Module Name: mBoot
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        31/05/2002
' Version:     01.00.00
' Description: Contains startup code for example 1 - SDI forms
' History:     01.00.00 Initial Release
' Notes:       If the StartUpPosition property of the form is *not*
'              '0 - Manual', then the gvbFormPos control will raise an
'              error. If any other StartUpPosition property setting, then
'              the desired effect won't be achieved.
'
'===========================================================================

Option Explicit

Private oF1 As fTestForm
Private of2 As fTestForm

Public gFormName As String

Private Sub Main()

    gFormName = "Test form 1"
    Set oF1 = New fTestForm

    gFormName = "Test form 2"
    Set of2 = New fTestForm
    
    oF1.Show
    of2.Show

End Sub

