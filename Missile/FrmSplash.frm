VERSION 5.00
Begin VB.Form FrmSplashL 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3870
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "FrmSplash.frx":0000
   ScaleHeight     =   3870
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmSplashL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************************************************'
' Reviewing comment reading order should be:  (it should make reading easier to understand)'
'******************************************************************************************'
' 1. read modfunctions to understand how animation works, very importatn variables and other global functions/subs
' 2. read frmnew to understand how a new game is made
' 3. read Frmconfig to understand how the keys are linked to the main game
' 4. read Main to understand how the game works
' 5. read the rest, they don't matter too much, mostly all separate and not linked to each other

Private Sub Form_Load()

'save original display
    ' Initialize the structure.
    Dm.dmSize = Len(Dm)
    ' Get the display settings for the current monitor and mode.
    RetVal = EnumDisplaySettings(vbNullString, ENUM_CURRENT_SETTINGS, Dm)

'*********************'
' checks display mode '
'*********************'
ResChanged = False                                                 ' saying no res change at start
If Dm.dmPelsWidth <> 1024 Then                                  ' if not at 1024

    Dim NewRes As DEVMODE
    NewRes.dmSize = Len(Dm)
    
    EnumDisplaySettings vbNullString, ENUM_CURRENT_SETTINGS, NewRes
    NewRes.dmBitsPerPel = 16
    NewRes.dmDisplayFrequency = 75
    NewRes.dmPelsWidth = 1024
    NewRes.dmPelsHeight = 768
    
    If ChangeDisplaySettings(NewRes, CDS_TEST) <> DISP_CHANGE_SUCCESSFUL Then
        NewRes.dmDisplayFrequency = 60      ' change to a lower frequency
        If ChangeDisplaySettings(NewRes, CDS_TEST) <> DISP_CHANGE_SUCCESSFUL Then
          MsgBox "Warning! Your Screen Resolution is not 1024x768x16" & vbNewLine & _
             "Your system cannot run at 1024x768x16! This program will Terminate!", vbExclamation, "Error"
            End
        Else
            GoTo ChangeResNow
        End If
    Else
ChangeResNow:
            ' set screen res
            ResChanged = True
            DModeChangeStat = ChangeDisplaySettings(NewRes, &H1)
            Select Case DModeChangeStat
            ' Check for errors, there should be none since i just enumerated the display setting, but just in case
            Case 0
                'MsgBox "The display settings change was successful", vbInformation
            Case 1
                MsgBox "The computer must be restarted in order for the graphics mode to work", vbQuestion
                End
            Case -1
                MsgBox "The display driver failed the specified graphics mode", vbCritical
                End
            Case -2
                MsgBox "The graphics mode is not supported", vbCritical
                End
            Case -3
                MsgBox "Unable to write settings to the registry", vbCritical
                ' Windows NT Only
                End
            Case -4
                MsgBox "An invalid set of flags was passed in", vbCritical
                End
            End Select
    End If
End If
OriginalCursorPath = GetSettingString(HKEY_CURRENT_USER, "control panel\cursors\", "arrow", "")     ' get the location of teh current cursor
If Left(OriginalCursorPath, 12) = "%SYSTEMROOT%" Then                                               ' if it's in teh system root, then convert the system root name to c:\windows
    OriginalCursorPath = "c:\windows" & Right(OriginalCursorPath, Len(OriginalCursorPath) - 12)
End If
FrmSplashL.Show                                                     ' show it
sndPlaySound App.Path + "\sound\establish.wav", SND_ASYNC Or SND_NODEFAULT
Load Main                                                           ' load main
End Sub
