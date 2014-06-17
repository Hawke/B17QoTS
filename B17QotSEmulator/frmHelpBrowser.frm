'******************************************************************************
' frmHelpBrowser.frm
'
' @author Preston V. McMurry III, http://www.prestonm.com
' @copyright (C) Copyright 2002, 2010 by Preston V. McMurry III, http://www.prestonm.com
'
' *****************************************************************************
'
' This file is part of B17QotS, the "B-17: Queen of the Skies" Emulator.
'
' B17QotS is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' B17QotS is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with B17QotS. If not, see <http://www.gnu.org/licenses/>.
'******************************************************************************
VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHelpBrowser 
   ClientHeight    =   5085
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   6540
   Icon            =   "frmHelpBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6540
   Visible         =   0   'False
   Begin VB.TextBox txtPageName 
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3734
      Left            =   50
      TabIndex        =   0
      Top             =   1215
      Width           =   5393
      ExtentX         =   9525
      ExtentY         =   6588
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6000
      Top             =   1500
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6540
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   6540
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Top             =   300
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpBrowser.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpBrowser.frx":11AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpBrowser.frx":148E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpBrowser.frx":1770
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpBrowser.frx":1A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpBrowser.frx":1D34
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHelpBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

'******************************************************************************
' Form_Load
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub Form_Load()
    On Error Resume Next
    
    Me.Show
    tbToolBar.Refresh
    Form_Resize

    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15
End Sub

'******************************************************************************
' Form_Activate
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub Form_Activate()
    StartingAddress = CurDir & "\" & txtPageName.Text
'assfucker
'    StartingAddress = App.Path & "\" & txtPageName.Text

    If Len(StartingAddress) > 0 Then
        cboAddress.AddItem StartingAddress
       
'cboAddress.Text = StartingAddress
'cboAddress.AddItem cboAddress.Text
       
       'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If

End Sub

'******************************************************************************
' brwWebBrowser_NavigateComplete2
'
' INPUT:  ???
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    
    Dim intIndex As Integer
    Dim blnFound As Boolean

    Call ParseCaptionFromCallerName
    
    For intIndex = 0 To cboAddress.ListCount - 1
        If cboAddress.List(intIndex) = brwWebBrowser.LocationURL Then
            blnFound = True
            Exit For
        End If
    Next intIndex
    
    mbDontNavigateNow = True
    
    If blnFound Then
        cboAddress.RemoveItem intIndex
    End If
    
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    
    cboAddress.ListIndex = 0
    
    mbDontNavigateNow = False

End Sub

'******************************************************************************
' cboAddress_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

'******************************************************************************
' cboAddress_KeyPress
'
' INPUT:  The ASCII value of the key the user pressed.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

'******************************************************************************
' Form_Resize
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub Form_Resize()
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub

'******************************************************************************
' timTimer_Timer
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Let the user know the form hasn't locked up.
'******************************************************************************
Private Sub timTimer_Timer()
    
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Call ParseCaptionFromCallerName
    Else
        Me.Caption = "Working ..."
    End If

End Sub

'******************************************************************************
' tbToolBar_ButtonClick
'
' INPUT:  The structure of the control the user clicked.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next

    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back":
            
            brwWebBrowser.GoBack
            
        Case "Forward":
            
            brwWebBrowser.GoForward
            
        Case "Refresh":
            
            brwWebBrowser.Refresh
        
        Case "Home":
            
            brwWebBrowser.GoHome
        
        Case "Search":
            
            brwWebBrowser.GoSearch
        
        Case "Stop":
            
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Call ParseCaptionFromCallerName
    
    End Select

End Sub

'******************************************************************************
' ParseCaptionFromCallerName
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Since the help files are not on a web server, the TITLE tag will not
'         be displayed in the browser's caption / title bar. This function
'         displays a readable caption, rather than the help file's name. The
'         help file name format is B17xxxHelpxxx.html, where "B17" must be the
'         prefix, ".html" must be the suffix, and "Help" must appear somewhere
'         in the file name. If the page is not a help file, then the TITLE tag
'         will be displayed.
'******************************************************************************
Private Sub ParseCaptionFromCallerName()
    Dim strTempCaption As String
    Dim strNewCaption As String
    Dim intCaptionLength As Integer
    Dim intCurrentLetterAsciiValue As Integer
    Dim intIndex As Integer

'    strTempCaption = txtPageName.Text
    strTempCaption = brwWebBrowser.LocationName

    If Left(strTempCaption, 3) = "B17" _
    And InStr(1, strTempCaption, "Help") > 0 _
    And Right(strTempCaption, 5) = ".html" Then

        ' We should be loading an .html help file associated with this
        ' emulator.

        intCaptionLength = Len(strTempCaption)

        ' Chop off the 5-character '.html' file extension.

        strTempCaption = Left(strTempCaption, (intCaptionLength - 5))

        intCaptionLength = Len(strTempCaption)

        ' Chop off the 3-character 'B17' file prefix.

        strTempCaption = Right(strTempCaption, (intCaptionLength - 3))
        
        intCaptionLength = Len(strTempCaption)

        ' Now that we have the basic name of the help file, insert spaces
        ' before every capital letter, except the first one.

        strNewCaption = Left(strTempCaption, 1)
        
        For intIndex = 2 To intCaptionLength
            intCurrentLetterAsciiValue = Asc(Mid(strTempCaption, intIndex, 1))
            
            ' ASCII values 65-90 are capital A to Z.
            
            If intCurrentLetterAsciiValue > 64 _
            And intCurrentLetterAsciiValue < 91 Then
                strNewCaption = strNewCaption & " "
            End If
            
            strNewCaption = strNewCaption & Mid(strTempCaption, intIndex, 1)
        Next intIndex
    
        Me.Caption = strNewCaption
    Else
        ' This is not an emulator help file.
    
        Me.Caption = brwWebBrowser.LocationName
    End If

End Sub

