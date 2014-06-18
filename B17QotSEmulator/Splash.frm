VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   """B-17 Queen of the Skies"" Emulator"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub Form_Load()
    
    Dim ret As Long

    'ret = mciSendString("OPEN sound/great-escape.mp3 Alias Sonido", 0, 0, 0)
    'ret = mciSendString("Play sonido", 0, 0, 0)
            
    ' Pause for dramatic effect. Display, approximately, on first drum beat.
    'Sleep 6000
    picSplash.Picture = LoadPicture(App.Path + "\image\SplashLg.jpg")
    picSplash.Height = ScaleY(picSplash.Picture.Height)
    picSplash.Width = ScaleX(picSplash.Picture.Width)
    frmSplash.Height = picSplash.Height
    frmSplash.Width = picSplash.Width
    Load frmMainMenu
End Sub

Private Sub picSplash_Click()
    Dim ret As Long
    
    frmMainMenu.Show
    'ret = mciSendString("CLOSE Sonido", 0, 0, 0)
    Unload Me
End Sub
