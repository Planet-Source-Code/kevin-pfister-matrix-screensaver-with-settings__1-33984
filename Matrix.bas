Attribute VB_Name = "ModMatrix"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Sub Main()
  Select Case Left$(UCase$(Command$), 2)
   Case "/A"         'change password
    'no password support (yet)
    'so why not show about dialog?
    'About
   Case "/C"         'config
    FrmConfig.Show
   Case "/P"         'preview
    'Not Yet
   Case "/S"         'display
    FrmMain.Show
  End Select
  If App.PrevInstance Then End    'Exit if there is a prev version running
End Sub
