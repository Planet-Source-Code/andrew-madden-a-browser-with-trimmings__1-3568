Attribute VB_Name = "Module1"
Public fMainForm As frmMain


Sub Main()
    frmSplash.Show
    
    frmSplash.Refresh
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        End
    End If
    Unload fLogin


    Set fMainForm = New frmMain
    Load fMainForm
    
    fMainForm.Show
    Unload frmSplash
End Sub

