Attribute VB_Name = "Module1"
Option Explicit

'Windows API to download image from url to file
Public Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long
    
Sub MovieRecommender()
    UserForm.Show
End Sub


