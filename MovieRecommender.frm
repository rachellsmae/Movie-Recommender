VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Movie Recommender"
   ClientHeight    =   11730
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9380.001
   OleObjectBlob   =   "MovieRecommender.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This connection string connects to the server
Const sConnString = "Driver={SQL Server Native Client 11.0};Server=13.58.23.232;Database=project;Uid=rachel;Pwd=password;Connection Timeout=10;"

Private Sub UserForm_Initialize()
'Add genres to combo box
With ComboBox1
    .AddItem "Action"
    .AddItem "Adventure"
    .AddItem "Animation"
    .AddItem "Biography"
    .AddItem "Comedy"
    .AddItem "Crime"
    .AddItem "Drama"
    .AddItem "Family"
    .AddItem "Fantasy"
    .AddItem "Film-Noir"
    .AddItem "History"
    .AddItem "Horror"
    .AddItem "Music"
    .AddItem "Musical"
    .AddItem "Mystery"
    .AddItem "Romance"
    .AddItem "Sci-Fi"
    .AddItem "Sport"
    .AddItem "Thriller"
    .AddItem "War"
    .AddItem "Western"

End With
End Sub

Private Sub CommandButton1_Click()
    
    'Show message box if genre is not selected
    If ComboBox1 = Null Then
        MsgBox ("Please choose a genre.")
    End If
        
    'Show message box if years are not selected
    If OptionButton1.Value = False And OptionButton2.Value = False Then
        MsgBox ("Please choose a range of years.")
    End If
    
    'Need to make an object to hold the connection and the results (generic)
    Dim adoCN As ADODB.Connection
    Dim Movie1Name As New ADODB.Recordset
    Dim Movie1Details As New ADODB.Recordset
    Dim Movie1Image As New ADODB.Recordset
    Dim Movie2Name As New ADODB.Recordset
    Dim Movie2Details As New ADODB.Recordset
    Dim Movie2Image As New ADODB.Recordset
    Dim Movie3Name As New ADODB.Recordset
    Dim Movie3Details As New ADODB.Recordset
    Dim Movie3Image As New ADODB.Recordset
    
    Dim Movie1NameV As Variant
    Dim Movie1DetailsV As Variant
    Dim Movie2NameV As Variant
    Dim Movie2DetailsV As Variant
    Dim Movie3NameV As Variant
    Dim Movie3DetailsV As Variant

    'Set the lower bound of years
    Dim LowerYear As Long
    Dim GenreOptions As Variant
    
    If OptionButton1.Value = True Then
        LowerYear = 1900
    ElseIf OptionButton2.Value = True Then
        LowerYear = 2000
    End If
    
    'Declare variables to download image to disk, load it, and then delete it
    Dim img As Long
    Dim str1URL As String
    Dim str2URL As String
    Dim str3URL As String
    Dim strFile As String
    strFile = "C:\Temp\Temp.jpg"
        
    'Create connection object and open database (generic)
    Set adoCN = New ADODB.Connection

    On Error GoTo FailedConnection
        adoCN.Open sConnString
    On Error GoTo 0

    On Error GoTo FailedQuery
        'Run query and return 1st ranking movie
        Set Movie1Name = adoCN.Execute("SELECT TOP 1 MovieName, Year " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC")
        Set Movie1Details = adoCN.Execute("SELECT TOP 1 Plot " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC")
        Set Movie1Image = adoCN.Execute("SELECT TOP 1 ImgURL " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC")
        Label3.Caption = Movie1Name.GetString(adClipString, , vbTab, vbCrLf)
        Label6.Caption = Movie1Details.GetString(adClipString, , vbTab, vbCrLf)
        
        str1URL = Movie1Image.GetString(adClipString, , vbTab, vbCrLf)
        ret = URLDownloadToFile(0, str1URL, strFile, 0, 0)
        Image1.Picture = LoadPicture(strFile)
    
    
        'Run query and return 2nd ranking movie
        Set Movie2Name = adoCN.Execute("SELECT MovieName, Year " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC " & _
                                    "OFFSET 1 Rows " & _
                                    "Fetch NEXT 1 Rows ONLY")
        Set Movie2Details = adoCN.Execute("SELECT Plot " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC " & _
                                    "OFFSET 1 Rows " & _
                                    "Fetch NEXT 1 Rows ONLY")
        Set Movie2Image = adoCN.Execute("SELECT ImgURL " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC " & _
                                    "OFFSET 1 Rows " & _
                                    "Fetch NEXT 1 Rows ONLY")
        Label4.Caption = Movie2Name.GetString(adClipString, , vbTab, vbCrLf)
        Label7.Caption = Movie2Details.GetString(adClipString, , vbTab, vbCrLf)
        
        str2URL = Movie2Image.GetString(adClipString, , vbTab, vbCrLf)
        ret = URLDownloadToFile(0, str2URL, strFile, 0, 0)
        Image2.Picture = LoadPicture(strFile)
        
        'Run query and return 3rd ranking movie
        Set Movie3Name = adoCN.Execute("SELECT MovieName, Year " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC " & _
                                    "OFFSET 2 Rows " & _
                                    "Fetch NEXT 1 Rows ONLY")
        Set Movie3Details = adoCN.Execute("SELECT Plot " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC " & _
                                    "OFFSET 2 Rows " & _
                                    "Fetch NEXT 1 Rows ONLY")
        Set Movie3Image = adoCN.Execute("SELECT ImgURL " & _
                                    "FROM Movies " & _
                                    "WHERE Genre LIKE '%" & ComboBox1.Value & "%' AND year BETWEEN " & LowerYear & " AND " & LowerYear + 100 & _
                                    " ORDER BY Ranking ASC " & _
                                    "OFFSET 2 Rows " & _
                                    "Fetch NEXT 1 Rows ONLY")
        Label5.Caption = Movie3Name.GetString(adClipString, , vbTab, vbCrLf)
        Label8.Caption = Movie3Details.GetString(adClipString, , vbTab, vbCrLf)
        
        str3URL = Movie3Image.GetString(adClipString, , vbTab, vbCrLf)
        ret = URLDownloadToFile(0, str3URL, strFile, 0, 0)
        Image3.Picture = LoadPicture(strFile)
        Kill strFile
        

On Error GoTo 0

'Movie1Name.Close
'Movie1Details.Close
'Movie2Name.Close
'Movie2Details.Close
'Movie3Name.Close
'Movie3Details.Close

adoCN.Close
Set adoCN = Nothing
Set rsReturn = Nothing
Exit Sub

FailedConnection:
Debug.Print Error(Err)
MsgBox ("Sorry... could not connect to database" & vbCrLf & "Error Message:  " & Error(Err))
Exit Sub

FailedQuery:
Debug.Print Error(Err)
MsgBox ("Sorry...the server could not execute your query" & vbCrLf & "Error Message: " & Error(Err))
Exit Sub

End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

