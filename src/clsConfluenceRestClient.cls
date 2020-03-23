VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfluenceRestClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private confluenceUrl As String
Private confluenceUsr As String
Private confluencePwd As String

Private Sub Class_initialize()

    If confluenceUrl = vbNullString Or confluenceUsr = vbNullString Or confluencePwd = vbNullString Then
        confluenceUrl = GetSetting("ExcelAddIn4Confluence", "Settings", "Confluence_url")
        confluenceUsr = GetSetting("ExcelAddIn4Confluence", "Settings", "Confluence_username")
        confluencePwd = GetSetting("ExcelAddIn4Confluence", "Settings", "Confluence_password")
    End If
      
    If confluenceUrl = vbNullString Or confluenceUsr = vbNullString Or confluencePwd = vbNullString Then
        frmConfluenceSettings.Show
    End If
        
End Sub

Private Function ConfluenceService(methods As String, url As String, Optional sData As Variant) As String
    
    Dim confluenceClient As Object
    Dim confluenceResponse As String

    Set confluenceClient = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With confluenceClient
        .Open methods, confluenceUrl & url, False
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(confluenceUsr & ":" & confluencePwd)
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .Send sData

        confluenceResponse = .ResponseText
        
    End With
        
    Set confluenceClient = Nothing
    
    ConfluenceService = confluenceResponse
    
End Function

Public Function GetConfluenceContent(confluenceID As String) As clsConfluenceContent

    Dim confluenceResponse As String
    Dim json As Object
    Dim confluenceContent As New clsConfluenceContent
    
    '//Cache ConfluenceContentResponse
    If confluenceContentCache Is Nothing Then Set confluenceContentCache = CreateObject("Scripting.Dictionary")
    
    If Not confluenceContentCache.Exists(confluenceID) Then
        confluenceContentCache.Add confluenceID, ConfluenceService("GET", "rest/api/content/" & confluenceID & "?expand=body.storage,version")
    End If

    Set json = JsonConverter.ParseJson(confluenceContentCache(confluenceID))
    
    With confluenceContent
        .ID = json("id")
        .confluenceType = json("type")
        .title = json("title")
        .version = json("version")("number")
        .body = json("body")("storage")("value")
    End With
    
    Set GetConfluenceContent = confluenceContent
    
End Function

Public Sub UpdateConfluenceContent(confluenceID As String, confluenceType As String, title As String, version As Integer, body As String)

   Dim data As String
      
   data = "{   ""type"":""" & confluenceType & """," & _
          "    ""title"":""" & title & """," & _
          "    ""body"":{""storage"":{""value"":" & JsonConverter.ConvertToJson(body) & ",""representation"":""storage""}}," & _
          "    ""version"":{""number"":" & version & "}}"
    
   Call ConfluenceService("put", "rest/api/content/" & confluenceID, data)
    
End Sub

Private Function EncodeBase64(Text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(Text, vbFromUnicode)
  
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
  
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
  
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text
  
    Set objNode = Nothing
    Set objXML = Nothing
End Function
