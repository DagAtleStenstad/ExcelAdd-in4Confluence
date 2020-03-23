VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfluenceContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pID As String
Private pConfluenceType As String
Private pTitle As String
Private pVersion As Integer
Private pBody As Variant

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(value As String)
    pID = value
End Property

Public Property Get confluenceType() As String
    confluenceType = pConfluenceType
End Property

Public Property Let confluenceType(value As String)
    pConfluenceType = value
End Property

Public Property Get title() As String
    title = pTitle
End Property

Public Property Let title(value As String)
    pTitle = value
End Property

Public Property Get version() As Integer
    version = pVersion
End Property

Public Property Let version(value As Integer)
    pVersion = value
End Property

Public Property Get body() As String
    body = pBody
End Property

Public Property Let body(value As String)
    pBody = value
End Property
