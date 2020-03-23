Attribute VB_Name = "modCore"
Option Explicit

Public activeCellAddress As String
Public excelAddInn4ConfluenceCommand As String
Private confluenceContentCache As Object

Public gclsAppEvents As clsAppEvents                          'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
  
Sub Auto_Open()                                               'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
    Set gclsAppEvents = New clsAppEvents
    Set gclsAppEvents.App = Application
End Sub

Function ConfluenceSettings() As String
    If gclsAppEvents Is Nothing Then Call Auto_Open
    
    activeCellAddress = ActiveCell.Address
    excelAddInn4ConfluenceCommand = "openConfluenceSettingsForm"
End Function

Function ConfluenceLookupTableValue(ConfluncePageID As String, GetColumn As String, WhereColumn As String, value As String) As String
    
    Dim confluenceClient As New clsConfluenceRestClient
    
    Dim confluenceContent As clsConfluenceContent
    Set confluenceContent = confluenceClient.GetConfluenceContent(ConfluncePageID)
    
    Dim getColumnNumber As Integer
    Dim whereColumnNumber As Integer
    
    Dim rowNumber As Integer
    
    getColumnNumber = findColumnNumber(confluenceContent.body, Trim(GetColumn))
    whereColumnNumber = findColumnNumber(confluenceContent.body, Trim(WhereColumn))
    
    rowNumber = findRowNumber(confluenceContent.body, value, whereColumnNumber)

    ConfluenceLookupTableValue = geTableCellValue(confluenceContent.body, getColumnNumber, rowNumber)
      
End Function

Private Function findColumnNumber(htmlSource As String, value As String) As Integer
    
    Dim html As New HTMLDocument
    html.body.innerHTML = htmlSource
    
    Dim columns As Object
    Dim column As HTMLHtmlElement
    Dim columnCounter As Integer
    
    Set columns = html.getElementsByTagName("tr")(0).getElementsByTagName("th")

    For Each column In columns
            
            Debug.Print Trim(column.innerText)
            If Trim(column.innerText) = value Then
                findColumnNumber = columnCounter
                Exit For
            End If
        
            columnCounter = columnCounter + 1
    Next

End Function

Private Function findRowNumber(htmlSource As String, value As String, ByVal column As Integer) As Integer

    Dim html As New HTMLDocument
    html.body.innerHTML = htmlSource
    
    Dim rows As Object
    Dim row As HTMLHtmlElement
    Dim cell As HTMLHtmlElement
    Dim rowCounter As Integer
    
    Set rows = html.getElementsByTagName("tr")
            
           For Each row In rows
            Set cell = row.getElementsByTagName("td")(column)
            
             If Not cell Is Nothing Then
                If InStr(1, cell.innerText, value) Then
                    findRowNumber = rowCounter
                    Exit For
                End If
            End If
            
            rowCounter = rowCounter + 1
        Next
       
End Function

Private Function geTableCellValue(htmlSource As String, columnNumber As Integer, rowNumber As Integer)
 
    Dim html As New HTMLDocument
    html.body.innerHTML = htmlSource
    
    If rowNumber = 0 Then Exit Function
    
    geTableCellValue = html.getElementsByTagName("tr")(rowNumber).getElementsByTagName("td")(columnNumber).innerText

End Function
