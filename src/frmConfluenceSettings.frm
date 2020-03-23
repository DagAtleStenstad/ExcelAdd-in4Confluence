VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfluenceSettings 
   Caption         =   "Settings"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   OleObjectBlob   =   "frmConfluenceSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfluenceSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
     
    lblConfluenceUrl.ForeColor = vbBlack
    lblConfluenceUsername.ForeColor = vbBlack
    lblConfluencePassword.ForeColor = vbBlack
    
    If txtConfluenceUrl = vbNullString Then
        lblConfluenceUrl.ForeColor = vbRed
        txtConfluenceUrl.SetFocus
        Exit Sub
    End If
    
    If txtConfluenceUsername = vbNullString Then
        lblConfluenceUsername.ForeColor = vbRed
        txtConfluenceUsername.SetFocus
        Exit Sub
    End If

    If txtConfluencePassword = vbNullString Then
        lblConfluencePassword.ForeColor = vbRed
        txtConfluencePassword.SetFocus
        Exit Sub
    End If
    
    SaveSetting "ExcelAddIn4Confluence", "Settings", "Confluence_url", txtConfluenceUrl
    SaveSetting "ExcelAddIn4Confluence", "Settings", "Confluence_username", txtConfluenceUsername
    SaveSetting "ExcelAddIn4Confluence", "Settings", "Confluence_password", txtConfluencePassword
    SaveSetting "ExcelAddIn4Confluence", "Settings", "Confluence_remember_password", chkRemember
    
    Unload Me

End Sub

Private Sub lblExcelAddin4Confluence_Click()
    ActiveWorkbook.FollowHyperlink "https://bitbucket.org/Stenstad/exceladd-in4confluence/src/master/"
End Sub

Private Sub UserForm_Initialize()

    txtConfluenceUrl = GetSetting("ExcelAddIn4Confluence", "Settings", "Confluence_url")
    txtConfluenceUsername = GetSetting("ExcelAddIn4Confluence", "Settings", "Confluence_username")
    
    If GetSetting("ExcelAddIn4Confluence", "Settings", "Confluence_remember_password") = "True" Then
       chkRemember.value = "True"
       txtConfluencePassword = GetSetting("ExcelAddin4Confluence", "Settings", "Confluence_password")
    End If
    
End Sub
