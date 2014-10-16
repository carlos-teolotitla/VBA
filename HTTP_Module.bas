Attribute VB_Name = "HTTP_Module"

Option Compare Database

Private Const Alfresco_Environment As String = "cms-qa.cengage.info"

Sub Order_Assignment()
    Dim ISBN13 As String
    Dim Alfresco_Assets As Collection
    ISBN13 = "9780538479158"
    Set Alfresco_Assets = Get_Alfresco_Assets(ISBN13)
    
End Sub

Public Function Get_Alfresco_Assets(ByVal ISBN13 As String) As Collection
    Dim ISBN_Worspace As String
    ISBN_Worspace = Get_ISBN_Workspace(ISBN13)
    Set Get_Alfresco_Assets = Get_Workspace_Siblings(ISBN_Worspace, ISBN13)
End Function

Private Function Get_Workspace_Siblings(ByVal Workspace As String, ByVal ISBN13 As String) As Collection

 Dim response As String
    Dim JSonObj As Object
    Dim result As String
    Dim asset As AlfrescoAsset
    Dim assets As Collection
    
    response = http_Resp("http://" + Alfresco_Environment + "/alfresco/service/slingshot/search?term=PRIMARYPARENT:%22" + Workspace + "%22%20AND%20TYPE:%22cengage:asset%22")
    Set JSonObj = JSON.parse(response)
    Set assets = New Collection
    
    For Each Item In JSonObj("items")
        Set asset = New AlfrescoAsset
        asset.ISBN = ISBN13
        asset.FILE_NAME = Item("name")
        asset.NOTES = Item("description")
        asset.FILE_TYPE = Item("node")("mimetypeDisplayName")
        Set asset.ITEM_TYPES = New Collection
        
        For Each category In Item("node")("properties")("cm:categories")
            asset.ITEM_TYPES.Add category("name")
        Next
        If AllowAsset(asset.ITEM_TYPES) Then
                assets.Add asset
        End If
    Next
    Set Get_Workspace_Siblings = assets
End Function

Private Function AllowAsset(ByVal categories As Collection) As Boolean
    Dim denidedCategories As Collection
    Dim result As Boolean
    result = True
    
    Set denidedCategories = New Collection
    denidedCategories.Add "Bookmap"
    denidedCategories.Add "Artwork"
    denidedCategories.Add "Archive Directory Structure"
    denidedCategories.Add "Readme"
    denidedCategories.Add "Font Set"
    denidedCategories.Add "XML"
    For Each category In categories
        For Each denidedCategory In denidedCategories
            If denidedCategory = category Then
                result = False
            End If
        Next
    Next
    AllowAsset = result
End Function

Private Function Get_ISBN_Workspace(ByVal ISBN13 As String) As String
    Dim response As String
    Dim JSonObj As Object
    Dim result As String
    response = http_Resp("http://" + Alfresco_Environment + "/alfresco/service/slingshot/search?term=cbib:isbn13:" + ISBN13)
    Set JSonObj = JSON.parse(response)
    result = JSonObj("items")(1)("nodeRef")
    Get_ISBN_Workspace = result
     
End Function



Private Function http_Resp(ByVal sReq As String) As String

    Dim byteData() As Byte
    Dim XMLHTTP As Object

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")

    XMLHTTP.Open "GET", sReq, False
    XMLHTTP.send
    byteData = XMLHTTP.responseBody
        
    Set XMLHTTP = Nothing

    http_Resp = StrConv(byteData, vbUnicode)

End Function
