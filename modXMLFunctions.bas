Attribute VB_Name = "modXMLFunctions"
  Private msXMLType As String   ' "attribute" or "value"
  Private msXMLStr As String    ' This is either the attribute name or Value name we are getting the value for
  Private msXMLAttrName As String
  'Public Const gcCLIENTSTATUS_PROSPECT As String = "Prospect"
 
 Private Function HandleTree(SearchedNodeName As String, Node As MSXML2.IXMLDOMNode) As String
  Dim ReturnString As String
  Dim XMLChild As IXMLDOMNode

  If SearchedNodeName = "" Then
          If msXMLType = "attribute" Then
            ReturnString = GetAttributeValue(Node)
          ElseIf msXMLType = "value" Then
            ReturnString = GetFieldValue(Node)
          End If

  Else
    For Each XMLChild In Node.childNodes
      If XMLChild.nodeName = SearchedNodeName Then
          ReturnString = HandleTree(GetNextParsedValue, XMLChild)
      End If
    Next
  End If

  HandleTree = ReturnString
End Function
Public Function HandleTree2(SearchedNodeName As String, Node As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode
  Dim ReturnNode As MSXML2.IXMLDOMNode
  Dim XMLChild As IXMLDOMNode

  If SearchedNodeName = "" Then
            Set ReturnNode = Node
          
  Else
    For Each XMLChild In Node.childNodes
      If XMLChild.nodeName = SearchedNodeName Then
          Set ReturnNode = HandleTree2(GetNextParsedValue, XMLChild)
      End If
    Next
  End If

  Set HandleTree2 = ReturnNode
End Function
Private Function GetAttributeValue(Node As MSXML2.IXMLDOMNode) As String
'This function gets the attribute named in XMLAttrName from thwe current node
  Dim x As Integer
  On Error GoTo GetAttributeValueErr
  If (Node.nodeType = 1) Then
    If (Node.Attributes.length > 0) Then
         GetAttributeValue = Node.Attributes.getNamedItem(msXMLAttrName).nodeTypedValue
    End If
  End If
GetAttributeValueExit:
  Exit Function

GetAttributeValueErr:
  GetAttributeValue = ""
  GoTo GetAttributeValueExit
  
End Function
Private Function GetFieldValue(Node As MSXML2.IXMLDOMNode) As String
  'This function displays the value within the current node.
  GetFieldValue = Node.nodeTypedValue
  
End Function
Public Function GetNextParsedValue() As String
'This Function parses the XMLStr variable to get the next node we are looking for.
  Dim ParseLoc As Integer
  Dim returnstr As String
  
    ParseLoc = InStr(1, msXMLStr, " ")
    If ParseLoc = 0 Then
      ParseLoc = Len(msXMLStr)
      returnstr = msXMLStr
    Else
      returnstr = Left(msXMLStr, ParseLoc - 1)
    End If
    msXMLStr = Right(msXMLStr, Len(msXMLStr) - ParseLoc)
    GetNextParsedValue = returnstr
End Function
Private Function GrabAttributeName() As String
'When the function is called there is a Attribute name
'stored at the end of some of the node paths.  This function
'grabs the last thing off the end of the string.
  Dim ParseLoc As Integer
  Dim returnstr As String
  
    ParseLoc = InStrRev(msXMLStr, "|")
    If ParseLoc = 0 Then
      ParseLoc = Len(msXMLStr)
      returnstr = msXMLStr
    Else
      returnstr = Right(msXMLStr, Len(msXMLStr) - ParseLoc)
    End If
    msXMLStr = Left(msXMLStr, ParseLoc - 1)
    GrabAttributeName = returnstr
End Function
Public Function GetXMLAttribute(Name As String, lnode As MSXML2.IXMLDOMNode) As String
  msXMLType = "attribute"
  msXMLStr = ""
  'msXMLAttrName = GrabAttributeName
  msXMLAttrName = Name
  GetXMLAttribute = HandleTree(GetNextParsedValue(), lnode)
  
End Function
Public Function GetXMLValue(Name As String, lnode As MSXML2.IXMLDOMNode) As String
  msXMLType = "value"
  msXMLStr = Name
  GetXMLValue = HandleTree(GetNextParsedValue(), lnode)
End Function
Public Sub SetXMLStr(str As String)
 msXMLStr = str
End Sub

