VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 'RBR Added Rotator Code 6/22/2011
Option Explicit
  Private mobjDal As SSnet_DAL.Functions
  Private scenario As String
  Dim mlcoIndividualID As Long
  Private Const ASPRedirect As String = "WorkflowRedirect.asp"
  Private mlErrorCode As Long
  Private mbProcessAllDemos As Boolean
  Private mbAllowBeBacks As Boolean
  Private mbUpdateAllScenarios As Boolean
  Private lsSitesDefaultAddress As String
  Private mbIsUpdate As Boolean
  Private mbDontUpdateBeBackDC As Boolean
  Private gHasNotRunXMLInNote As Boolean
  Private mobjKernel As SSnet_Kernel.Functions
  Private mcVirtualPath As String
  Private mcDefDivRank  As String
  Private mcLeadText As String
  Private mlBuilderID As Long
  Private mlDivisionID As Long
  Private mlTempDivisionID As Long
  Private mlSubdivID As Long
  Private mlUserID As Long
  Private mcINISection As String
  Private mlClientID As Long
  Private mlBeBackID As Long
  Private mlMasterReportID As Long
  Private mlIndividualID As Long
  Private defDemo1 As String
  Private defDemo2 As String
  Private defDemo3 As String
  Private defDemo4 As String
  Private defDemo5 As String
  Private defDemo6 As String
  Private mcAttachments As String
  Private mbSendEmail As Boolean
  Private mcRespCCAddr As String
  Private mcRespBCCAddr As String
  Private mcRespSubj As String
  Private lsLeadType As String
  Private lsSource As String
  Private lsIsWebLead As String
  Private lsStateName As String
  Private lsMarketName As String
  Private lsSendResponse As String
  Private lsContactType As String
  Private lsBuilderNumber As String
  Private lsBuilderName As String
  Private lsCommunityNumber As String
  Private lsCommunityName As String
  Private lsMasterCommunity As String
  Private lsPlanName As String
  Private lsPlanNumber As String
  Private lsSpecNumber As String
  Private lsSpecAddress As String
  Private lsPrice As String
  Private lCountryID As Long
  Private lsTwitterID As String
  Private lsTwitterName As String
  
  Private lsVisitDate As String
  Private lsFirstName As String
  Private lsTitle As String
  Private lsLastName As String
  Private lsEmail As String
  Private lsPhone As String
  Private lsWorkPhone As String
  Private lsMobilePhone As String
  Private lsPager As String
  Private lsFax As String
  Private lsStreetAddress As String
  Private lsStreetAddress2 As String
  Private lsCity As String
  Private lsState As String
  Private lsPostalCode As String
  Private lsCountry As String
  
  Private lscoFirstName As String
  Private lscoTitle As String
  Private lscoLastName As String
  Private lscoEmail As String
  Private lscoPhone As String
  Private lscoWorkPhone As String
  Private lscoMobilePhone As String
  Private lscoPager As String
  Private lscoFax As String
  Private lscoStreetAddress As String
  Private lscoStreetAddress2 As String
  Private lscoCity As String
  Private lscoState As String
  Private lscoPostalCode As String
  Private lscoCountry As String
  Private lscotateID As Long
  
  Private bValidateByEmailInsteadOfaddress As Boolean
  
  
  Private lIndividualID As Long
  Private lsIPAddress As String
  Private lsPrefPriceLow As String
  Private lsPrefPricehigh As String
  Private lsFinancing As String
  Private lsReasonForBuying As String
  Private lsMoveInDate As String
  Private lsComments As String
  Private lsCountElements As String
  Private lsLeadDetails As String
  Private lsStatusName As String
  Private lsUserName As String
  Private lsRank As String
  
  Private mrsDefaults As ADODB.Recordset
  Private QualNode As MSXML2.IXMLDOMNode
  
  
  Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
  Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  
  
Public Function ImportXMLFileWS(ByVal xmlString As String, INIFile As String, ByRef lIndividualID As Long, ByRef ErrorCode As Long) As String
  ' This routine accepts a filename representing an XML data transfer. It then gets this
  ' file, processes it and then returns any error messages as the return value for the
  ' routine. If all went well, the return will be blank thus signifying success.
  
  Dim XMLFile As Integer
  Dim XMLFileStr As String
  Dim NextLine As String
  Dim xmlDoc As New MSXML2.DOMDocument
  Dim XMLChild As MSXML2.IXMLDOMNode
  Dim XMLChild2 As MSXML2.IXMLDOMNode
  Dim cError As String
  Dim cErr As String
  Dim BuilderNode As MSXML2.IXMLDOMNode
  Dim x As String
  Dim SalesCenterImportNode As MSXML2.IXMLDOMNode
  
  On Error GoTo ProcessError

  

  Set xmlDoc = New MSXML2.DOMDocument
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  If xmlDoc.loadXML(xmlString) = False Then Err.Raise 100, , "Xml File did not parse please correct the problem and try again."

  ' Instantiate the kernel and DAL components.
  Set mobjKernel = New SSnet_Kernel.Functions
  Set mobjDal = New SSnet_DAL.Functions
  
  cError = mobjDal.SetINIFile(INIFile)
  If cErr <> "" Then cError = cError & vbCrLf & cErr
  
  'SetXMLStr ("BHILeads")
  Set SalesCenterImportNode = xmlDoc.childNodes.Item(1)
  
  'Set SalesCenterImportNode = HandleTree2(GetNextParsedValue(), xmlDoc)
 
  
  For Each XMLChild In SalesCenterImportNode.childNodes
    If XMLChild.nodeName = "Lead" Then
        Set XMLChild2 = XMLChild.cloneNode(True)
        mcLeadText = DumpLeadFile(XMLChild2)

        
        Call Init 'Initialize Data
        cErr = HandleLead(XMLChild)
        cErr = ProcessData 'Put Data Into DB
        If cErr <> "" Then cError = cError & vbCrLf & cErr
    End If
  Next


ProcessExit:
  ImportXMLFileWS = cError
  Set xmlDoc = Nothing
  ErrorCode = mlErrorCode
  Set XMLChild = Nothing
  Set XMLChild2 = Nothing
  Set BuilderNode = Nothing
  Set SalesCenterImportNode = Nothing
  Exit Function
  
ProcessError:
  cError = Err.Description
   mlErrorCode = 300
  Resume ProcessExit
End Function

Public Function ImportXMLFile(ByVal cfname As String, cINISection As String) As String
  ' This routine accepts a filename representing an XML data transfer. It then gets this
  ' file, processes it and then returns any error messages as the return value for the
  ' routine. If all went well, the return will be blank thus signifying success.
  
  Dim XMLFile As Integer
  Dim XMLFileStr As String
  Dim NextLine As String
  Dim xmlDoc As New MSXML2.DOMDocument
  Dim XMLChild As MSXML2.IXMLDOMNode
  Dim XMLChild2 As MSXML2.IXMLDOMNode
  Dim cError As String
  Dim cErr As String
  Dim BuilderNode As MSXML2.IXMLDOMNode
  Dim x As String
  Dim SalesCenterImportNode As MSXML2.IXMLDOMNode
    
  On Error GoTo ProcessError

  

  Set xmlDoc = New MSXML2.DOMDocument
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.Load (cfname)
  
  'Instantiate the kernel and DAL components.
  Set mobjKernel = New SSnet_Kernel.Functions
  Set mobjDal = New SSnet_DAL.Functions
  
  
  'SetXMLStr ("BHILeads")
  Set SalesCenterImportNode = xmlDoc.childNodes.Item(1)
  
  'Set SalesCenterImportNode = HandleTree2(GetNextParsedValue(), xmlDoc)
 
  
  For Each XMLChild In SalesCenterImportNode.childNodes
    If XMLChild.nodeName = "Lead" Then
        Set XMLChild2 = XMLChild.cloneNode(True)
        mcLeadText = DumpLeadFile(XMLChild2)

        Call Init 'Initialize Data
        cErr = HandleLead(XMLChild)
        cErr = ProcessData 'Put Data Into DB
        If cErr <> "" Then cError = cError & vbCrLf & cErr
    End If
  Next


ProcessExit:
  ImportXMLFile = cError
  Set xmlDoc = Nothing
  Set XMLChild = Nothing
  Set XMLChild2 = Nothing
  Set BuilderNode = Nothing
  Set SalesCenterImportNode = Nothing
  Exit Function
  
ProcessError:
  cError = cfname & " - " & Err.Description
  Resume ProcessExit
End Function
Private Function HandleLead(Node As MSXML2.IXMLDOMNode) As String

  Dim XMLChild As MSXML2.IXMLDOMNode
  Dim rsBuilder As ADODB.Recordset
  Dim cErr As String
  Dim cError As String
  On Error GoTo HandleLeadError

  
  lsSource = GetXMLAttribute("Source", Node)
  lsLeadType = GetXMLAttribute("LeadType", Node)
 
  For Each XMLChild In Node.childNodes
    If XMLChild.nodeName = "Contact" Then
        cErr = HandleContact(XMLChild)
        If cErr <> "" Then cError = cError & vbCrLf & cErr
    End If
  Next
  For Each XMLChild In Node.childNodes
    If XMLChild.nodeName = "Qualifications" Then
        cErr = HandleQualification(XMLChild)
        If cErr <> "" Then cError = cError & vbCrLf & cErr
    End If
  Next
  For Each XMLChild In Node.childNodes
    If XMLChild.nodeName = "PropertyInterest" Then
        cErr = HandlePropertyInterest(XMLChild)
        If cErr <> "" Then cError = cError & vbCrLf & cErr
    End If
  Next
  For Each XMLChild In Node.childNodes
    If XMLChild.nodeName = "Telemetry" Then
        cErr = HandleTelemetry(XMLChild)
        If cErr <> "" Then cError = cError & vbCrLf & cErr
    End If
  Next
 

HandleLeadExit:
  On Error Resume Next
  HandleLead = cError
  Set XMLChild = Nothing
  Set rsBuilder = Nothing
  Exit Function

HandleLeadError:
  mlErrorCode = 201 'Error Occured while processing the XML Lead Error is unknown General Error
  cError = Err.Description
  GoTo HandleLeadExit

End Function

Private Function HandleContact(Node As MSXML2.IXMLDOMNode) As String

  Dim bMatchFound As Boolean
  Dim vFields As Variant
  Dim vFilter As Variant
  Dim XMLChild As MSXML2.IXMLDOMNode
  Dim cErr As String
  Dim cError As String
  
  
  On Error GoTo HandleContactError
  
   lsFirstName = GetXMLAttribute("FirstName", Node)
   lsTitle = GetXMLAttribute("Title", Node)
   lsLastName = GetXMLAttribute("LastName", Node)
   lsEmail = GetXMLAttribute("Email", Node)
   lsPhone = GetXMLAttribute("Phone", Node)
   lsWorkPhone = GetXMLAttribute("WorkPhone", Node)
   lsFax = GetXMLAttribute("Fax", Node)
   lsPager = GetXMLAttribute("Pager", Node)
   lsMobilePhone = GetXMLAttribute("MobilePhone", Node)
   lsStreetAddress = GetXMLAttribute("StreetAddress", Node)
   lsStreetAddress2 = GetXMLAttribute("StreetAddress2", Node)
   lsCity = GetXMLAttribute("City", Node)
   lsState = GetXMLAttribute("State", Node)
   lsPostalCode = GetXMLAttribute("PostalCode", Node)
   lsCountry = GetXMLAttribute("Country", Node)
   lsVisitDate = GetXMLAttribute("VisitDate", Node)
   lsIPAddress = GetXMLAttribute("IPAddress", Node)
   lsContactType = GetXMLAttribute("ContactType", Node)
   lsSendResponse = GetXMLAttribute("SendResponse", Node)
   lsIsWebLead = GetXMLAttribute("IsWebLead", Node)
   lsUserName = GetXMLAttribute("UserName", Node)
   lsTwitterID = GetXMLAttribute("TwitterID", Node)
   lsTwitterName = GetXMLAttribute("TwitterName", Node)
   lsRank = GetXMLAttribute("Rank", Node)
   Dim lsAllowBeBacks As String
   Dim lsProcessAllScenarios As String
   Dim lsUpdateAllDemographics As String
   Dim lsDontUpdateBeBackDC As String
   Dim lsValidateByEmailInsteadOfaddress As String
  
   lsSitesDefaultAddress = GetXMLAttribute("SitesDefaultAddress", Node)
   
   lsValidateByEmailInsteadOfaddress = GetXMLAttribute("ValidateByEmail", Node)
   If UCase(lsValidateByEmailInsteadOfaddress) <> "YES" Then
        bValidateByEmailInsteadOfaddress = False
    Else
         bValidateByEmailInsteadOfaddress = True
   End If
     
   lsDontUpdateBeBackDC = GetXMLAttribute("DontUpdateBeBackDemos", Node)
   If UCase(lsDontUpdateBeBackDC) <> "YES" Then
        mbDontUpdateBeBackDC = False
    Else
         mbDontUpdateBeBackDC = True
   End If
   
   lsAllowBeBacks = GetXMLAttribute("AllowBeBacks", Node)
   If UCase(lsAllowBeBacks) <> "NO" Then
        mbAllowBeBacks = True
    Else
         mbAllowBeBacks = False
   End If
   
   lsProcessAllScenarios = GetXMLAttribute("ProcessAllScenarios", Node)
   If UCase(lsProcessAllScenarios) = "YES" Then
        mbUpdateAllScenarios = True
    Else
        mbUpdateAllScenarios = False
   End If
   
   lsUpdateAllDemographics = GetXMLAttribute("UpdateAllDemographics", Node)
   If UCase(lsUpdateAllDemographics) = "YES" Then
        mbProcessAllDemos = True
    Else
        mbProcessAllDemos = False
        
   End If
  
    
   
   
   
   lscoFirstName = GetXMLAttribute("CoFirstName", Node)
   lscoTitle = GetXMLAttribute("CoTitle", Node)
   lscoLastName = GetXMLAttribute("CoLastName", Node)
   lscoEmail = GetXMLAttribute("CoEmail", Node)
   lscoPhone = GetXMLAttribute("CoPhone", Node)
   lscoWorkPhone = GetXMLAttribute("CoWorkPhone", Node)
   lscoFax = GetXMLAttribute("CoFax", Node)
   lscoPager = GetXMLAttribute("CoPager", Node)
   lscoMobilePhone = GetXMLAttribute("CoMobilePhone", Node)
   lscoStreetAddress = GetXMLAttribute("CoStreetAddress", Node)
   lscoStreetAddress2 = GetXMLAttribute("CoStreetAddress2", Node)
   lscoCity = GetXMLAttribute("CoCity", Node)
   lscoState = GetXMLAttribute("CoState", Node)
   lscoPostalCode = GetXMLAttribute("CoPostalCode", Node)
   lscoCountry = GetXMLAttribute("CoCountry", Node)
   
   
   
HandleContactExit:
  On Error Resume Next
  HandleContact = cError
  Set XMLChild = Nothing
  Exit Function

HandleContactError:
  mlErrorCode = 202 'Error Occured while processing the XML Lead Error is unknown but Occured in the Handle Contact Section
  cError = "HandleContact: ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  GoTo HandleContactExit
End Function


Private Function HandleQualification(Node As MSXML2.IXMLDOMNode) As String
  Dim cErr As String
  Dim cError As String
  Dim vFields As Variant
  Dim vFilter As Variant

  On Error GoTo HandleQualificationError
  
  Set QualNode = Node.cloneNode(False)
  

HandleQualificationExit:
  On Error Resume Next
  HandleQualification = cError
  Exit Function

HandleQualificationError:
  mlErrorCode = 203 'Error Occured while processing the XML Lead Error is unknown but Occured in the Handle Qualification Section
  cError = "HandleQualification: ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  GoTo HandleQualificationExit
End Function



Private Function HandlePropertyInterest(Node As MSXML2.IXMLDOMNode) As String

  Dim vFields As Variant
  Dim vFilter As Variant
  Dim XMLChild As MSXML2.IXMLDOMNode
  Dim cErr As String
  Dim cError As String
  On Error GoTo HandlePropertyInterestError
  

  
   lsStateName = GetXMLAttribute("StateName", Node)
   lsMarketName = GetXMLAttribute("MarketName", Node)
   lsBuilderNumber = GetXMLAttribute("BuilderNumber", Node)
   lsBuilderName = GetXMLAttribute("BuilderName", Node)
   lsCommunityNumber = GetXMLAttribute("CommunityNumber", Node)
   lsCommunityName = GetXMLAttribute("CommunityName", Node)
   lsMasterCommunity = GetXMLAttribute("MasterCommunity", Node)
   lsPlanName = GetXMLAttribute("PlanName", Node)
   lsPlanNumber = GetXMLAttribute("PlanNumber", Node)
   lsSpecNumber = GetXMLAttribute("SpecNumber", Node)
   lsSpecAddress = GetXMLAttribute("SpecAddress", Node)
   lsPrice = GetXMLAttribute("Price", Node)
   

   
 
HandlePropertyInterestExit:
  On Error Resume Next
  HandlePropertyInterest = cError
  Set XMLChild = Nothing
  Exit Function

HandlePropertyInterestError:
  mlErrorCode = 204 'Error Occured while processing the XML Lead Error is unknown but Occured in the Handle Property Interest Section
  cError = "HandlePropertyInterest: ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  GoTo HandlePropertyInterestExit
End Function


Private Function HandleTelemetry(Node As MSXML2.IXMLDOMNode) As String

  Dim XMLChild As MSXML2.IXMLDOMNode
  Dim cErr As String
  Dim cError As String
  Dim vFields As Variant
  Dim vFilter As Variant
  On Error GoTo HandleTelemetryError
  
   lsCountElements = GetXMLAttribute("CountElements", Node)
   lsLeadDetails = GetXMLAttribute("LeadDetails", Node)
  
HandleTelemetryExit:
  On Error Resume Next
  HandleTelemetry = cError
  Set XMLChild = Nothing
  Exit Function

HandleTelemetryError:
  mlErrorCode = 205 'Error Occured while processing the XML Lead Error is unknown but Occured in the Handle Telemetry Section
  cError = "HandleTelemetry: ON  Builder:" & lsBuilderName & "  Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  GoTo HandleTelemetryExit
End Function

Private Function GetEleadsData() As String
  Dim cError As String
  Dim vFields As Variant
  Dim vFilter As Variant
  On Error GoTo GetEleadsDataErr
  
  
  Set mrsDefaults = New ADODB.Recordset
  mrsDefaults.CursorLocation = adUseClient
  mrsDefaults.LockType = adLockBatchOptimistic
  vFields = Array("ID", "BuilderID", "ControllingID", "Category", "DefaultID", "DefaultValue", "NumCompare", "DateCompare", "MinRangeValue", "MaxRangeValue", "Note")
  vFilter = Array("BuilderID=" & CStr(mlBuilderID))
  cError = mobjDal.GetData("tbElead", mrsDefaults, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  
   
GetEleadsDataExit:
  On Error Resume Next
  GetEleadsData = cError
  Exit Function

GetEleadsDataErr:
  mlErrorCode = 206 'Error Occured while Retrieving Data from tbElead.  Without This Data Eleads Doesn't Work
  cError = "GetEleadsData: ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  GoTo GetEleadsDataExit
End Function


Private Function InsertClient() As String
  Dim cError As String
  Dim iItem As Long

  Dim rsIndividual As ADODB.Recordset
  Dim rsClient As ADODB.Recordset
  Dim rsMultiSalesAgent As ADODB.Recordset
  Dim rsBuyerNums As ADODB.Recordset
  Dim rsBeBack As ADODB.Recordset
  Dim rsClientStatus As ADODB.Recordset
  Dim rsDC As ADODB.Recordset
  Dim rsMember As ADODB.Recordset
  Dim rsMember2 As ADODB.Recordset
  Dim rsMemberType As ADODB.Recordset
  Dim lIndividualID As Long
  Dim lClientRankID As Long
  Dim lClientStatus As Long
  Dim lNewIndividualID  As Long
  Dim lRecID As Long
  Dim lMemberTypeID As Long
  Dim lMasterPlanID As Long
  Dim lPhasePlanID As Long
  Dim lLotID As Long
  Dim lUserID As Long
  Dim lNewClientID As Long
  Dim bMatchFound As Boolean
  Dim rsLot As ADODB.Recordset
  Dim rsUser As ADODB.Recordset
  Dim rsMasterPlan As ADODB.Recordset
  Dim rsIPAddress As ADODB.Recordset
  Dim rsPhasePlan As ADODB.Recordset
  Dim rsCheckBuyer As ADODB.Recordset
  Dim vFields As Variant
  Dim vFilter As Variant
  Dim lStateID As Long
  Dim lCountryID As Long
  
  On Error GoTo InsertClientErr
  

  
 
 
  
  
  

  lPhasePlanID = 0
  gHasNotRunXMLInNote = True
  
  Set rsPhasePlan = New ADODB.Recordset
  rsPhasePlan.CursorLocation = adUseClient
  rsPhasePlan.LockType = adLockBatchOptimistic
  vFields = Array("ID")
  vFilter = Array("SubdivisionID = '" & mlSubdivID & "'", "FloorPlanCode = '" & lsPlanNumber & "'")
  cError = mobjDal.GetData("vuSEPhasePlan", rsPhasePlan, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsPhasePlan.EOF = False Then
    lPhasePlanID = rsPhasePlan("ID")
  End If



  mrsDefaults.Filter = ""
  mrsDefaults.Filter = "Category = 'Rank'"
  If mrsDefaults.RecordCount > 0 Or lsRank <> "" Or mcDefDivRank <> "" Then
      
    Dim cRank As String
    If mcDefDivRank <> "" Then
      cRank = mcDefDivRank
    Else
        If mrsDefaults.RecordCount > 0 Then
            cRank = mrsDefaults("DefaultValue")
        End If
    End If
    
    
    If lsRank <> "" Then
        cRank = lsRank
    End If
    
    Set rsDC = New ADODB.Recordset
    rsDC.CursorLocation = adUseClient
    rsDC.LockType = adLockBatchOptimistic
    vFields = Array("ID")
    vFilter = Array("SystemName = 'Rank'", _
                    "SubdivisionID=" & CStr(mlSubdivID), _
                    "Description='" & mobjDal.CorrectSQLText(cRank) & "'")
    cError = mobjDal.GetData("vuGetSubdivisionDatacodes", rsDC, vFields, vFilter)
    If cError <> "" Then
        mlErrorCode = 214 'Error Occured While Retrieving Rank.  Individual Was Created or Updated but Not Client etc..
        Err.Raise 100, , cError
      End If
    With rsDC
      If .RecordCount = 0 Then
         mlErrorCode = 215 'Unable to determine datacode for 'Rank'
        Err.Raise 100, , "Unable to determine datacode for 'Rank'"
      End If
      lClientRankID = .Fields("ID").Value
      
    End With
  Else
   mlErrorCode = 215 'Unable to determine datacode for 'Rank'
   Err.Raise 100, , "Unable to determine datacode for 'Rank'"
  End If
  

  cError = GetSpecLot(lLotID)
  If cError <> "" Then Err.Raise 100, , cError
  
  If lsFirstName = "" Then lsFirstName = lsEmail
  If lsLastName = "" Then lsLastName = lsEmail
  If lsStreetAddress = "" Then lsStreetAddress = lsEmail

  
  
  lIndividualID = 0   ' Default to indicate new individual.
  cError = CheckIndividualUnique(lsFirstName, _
                                        lsLastName, _
                                        lsStreetAddress, _
                                        lsEmail, _
                                        lIndividualID, _
                                        bMatchFound)
   
    If bMatchFound = False Then
      Set rsIndividual = New ADODB.Recordset
      rsIndividual.CursorLocation = adUseClient
      rsIndividual.LockType = adLockBatchOptimistic
      vFields = Array("ID", "BuilderID", "Title", "FirstName", "MI", "LastName", "JobPosition", _
                      "Address1", "Address2", "City", "State", "Country", "CountryID", "StateID", "Zip", "HomePhone", _
                      "Fax", "WorkPhone", "WorkExt", "MobilePhone", "Pager", "Email", "TwitterID", "TwitterName")
      vFilter = Array("ID=0")
      cError = mobjDal.GetData("Individual", rsIndividual, vFields, vFilter)
      
      
      If cError <> "" Then
        mlErrorCode = 213 'Error Occured while Retrieving Individual Data.  from tbIndividual. Lead was not Processed
        Err.Raise 100, , cError
      End If
      rsIndividual.AddNew
      
      If mlBuilderID <> 0 Then rsIndividual("BuilderID") = mlBuilderID
      If Len(lsState) > 2 Then lsState = ""
      If lsFirstName <> "" Then
        cError = SetFieldValue(rsIndividual("FirstName"), lsFirstName, False, True, False, False, "FirstName and Email")
        If cError <> "" Then Err.Raise 100, , cError
        'rsIndividual("FirstName") = lsFirstName
      Else
        cError = SetFieldValue(rsIndividual("FirstName"), lsEmail, False, True, False, False, "FirstName and Email")
        If cError <> "" Then Err.Raise 100, , cError
        
      End If
      cError = SetFieldValue(rsIndividual("Title"), lsTitle)
      If cError <> "" Then Err.Raise 100, , cError
      
      If lsLastName <> "" Then
        'rsIndividual("LastName") = lsLastName
        cError = SetFieldValue(rsIndividual("LastName"), lsLastName, False, True, False, False, "LastName and Email")
        If cError <> "" Then Err.Raise 100, , cError
      
      Else
        cError = SetFieldValue(rsIndividual("LastName"), lsEmail, False, True, False, False, "LastName and Email")
        If cError <> "" Then Err.Raise 100, , cError

      End If
      cError = SetFieldValue(rsIndividual("Email"), lsEmail)
      If cError <> "" Then Err.Raise 100, , cError
      
      
      cError = SetFieldValue(rsIndividual("TwitterID"), lsTwitterID)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsIndividual("TwitterName"), lsTwitterName)
      If cError <> "" Then Err.Raise 100, , cError
      
      
      cError = SetFieldValue(rsIndividual("HomePhone"), lsPhone)
      If cError <> "" Then Err.Raise 100, , cError
      
      If lsWorkPhone <> "" Then
        cError = SetFieldValue(rsIndividual("WorkPhone"), lsWorkPhone)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      cError = SetFieldValue(rsIndividual("Fax"), lsFax)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsIndividual("MobilePhone"), lsMobilePhone)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsIndividual("Pager"), lsPager)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsIndividual("Address2"), lsStreetAddress2)
      If cError <> "" Then Err.Raise 100, , cError
      
      
      If lsStreetAddress <> "" Then
         cError = SetFieldValue(rsIndividual("Address1"), lsStreetAddress, False, True, False, False, "Address and Email")
         If cError <> "" Then Err.Raise 100, , cError
      Else
         cError = SetFieldValue(rsIndividual("Address1"), lsEmail, False, True, False, False, "Address and Email")
         If cError <> "" Then Err.Raise 100, , cError
      End If
      
     
       
       cError = SetFieldValue(rsIndividual("City"), lsCity)
       If cError <> "" Then Err.Raise 100, , cError
    
        If UCase(lsCountry) = "US" Then lsCountry = "USA"
        lCountryID = GetCountryID(lsCountry)
        cError = SetFieldValue(rsIndividual("Country"), lsCountry)
        If cError <> "" Then Err.Raise 100, , cError
        If lCountryID > 0 Then
          cError = SetFieldValue(rsIndividual("CountryID"), CStr(lCountryID))
          If cError <> "" Then Err.Raise 100, , cError
        End If
    
       lStateID = GetStateID(lsState, lCountryID)
       If lStateID > 0 Then
        cError = SetFieldValue(rsIndividual("StateID"), CStr(lStateID))
        If cError <> "" Then Err.Raise 100, , cError
       End If
       
       cError = SetFieldValue(rsIndividual("State"), lsState)
       If cError <> "" Then Err.Raise 100, , cError
      
       cError = SetFieldValue(rsIndividual("Zip"), lsPostalCode)
       If cError <> "" Then Err.Raise 100, , cError
      
      
     
      rsIndividual.Update
    Else
      Set rsIndividual = New ADODB.Recordset
      rsIndividual.CursorLocation = adUseClient
      rsIndividual.LockType = adLockBatchOptimistic
      vFields = Array("ID", "BuilderID", "Title", "FirstName", "MI", "LastName", "JobPosition", _
                      "Address1", "Address2", "City", "State", "Country", "CountryID", "StateID", "Zip", "HomePhone", _
                      "Fax", "WorkPhone", "WorkExt", "MobilePhone", "Pager", "Email")
      vFilter = Array("ID=" & CStr(lIndividualID))
      cError = mobjDal.GetData("Individual", rsIndividual, vFields, vFilter)
      If cError <> "" Then
        mlErrorCode = 213 'Error Occured while Retrieving Individual Data.  from tbIndividual. Lead was not Processed
        Err.Raise 100, , cError
      End If
     
      If lsFirstName <> "" Then
        cError = SetFieldValue(rsIndividual("FirstName"), lsFirstName, False, True)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      If Len(lsState) > 2 Then lsState = ""
      If lsTitle <> "" Then
        cError = SetFieldValue(rsIndividual("Title"), lsTitle)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lsLastName <> "" Then
        'rsIndividual("LastName") = lsLastName
        cError = SetFieldValue(rsIndividual("LastName"), lsLastName)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lsEmail <> "" Then
        'rsIndividual("Email") = lsEmail
        cError = SetFieldValue(rsIndividual("Email"), lsEmail)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lsPhone <> "" Then
        'rsIndividual("HomePhone") = lsPhone
        cError = SetFieldValue(rsIndividual("HomePhone"), lsPhone)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lsStreetAddress <> "" And lsStreetAddress <> lsSitesDefaultAddress Then
        'rsIndividual("Address1") = lsStreetAddress
        cError = SetFieldValue(rsIndividual("Address1"), lsStreetAddress)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lsCity <> "" Then
        'rsIndividual("City") = lsCity
        cError = SetFieldValue(rsIndividual("City"), lsCity)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If UCase(lsCountry) = "US" Then lsCountry = "USA"
        lCountryID = GetCountryID(lsCountry)
        cError = SetFieldValue(rsIndividual("Country"), lsCountry)
        If cError <> "" Then Err.Raise 100, , cError
        If lCountryID > 0 Then
          cError = SetFieldValue(rsIndividual("CountryID"), CStr(lCountryID))
          If cError <> "" Then Err.Raise 100, , cError
        End If
    
       lStateID = GetStateID(lsState, lCountryID)
       If lStateID > 0 Then
        cError = SetFieldValue(rsIndividual("StateID"), CStr(lStateID))
        If cError <> "" Then Err.Raise 100, , cError
       End If
      
      If lsState <> "" Then
        'rsIndividual("State") = lsState
        cError = SetFieldValue(rsIndividual("State"), lsState)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lsPostalCode <> "" Then
        'rsIndividual("Zip") = lsPostalCode
        cError = SetFieldValue(rsIndividual("Zip"), lsPostalCode)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lsCountry <> "" Then
        If UCase(lsCountry) = "US" Then lsCountry = "USA"
        cError = SetFieldValue(rsIndividual("Country"), lsCountry)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lsWorkPhone <> "" Then
        cError = SetFieldValue(rsIndividual("WorkPhone"), lsWorkPhone)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lsFax <> "" Then
        cError = SetFieldValue(rsIndividual("Fax"), lsFax)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lsMobilePhone <> "" Then
        cError = SetFieldValue(rsIndividual("MobilePhone"), lsMobilePhone)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lsPager <> "" Then
        cError = SetFieldValue(rsIndividual("Pager"), lsPager)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lsStreetAddress <> "" Then
        cError = SetFieldValue(rsIndividual("Address2"), lsStreetAddress2)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      rsIndividual.Update
    End If
    
    If bMatchFound = False Then
      cError = mobjDal.CreateData("Individual", rsIndividual, rsIndividual.Bookmark, mlIndividualID)
      lIndividualID = mlIndividualID
      
      
        ' Next, see if we need to save the Buyer's numbers. This will not need to be done
       ' if we found this Individual already in the database.
         ' Yes...
         Set rsBuyerNums = New ADODB.Recordset
         rsBuyerNums.CursorLocation = adUseClient
         rsBuyerNums.LockType = adLockBatchOptimistic
         vFields = Array("ID", "MilitaryStatus", "MonthlyIncome", "OtherMonthlyIncome", _
                         "MonthlyCreditCard", "MonthlyStudentLoan", "MonthlyCarLoan", _
                         "MonthlyChildSupport", "MonthlyAlimony", "MonthlyOther", "IndividualID", "SSN")
         vFilter = Array("IndividualID=0")
         cError = mobjDal.GetData("BuyerInfo", rsBuyerNums, vFields, vFilter)
         If cError <> "" Then
            mlErrorCode = 220 'Error Occured While Getting BuyerInfo Record
            Err.Raise 100, , cError
          End If
         ' Copy the information across.
         With rsBuyerNums
           .AddNew
           For iItem = 1 To rsBuyerNums.Fields.count - 1
             rsBuyerNums.Fields(iItem).Value = rsBuyerNums(iItem).Value
           Next iItem
           .Fields("IndividualID").Value = lIndividualID
           .Fields("MilitaryStatus").Value = "N"
           .Fields("SSN").Value = ""
           .Update
           ' Save it to the database.
           cError = mobjDal.CreateData("BuyerInfo", rsBuyerNums, .Bookmark)
           If cError <> "" Then
             mlErrorCode = 221 'Error Occured While Creating BuyerInfo Record
             Err.Raise 100, , cError
           End If
         End With
       
    Else
      mlErrorCode = 101 'Warning: Individual already Existed.  Updating Current Individual and Moving On with Client.
      cError = mobjDal.UpdateData("Individual", rsIndividual, rsIndividual.Bookmark)
      mlIndividualID = lIndividualID
    End If
    If cError <> "" Then
      mlErrorCode = 211 'Error Occured While Updating or Creating Individual.  Client and Member Will Not Be Processed
      Err.Raise 100, , cError
    End If
     
    
    If lsIPAddress <> "" Then
      Set rsIPAddress = New ADODB.Recordset
      rsIPAddress.CursorLocation = adUseClient
      rsIPAddress.LockType = adLockBatchOptimistic
      vFields = Array("*")
      vFilter = Array("IndividualID = " & mlIndividualID, "IPAddress= '" & lsIPAddress & "'")
      cError = mobjDal.GetData("tbIndividualIP", rsIPAddress, vFields, vFilter)
      If cError <> "" Then Err.Raise 100, , cError
      If rsIPAddress.EOF = True Then
          rsIPAddress.AddNew
          rsIPAddress("IndividualID") = mlIndividualID
          rsIPAddress("IPAddress") = lsIPAddress
          rsIPAddress.Update
          cError = mobjDal.CreateData("tbIndividualIP", rsIPAddress, rsIPAddress.Bookmark)
      End If
    End If
   
   
    
    mrsDefaults.Filter = ""
    mrsDefaults.Filter = "Category = 'Client Status'"
    If mrsDefaults.RecordCount > 0 Then
      'try to get Subdiv from XML
      ' Retrieve the 'Customer' status choices.
      Set rsClientStatus = New ADODB.Recordset
      rsClientStatus.CursorLocation = adUseClient
      rsClientStatus.LockType = adLockBatchOptimistic
      vFields = Array("*")
      vFilter = Array("1=1")
      cError = mobjDal.GetData("ClientStatus", rsClientStatus, vFields, vFilter)
       If cError <> "" Then
        mlErrorCode = 212 'Error Occured While Retriving Client Status.  Individual Was Created or Updated but Not Client etc..
        Err.Raise 100, , cError
      End If
      If lsContactType = "" Then
        rsClientStatus.Filter = "Status='" & mobjDal.CorrectSQLText(mrsDefaults("DefaultValue")) & "'"
        lsStatusName = mrsDefaults("DefaultValue")
      Else
        rsClientStatus.Filter = "Status='" & mobjDal.CorrectSQLText(lsContactType) & "'"
        lsStatusName = lsContactType
        If rsClientStatus.RecordCount = 0 Then
          rsClientStatus.Filter = ""
          rsClientStatus.Filter = "Status='" & mobjDal.CorrectSQLText(mrsDefaults("DefaultValue")) & "'"
          lsStatusName = mrsDefaults("DefaultValue")
        End If
      End If
      
      lClientStatus = rsClientStatus("ID").Value
    Else
      ' Retrieve the 'Customer' status choices.
      Set rsClientStatus = New ADODB.Recordset
      rsClientStatus.CursorLocation = adUseClient
      rsClientStatus.LockType = adLockBatchOptimistic
      vFields = Array("ID")
      vFilter = Array("Status='" & mobjDal.CorrectSQLText(gcCLIENTSTATUS_PROSPECT) & "'")
      cError = mobjDal.GetData("ClientStatus", rsClientStatus, vFields, vFilter)
      If cError <> "" Then
        mlErrorCode = 214 'Error Occured While Retrieving  ClientStatus.  Individual Was Created or Updated but Not Client etc..
        Err.Raise 100, , cError
      End If
      lClientStatus = rsClientStatus("ID").Value
      lsStatusName = gcCLIENTSTATUS_PROSPECT
    End If
    
    
     
    mbIsUpdate = False
    
    ' Check Client Record for uniqueness
    Set rsCheckBuyer = New ADODB.Recordset
    rsCheckBuyer.CursorLocation = adUseClient
    rsCheckBuyer.LockType = adLockBatchOptimistic
    vFields = Array("ID, UserID, SubdivisionID")
    If mbUpdateAllScenarios Then
        vFilter = Array("IndividualID='" & CStr(mlIndividualID) & "'")
        cError = mobjDal.GetData("vuSEClient", rsCheckBuyer, vFields, vFilter)
        If cError <> "" Then
            mlErrorCode = 216 'Error Occured While checking for BeBack Client Was not Created..
            Err.Raise 100, , cError
        End If
        Dim bProcessed1Client As Boolean
        bProcessed1Client = False
        
        Do While rsCheckBuyer.EOF = False
            mbIsUpdate = True
            mlSubdivID = rsCheckBuyer("SubdivisionID")
            bProcessed1Client = True
       
          'buyer exists for individual in this subdiv
          'Err.Raise 100, , "Buyer exists for Individual in this Subdivision"
          
          'I am to create a be-back record and if any of the Selected Demographics are be-back demographics
          'I am to insert them into the be-back record.
          
            mlClientID = rsCheckBuyer("ID")
            mlUserID = rsCheckBuyer("UserID")
            mlBeBackID = 0
          
            If mbProcessAllDemos = True Then
                ProcessDemos (False)
            End If
            
          
             If mbAllowBeBacks = True Then
                mlClientID = rsCheckBuyer("ID")
                mlUserID = rsCheckBuyer("UserID")
                Set rsBeBack = New ADODB.Recordset
                rsBeBack.CursorLocation = adUseClient
                rsBeBack.LockType = adLockBatchOptimistic
                vFields = Array("ID", "ClientID", "VisitDate", "BeBackDemo1ID", "BeBackDemo2ID", "BeBackDemo3ID")
                vFilter = Array("ID=0")
                cError = mobjDal.GetData("tbBeBack", rsBeBack, vFields, vFilter)
                If cError <> "" Then
                mlErrorCode = 217 'Error Occured While Creating BeBack Client Already Existed
                Err.Raise 100, , cError
                End If
                If rsBeBack.EOF = True Then
                    rsBeBack.AddNew
                    rsBeBack("ClientID").Value = mlClientID
                    If IsDate(lsVisitDate) = True Then
                      If Len(lsVisitDate) <= 5 And ConvertDateTimeToDate(CDate(lsVisitDate)) > Now Then
                        rsBeBack("VisitDate").Value = ConvertDateTimeToDate(CDate(Now))
                      Else
                        rsBeBack("VisitDate").Value = ConvertDateTimeToDate(CDate(lsVisitDate))
                      End If
                    Else
                      rsBeBack("VisitDate").Value = ConvertDateTimeToDate(Now)
                    End If
                    rsBeBack("BeBackDemo1ID").Value = Null
                    rsBeBack("BeBackDemo2ID").Value = Null
                    rsBeBack("BeBackDemo3ID").Value = Null
                    rsBeBack.Update
                    ' Save this.
                    
                    cError = mobjDal.CreateData("tbBeBack", rsBeBack, rsBeBack.Bookmark, lRecID)
                    If cError <> "" Then
                      mlErrorCode = 217 'Error Occured While Creating BeBack Client Already Existed
                      Err.Raise 100, , cError
                    End If
                    mlBeBackID = lRecID
                    ProcessDemos (False)
                End If
            End If
            
             ProcessDemos (True)
            
            If lsEmail <> "" Then
              If UCase(lsSendResponse) <> "NO" Then
                cError = SendResponse(lsEmail)
              End If
            End If
            
            
          
         
          
          rsCheckBuyer.MoveNext
          
        Loop
        If bProcessed1Client = True Then GoTo InsertClientExit
        
    Else
        
        If lLotID <> 0 Then
          vFilter = Array("IndividualID='" & CStr(mlIndividualID) & "'", "SubdivisionID = " & CStr(mlSubdivID), "LotID = " & CStr(lLotID))
        Else
          vFilter = Array("IndividualID='" & CStr(mlIndividualID) & "'", "SubdivisionID = " & CStr(mlSubdivID), "LotID is Null")
        End If
   
        cError = mobjDal.GetData("vuSEClient", rsCheckBuyer, vFields, vFilter)
        If cError <> "" Then
            mlErrorCode = 216 'Error Occured While checking for BeBack Client Was not Created..
            Err.Raise 100, , cError
          End If
        If rsCheckBuyer.EOF = False Then
            mbIsUpdate = True
            
          'buyer exists for individual in this subdiv
          'Err.Raise 100, , "Buyer exists for Individual in this Subdivision"
          
          'I am to create a be-back record and if any of the Selected Demographics are be-back demographics
          'I am to insert them into the be-back record.
            If mbProcessAllDemos = True Then
                ProcessDemos (False)
            End If
            
          
            If mbAllowBeBacks = True Then
                mlClientID = rsCheckBuyer("ID")
                mlUserID = rsCheckBuyer("UserID")
                Set rsBeBack = New ADODB.Recordset
                rsBeBack.CursorLocation = adUseClient
                rsBeBack.LockType = adLockBatchOptimistic
                vFields = Array("ID", "ClientID", "VisitDate", "BeBackDemo1ID", "BeBackDemo2ID", "BeBackDemo3ID")
                vFilter = Array("ID=0")
                cError = mobjDal.GetData("tbBeBack", rsBeBack, vFields, vFilter)
                If cError <> "" Then
                mlErrorCode = 217 'Error Occured While Creating BeBack Client Already Existed
                Err.Raise 100, , cError
                End If
                If rsBeBack.EOF = True Then
                    rsBeBack.AddNew
                    rsBeBack("ClientID").Value = mlClientID
                    If IsDate(lsVisitDate) = True Then
                      If Len(lsVisitDate) <= 5 And ConvertDateTimeToDate(CDate(lsVisitDate)) > Now Then
                        rsBeBack("VisitDate").Value = ConvertDateTimeToDate(CDate(Now))
                      Else
                        rsBeBack("VisitDate").Value = ConvertDateTimeToDate(CDate(lsVisitDate))
                      End If
                    Else
                      rsBeBack("VisitDate").Value = ConvertDateTimeToDate(Now)
                    End If
                    rsBeBack("BeBackDemo1ID").Value = Null
                    rsBeBack("BeBackDemo2ID").Value = Null
                    rsBeBack("BeBackDemo3ID").Value = Null
                    rsBeBack.Update
                    ' Save this.
                    
                    cError = mobjDal.CreateData("tbBeBack", rsBeBack, rsBeBack.Bookmark, lRecID)
                    If cError <> "" Then
                      mlErrorCode = 217 'Error Occured While Creating BeBack Client Already Existed
                      Err.Raise 100, , cError
                    End If
                    mlBeBackID = lRecID
                    ProcessDemos (False)
                End If
            End If
            
            ProcessDemos (True)
            
            If lsEmail <> "" Then
              If UCase(lsSendResponse) <> "NO" Then
                cError = SendResponse(lsEmail)
              End If
            End If
            
            
            GoTo InsertClientExit
         
        End If
     End If

    
        ' Client...
      Set rsClient = New ADODB.Recordset
      rsClient.CursorLocation = adUseClient
      rsClient.LockType = adLockBatchOptimistic
      vFields = Array("ID", "ClientStatusID", "UserID", "LotID", "SubdivisionID", "ClientRankID", "Salutation", "VisitDate", "PlanID", "LeadSource")
      vFilter = Array("ID=0")
      cError = mobjDal.GetData("Client", rsClient, vFields, vFilter)
      If cError <> "" Then
        mlErrorCode = 218 'Error Occured While Getting Blank Client Record
        Err.Raise 100, , cError
      End If
    
        
       rsClient.AddNew
      
       rsClient("ClientStatusID").Value = lClientStatus
       rsClient("ClientRankID").Value = lClientRankID
       If UCase(lsIsWebLead) <> "NO" Then
        rsClient("LeadSource").Value = giELEAD
       End If
       If IsDate(lsVisitDate) = True Then
           If Len(lsVisitDate) <= 5 And ConvertDateTimeToDate(CDate(lsVisitDate)) > Now Then
            rsClient("VisitDate").Value = ConvertDateTimeToDate(CDate(Now))
          Else
            rsClient("VisitDate").Value = ConvertDateTimeToDate(CDate(lsVisitDate))
          End If
        Else
          rsClient("VisitDate").Value = ConvertDateTimeToDate(Now)
        End If
       'rsClient("VisitDate").Value = ConvertDateTimeToDate(Now)
       mlUserID = GetUserID(mlSubdivID)
       Set rsUser = New ADODB.Recordset
       rsUser.CursorLocation = adUseClient
       rsUser.LockType = adLockBatchOptimistic
       vFields = Array("UserID")
       vFilter = Array("Active = 1", "UserID= " & mlUserID, "SubdivisionID = " & mlSubdivID)
       cError = mobjDal.GetData("vuGetUsersBySubdivision", rsUser, vFields, vFilter)
       If cError <> "" Then Err.Raise 100, , cError
       If rsUser.EOF = True Then
         mlErrorCode = 207 'Error Occured while Retrieving Lot Data from tbLot in the Insert Individual Function.  Lead was not Processed
         Err.Raise 100, , "User Is Not Active in the Subdivision, the Client will not be created Please re-activate the user or assign another Default User for the Subdivision"
       End If
        
       rsClient("UserID").Value = mlUserID
       rsClient("SubdivisionID").Value = mlSubdivID
       
       'If lPhasePlanID <> 0 Then
       '   rsClient("PlanID").Value = lPhasePlanID
       'End If
       rsClient.Update
      
     
       ' Use the existing Individual.
       lNewIndividualID = lIndividualID
       
       cError = mobjDal.CreateData("Client", rsClient, rsClient.Bookmark, lNewClientID)
       
       If cError <> "" Then
          mlErrorCode = 219 'Error Occured While Creating Client Record Client Was Not created NOr was member
          Err.Raise 100, , cError
        End If
       mlClientID = lNewClientID
       
       If lLotID <> 0 And lNewClientID <> 0 Then
          cError = SetLot(lLotID, lNewClientID)
          If cError <> "" Then Err.Raise 100, , cError
       End If
       
      
     
       
       ' Lastly, as this was a Create operation, we will need to also create a new
       ' record in the Membership datagroup.
       ' Retrieve the member type.
       Set rsMemberType = New ADODB.Recordset
       rsMemberType.CursorLocation = adUseClient
       rsMemberType.LockType = adLockBatchOptimistic
       vFields = Array("ID")
       vFilter = Array("MemberType='Buyer'")
       cError = mobjDal.GetData("MemberType", rsMemberType, vFields, vFilter)
       If cError <> "" Then
         mlErrorCode = 222 'Error Occured While Getting Member Type Data
         Err.Raise 100, , cError
       End If
       lMemberTypeID = rsMemberType("ID").Value
       
       
       ' Retrieve an empty Member structure and populate it with the required values.
       Set rsMember2 = New ADODB.Recordset
       rsMember2.CursorLocation = adUseClient
       rsMember2.LockType = adLockBatchOptimistic
       
       
       
       vFields = Array("*")
       vFilter = Array("ID=0")
       cError = mobjDal.GetData("Member", rsMember2, vFields, vFilter)
       
       If cError <> "" Then
         mlErrorCode = 223 'Error Occured While Getting Blank Member Data
         Err.Raise 100, , cError
       End If
       
       rsMember2.AddNew
       rsMember2.Fields("IndividualID").Value = lNewIndividualID
       rsMember2.Fields("MemberTypeID").Value = lMemberTypeID
       rsMember2.Fields("ClientID").Value = lNewClientID
       
       scenario = GetScenario(lNewIndividualID)
       rsMember2.Fields("Scenario").Value = scenario

       cError = mobjDal.CreateData("tbMember", rsMember2, rsMember2.Bookmark)
       
       If cError <> "" Then Err.Raise 100, , cError


        cError = InsertCoBuyer()


       ' 3/2/04 DSP We need to add the selected sales agent to the multi-agent table.
       Set rsMultiSalesAgent = New ADODB.Recordset
       rsMultiSalesAgent.CursorLocation = adUseClient
       rsMultiSalesAgent.LockType = adLockBatchOptimistic
       vFields = Array("ID", "ClientID", "UserID", "SalePercentage")
       vFilter = Array("ID=0")
       cError = mobjDal.GetData("tbClientSalesAgent", rsMultiSalesAgent, vFields, vFilter)
       If cError <> "" Then Err.Raise 100, , cError
       rsMultiSalesAgent.AddNew
       rsMultiSalesAgent("ClientID").Value = lNewClientID
       rsMultiSalesAgent("UserID").Value = mlUserID
       rsMultiSalesAgent("SalePercentage").Value = 100
       rsMultiSalesAgent.Update
       cError = mobjDal.CreateData("tbClientSalesAgent", rsMultiSalesAgent, rsMultiSalesAgent.Bookmark)
       If cError <> "" Then
         mlErrorCode = 225 'Error Occured While creating ClientSalesAgent Record
         Err.Raise 100, , cError
       End If
       cError = ProcessDemos(True)
       
      If lsEmail <> "" Then
        If UCase(lsSendResponse) <> "NO" Then
            cError = SendResponse(lsEmail)
          End If
      End If
     
      
InsertClientExit:
  On Error Resume Next
  InsertClient = cError
  Set rsIndividual = Nothing
  Set rsClient = Nothing
  Set rsMultiSalesAgent = Nothing
  Set rsBuyerNums = Nothing
  Set rsBeBack = Nothing
  Set rsClientStatus = Nothing
  Set rsDC = Nothing
  Set rsMember = Nothing
  Set rsMember2 = Nothing
  Set rsMemberType = Nothing
  
  Set rsUser = Nothing
  Set rsMasterPlan = Nothing
  Set rsPhasePlan = Nothing
  Set rsCheckBuyer = Nothing
  Exit Function



InsertClientErr:
  cError = "InsertClient: ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  GoTo InsertClientExit
End Function

Private Function InsertCoBuyer() As String

If lscoStreetAddress <> "" Or lscoLastName <> "" Or lscoEmail <> "" Or lscoFirstName <> "" Then
 Dim lcoIndividualID As Long
 Dim bMatchFound As Boolean
 Dim cError As String
 Dim rsCoBuyer As ADODB.Recordset
 Dim vFields As Variant
  Dim vFilter As Variant
 
 

  If lscoFirstName = "" Then lscoFirstName = lscoEmail
  If lscoLastName = "" Then lscoLastName = lscoEmail
  If lscoStreetAddress = "" Then lscoStreetAddress = lscoEmail

  
  
  lcoIndividualID = 0   ' Default to indicate new individual.
  cError = mobjKernel.CheckIndividualUnique(mobjDal, _
                                        lscoFirstName, _
                                        lscoLastName, _
                                        lscoStreetAddress, _
                                        lcoIndividualID, _
                                        bMatchFound)
                                              
    If bMatchFound = False Then
      Set rsCoBuyer = New ADODB.Recordset
      rsCoBuyer.CursorLocation = adUseClient
      rsCoBuyer.LockType = adLockBatchOptimistic
      vFields = Array("ID", "BuilderID", "Title", "FirstName", "MI", "LastName", "JobPosition", _
                      "Address1", "Address2", "City", "State", "Country", "CountryID", "StateID", "Zip", "HomePhone", _
                      "Fax", "WorkPhone", "WorkExt", "MobilePhone", "Pager", "Email")
      vFilter = Array("ID=0")
      cError = mobjDal.GetData("Individual", rsCoBuyer, vFields, vFilter)
      
      
      If cError <> "" Then
        mlErrorCode = 213 'Error Occured while Retrieving Individual Data.  from tbIndividual. Lead was not Processed
        Err.Raise 100, , cError
      End If
      rsCoBuyer.AddNew
      
      If mlBuilderID <> 0 Then rsCoBuyer("BuilderID") = mlBuilderID
      If Len(lscoState) > 2 Then lscoState = ""
      If lscoFirstName <> "" Then
        cError = SetFieldValue(rsCoBuyer("FirstName"), lscoFirstName, False, True, False, False, "FirstName and Email")
        If cError <> "" Then Err.Raise 100, , cError
        'rsCoBuyer("FirstName") = lscoFirstName
      Else
        cError = SetFieldValue(rsCoBuyer("FirstName"), lscoEmail, False, True, False, False, "FirstName and Email")
        If cError <> "" Then Err.Raise 100, , cError
        
      End If
      cError = SetFieldValue(rsCoBuyer("Title"), lscoTitle)
      If cError <> "" Then Err.Raise 100, , cError
      
      If lscoLastName <> "" Then
        'rsCoBuyer("LastName") = lscoLastName
        cError = SetFieldValue(rsCoBuyer("LastName"), lscoLastName, False, True, False, False, "LastName and Email")
        If cError <> "" Then Err.Raise 100, , cError
      
      Else
        cError = SetFieldValue(rsCoBuyer("LastName"), lscoEmail, False, True, False, False, "LastName and Email")
        If cError <> "" Then Err.Raise 100, , cError

      End If
      cError = SetFieldValue(rsCoBuyer("Email"), lscoEmail)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsCoBuyer("HomePhone"), lscoPhone)
      If cError <> "" Then Err.Raise 100, , cError
      
      
      cError = SetFieldValue(rsCoBuyer("WorkPhone"), lscoWorkPhone)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsCoBuyer("Fax"), lscoFax)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsCoBuyer("MobilePhone"), lscoMobilePhone)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsCoBuyer("Pager"), lscoPager)
      If cError <> "" Then Err.Raise 100, , cError
      
      cError = SetFieldValue(rsCoBuyer("Address2"), lscoStreetAddress2)
      If cError <> "" Then Err.Raise 100, , cError
      
      
      If lscoStreetAddress <> "" Then
         cError = SetFieldValue(rsCoBuyer("Address1"), lscoStreetAddress, False, True, False, False, "Address and Email")
         If cError <> "" Then Err.Raise 100, , cError
      Else
         cError = SetFieldValue(rsCoBuyer("Address1"), lscoEmail, False, True, False, False, "Address and Email")
         If cError <> "" Then Err.Raise 100, , cError
      End If
      
     
       
       cError = SetFieldValue(rsCoBuyer("City"), lscoCity)
       If cError <> "" Then Err.Raise 100, , cError
    
        If UCase(lscoCountry) = "US" Then lscoCountry = "USA"
        lCountryID = GetCountryID(lscoCountry)
        cError = SetFieldValue(rsCoBuyer("Country"), lscoCountry)
        If cError <> "" Then Err.Raise 100, , cError
      
        If lCountryID > 0 Then
          cError = SetFieldValue(rsCoBuyer("CountryID"), CStr(lCountryID))
          If cError <> "" Then Err.Raise 100, , cError
        End If
    
       lscotateID = GetStateID(lscoState, lCountryID)
       If lscotateID > 0 Then
        cError = SetFieldValue(rsCoBuyer("StateID"), CStr(lscotateID))
        If cError <> "" Then Err.Raise 100, , cError
       End If
       
       cError = SetFieldValue(rsCoBuyer("State"), lscoState)
       If cError <> "" Then Err.Raise 100, , cError
      
       cError = SetFieldValue(rsCoBuyer("Zip"), lscoPostalCode)
       If cError <> "" Then Err.Raise 100, , cError
      
      
     
      rsCoBuyer.Update
    Else
      Set rsCoBuyer = New ADODB.Recordset
      rsCoBuyer.CursorLocation = adUseClient
      rsCoBuyer.LockType = adLockBatchOptimistic
      vFields = Array("ID", "BuilderID", "Title", "FirstName", "MI", "LastName", "JobPosition", _
                      "Address1", "Address2", "City", "State", "Country", "Zip", "HomePhone", _
                      "Fax", "WorkPhone", "WorkExt", "MobilePhone", "Pager", "Email")
      vFilter = Array("ID=" & CStr(lcoIndividualID))
      cError = mobjDal.GetData("Individual", rsCoBuyer, vFields, vFilter)
      If cError <> "" Then
        mlErrorCode = 213 'Error Occured while Retrieving Individual Data.  from tbIndividual. Lead was not Processed
        Err.Raise 100, , cError
      End If
     
      If lscoFirstName <> "" Then
        cError = SetFieldValue(rsCoBuyer("FirstName"), lscoFirstName, False, True)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      If Len(lscoState) > 2 Then lscoState = ""
      If lscoTitle <> "" Then
        cError = SetFieldValue(rsCoBuyer("Title"), lscoTitle)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lscoLastName <> "" Then
        'rsCoBuyer("LastName") = lscoLastName
        cError = SetFieldValue(rsCoBuyer("LastName"), lscoLastName)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lscoEmail <> "" Then
        'rsCoBuyer("Email") = lscoEmail
        cError = SetFieldValue(rsCoBuyer("Email"), lscoEmail)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lscoPhone <> "" Then
        'rsCoBuyer("HomePhone") = lscoPhone
        cError = SetFieldValue(rsCoBuyer("HomePhone"), lscoPhone)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lscoStreetAddress <> "" Then
        'rsCoBuyer("Address1") = lscoStreetAddress
        cError = SetFieldValue(rsCoBuyer("Address1"), lscoStreetAddress)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lscoCity <> "" Then
        'rsCoBuyer("City") = lscoCity
        cError = SetFieldValue(rsCoBuyer("City"), lscoCity)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lscoState <> "" Then
        'rsCoBuyer("State") = lscoState
        cError = SetFieldValue(rsCoBuyer("State"), lscoState)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lscoPostalCode <> "" Then
        
        cError = SetFieldValue(rsCoBuyer("Zip"), lscoPostalCode)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lscoCountry <> "" Then
        If UCase(lscoCountry) = "US" Then lscoCountry = "USA"
        cError = SetFieldValue(rsCoBuyer("Country"), lscoCountry)
        If cError <> "" Then Err.Raise 100, , cError

      End If
      
      If lscoWorkPhone <> "" Then
        cError = SetFieldValue(rsCoBuyer("WorkPhone"), lscoWorkPhone)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lscoFax <> "" Then
        cError = SetFieldValue(rsCoBuyer("Fax"), lscoFax)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lscoMobilePhone <> "" Then
        cError = SetFieldValue(rsCoBuyer("MobilePhone"), lscoMobilePhone)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lscoPager <> "" Then
        cError = SetFieldValue(rsCoBuyer("Pager"), lscoPager)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      If lscoStreetAddress <> "" Then
        cError = SetFieldValue(rsCoBuyer("Address2"), lscoStreetAddress2)
        If cError <> "" Then Err.Raise 100, , cError
      End If
      
      rsCoBuyer.Update
    End If
    
    If bMatchFound = False Then
      cError = mobjDal.CreateData("Individual", rsCoBuyer, rsCoBuyer.Bookmark, mlcoIndividualID)
      lcoIndividualID = mlcoIndividualID
      
        Dim rsBuyerNums As New ADODB.Recordset
        
        Set rsBuyerNums = New ADODB.Recordset
        rsBuyerNums.CursorLocation = adUseClient
        rsBuyerNums.LockType = adLockBatchOptimistic
        vFields = Array("ID", "MilitaryStatus", "MonthlyIncome", "OtherMonthlyIncome", _
                       "MonthlyCreditCard", "MonthlyStudentLoan", "MonthlyCarLoan", _
                       "MonthlyChildSupport", "MonthlyAlimony", "MonthlyOther", "IndividualID", "SSN")
        vFilter = Array("IndividualID=0")
        cError = mobjDal.GetData("BuyerInfo", rsBuyerNums, vFields, vFilter)
        If cError <> "" Then
          mlErrorCode = 220 'Error Occured While Getting BuyerInfo Record
          Err.Raise 100, , cError
        End If
        ' Copy the information across.
        Dim iItem As Integer
        
        With rsBuyerNums
         .AddNew
         For iItem = 1 To rsBuyerNums.Fields.count - 1
           rsBuyerNums.Fields(iItem).Value = rsBuyerNums(iItem).Value
         Next iItem
         .Fields("IndividualID").Value = mlcoIndividualID
         .Fields("MilitaryStatus").Value = "N"
         .Fields("SSN").Value = ""
         .Update
         ' Save it to the database.
         cError = mobjDal.CreateData("BuyerInfo", rsBuyerNums, .Bookmark)
         If cError <> "" Then
           mlErrorCode = 221 'Error Occured While Creating BuyerInfo Record
           Err.Raise 100, , cError
         End If
        End With
     
    Else
      mlErrorCode = 101 'Warning: Individual already Existed.  Updating Current Individual and Moving On with Client.
      cError = mobjDal.UpdateData("Individual", rsCoBuyer, rsCoBuyer.Bookmark)
      mlcoIndividualID = lcoIndividualID
    End If
    Dim cErr As String
    
    cErr = insertCoMember(mlcoIndividualID, mlClientID, scenario)


End If

End Function


Private Function insertCoMember(indID As Long, clientID As Long, lscenario As String) As String
Dim lMemberTypeID As Long
Dim cError As String
 
 Dim vFields As Variant
  Dim vFilter As Variant
Dim lrsMemberType As ADODB.Recordset
Dim lrsMember2 As ADODB.Recordset
' Lastly, as this was a Create operation, we will need to also create a new
       ' record in the Membership datagroup.
       ' Retrieve the member type.
       Set lrsMemberType = New ADODB.Recordset
       lrsMemberType.CursorLocation = adUseClient
       lrsMemberType.LockType = adLockBatchOptimistic
       vFields = Array("ID")
       vFilter = Array("MemberType='Co-Buyer'")
       cError = mobjDal.GetData("MemberType", lrsMemberType, vFields, vFilter)
       If cError <> "" Then
         mlErrorCode = 222 'Error Occured While Getting Member Type Data
         Err.Raise 100, , cError
       End If
       lMemberTypeID = lrsMemberType("ID").Value
       
       
       ' Retrieve an empty Member structure and populate it with the required values.
       Set lrsMember2 = New ADODB.Recordset
       lrsMember2.CursorLocation = adUseClient
       lrsMember2.LockType = adLockBatchOptimistic
       

 vFields = Array("*")
       vFilter = Array("ID=0")
       cError = mobjDal.GetData("Member", lrsMember2, vFields, vFilter)
       
       If cError <> "" Then
         mlErrorCode = 223 'Error Occured While Getting Blank Member Data
         Err.Raise 100, , cError
       End If
       
       lrsMember2.AddNew
       lrsMember2.Fields("IndividualID").Value = indID
       lrsMember2.Fields("MemberTypeID").Value = lMemberTypeID
       lrsMember2.Fields("ClientID").Value = clientID
    
       lrsMember2.Fields("Scenario").Value = scenario

       cError = mobjDal.CreateData("tbMember", lrsMember2, lrsMember2.Bookmark)
End Function


Private Function SendToSalesAgent() As String
  Dim rsMasterReport As ADODB.Recordset
  Dim rsReportDirectory As ADODB.Recordset
  Dim rsReportRecordset As ADODB.Recordset
  Dim rsMasterPlan As ADODB.Recordset
  Dim rsResponseInfo As ADODB.Recordset
  Dim lMasterPlanID As Long
  Dim lMasterReportID As Long
  Dim cProcessViewName As String
  Dim cProcessFileName As String
  Dim cReportDirectory As String
  Dim cUserEmail As String
  Dim cFromAddr As String
  Dim cToAddr As String
  Dim bSendCopytoAgent As Boolean
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cMailContent As String
  Dim cFromAddress As String
  Dim cError As String
  Dim cScreenandID As String
  Dim cWorkflowURL As String
  Dim WokflowLink As String
  Dim RedirectURL As String
  
mcRespCCAddr = ""
cToAddr = ""

  GoSub GetResponseInfo
  
  If mbSendEmail = True Then
    GoSub GetMasterReportInfo
    GoSub GetMasterReportRecordset
     
    'Get info to send the Login Link through the response
    GoSub GetBuilderWorkflowURL
    If cWorkflowURL <> "" Then
        cScreenandID = PrefilZeros(giCUSTOMER, giCUSTOMER_NOTES) & CStr(mlClientID)
        RedirectURL = cWorkflowURL & ASPRedirect & "?WorkflowData=" & cScreenandID
        WokflowLink = RedirectURL
    Else
        WokflowLink = ""
    End If
        
    cMailContent = MergeEmailContent(rsReportRecordset, cProcessFileName, WokflowLink)
    If bSendCopytoAgent = True Then
        GoSub GetUserEmail
        If cToAddr = "" Then
            cToAddr = cUserEmail
        Else
            mcRespCCAddr = cUserEmail
        End If
        
    End If
    
    'cFromAddress = GetFromAddress()
    If Trim(cFromAddr) = "" Then Err.Raise 100, , "No email address was found to serve as the ‘From’ address.  No email was sent."
    cError = SendEmail(cToAddr, cFromAddr, mcRespSubj, mcRespCCAddr, "", cMailContent, 0)
  End If

SendResponseExit:
  SendToSalesAgent = cError
  Exit Function
  Return
  
GetResponseInfo:
  Set rsResponseInfo = New ADODB.Recordset
  rsResponseInfo.CursorLocation = adUseClient
  rsResponseInfo.LockType = adLockBatchOptimistic
  vFields = Array("*")
  vFilter = Array("ResponseType='A'", "SubdivisionID=" & mlSubdivID)
  cError = mobjDal.GetData("tbEleadAutoResponse", rsResponseInfo, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsResponseInfo.EOF = False Then
    mbSendEmail = "" & rsResponseInfo("SendEmail").Value
    mcAttachments = "" & rsResponseInfo("Attachments").Value
    cFromAddr = "" & rsResponseInfo("FromAddr").Value
    cToAddr = "" & rsResponseInfo("ToAddr").Value
    mcRespSubj = "" & rsResponseInfo("Subject").Value
    bSendCopytoAgent = "" & rsResponseInfo("SendCopy").Value
    mlMasterReportID = "" & rsResponseInfo("MasterReportID").Value
  Else
    mbSendEmail = False
  End If
  Return

GetMasterReportInfo:
  Set rsMasterReport = New ADODB.Recordset
  rsMasterReport.CursorLocation = adUseClient
  rsMasterReport.LockType = adLockBatchOptimistic
  
  vFields = Array("FileName", "DataGroup")
  vFilter = Array("ID=" & CStr(mlMasterReportID))
  cError = mobjDal.GetData("MasterReports", rsMasterReport, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsMasterReport.EOF = False Then
    cProcessViewName = "" & rsMasterReport.Fields("DataGroup").Value
   
  End If
  
  Set rsReportDirectory = New ADODB.Recordset
  rsReportDirectory.CursorLocation = adUseClient
  rsReportDirectory.LockType = adLockBatchOptimistic
  vFields = Array("ReportDirectory")
  vFilter = Array("ID=" & mlBuilderID)
  cError = mobjDal.GetData("Builder", rsReportDirectory, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  
  cReportDirectory = ""
  If rsReportDirectory.EOF = False Then
    cReportDirectory = "" & rsReportDirectory.Fields("ReportDirectory").Value
  End If
  
  If rsMasterReport.EOF = False Then
    cProcessViewName = "" & rsMasterReport.Fields("DataGroup").Value
    cProcessFileName = "" & rsMasterReport.Fields("FileName").Value
    If Right(Trim(cReportDirectory), 1) <> "\" Then
      cReportDirectory = cReportDirectory & "\"
    End If
    cProcessFileName = cReportDirectory & cProcessFileName
  Else
    cProcessViewName = ""
    cProcessFileName = ""
  End If
  Return

GetUserEmail:
  cUserEmail = ""
  Dim rsUser As ADODB.Recordset
  Set rsUser = New ADODB.Recordset
  rsUser.CursorLocation = adUseClient
  rsUser.LockType = adLockBatchOptimistic
  vFields = Array("Email")
  vFilter = Array("ID= " & mlUserID)
  cError = mobjDal.GetData("User", rsUser, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsUser.EOF = False Then
    If "" & rsUser("Email") <> "" Then
      cUserEmail = rsUser("Email")
    End If
  End If
  Return

GetBuilderWorkflowURL:
  cWorkflowURL = ""
  Dim rsBuilder As ADODB.Recordset
  Set rsBuilder = New ADODB.Recordset
  rsBuilder.CursorLocation = adUseClient
  rsBuilder.LockType = adLockBatchOptimistic
  vFields = Array("WorkflowURL")
  vFilter = Array("ID= " & mlBuilderID)
  cError = mobjDal.GetData("Builder", rsBuilder, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsBuilder.EOF = False Then
    If "" & rsBuilder("WorkflowURL") <> "" Then
      cWorkflowURL = rsBuilder("WorkflowURL")
    End If
  End If
  Return
  
GetMasterReportRecordset:
  Set rsReportRecordset = New ADODB.Recordset
  rsReportRecordset.CursorLocation = adUseClient
  rsReportRecordset.LockType = adLockBatchOptimistic
  vFields = Array("*")
  vFilter = Array("RecordID =" & mlClientID, "UserID=" & mlUserID)
  cError = mobjDal.GetData(cProcessViewName, rsReportRecordset, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  Return
  
SendResponseError:
  cError = Err.Description
  Resume SendResponseExit

End Function
Private Function PrefilZeros(Integer1 As Integer, Integer2 As Integer) As String
  Dim Int1 As String
  Dim Int2 As String

  Int1 = CStr(Integer1)
  Int2 = CStr(Integer2)

  If Len(Int1) = 1 Then Int1 = "0" & Int1
  If Len(Int2) = 1 Then Int2 = "0" & Int2
  PrefilZeros = Int1 & Int2

End Function
Private Sub Init()
  Dim cError As String
  
    mlBuilderID = 0
    mlDivisionID = 0
    mlTempDivisionID = 0
    mlSubdivID = 0
    mlUserID = 0
    mcINISection = 0
    mlClientID = 0
    mlIndividualID = 0
    mlBeBackID = 0
    lsLeadType = ""
    lsSource = ""
    lsStateName = ""
    lsMarketName = ""
    lsBuilderNumber = ""
    lsBuilderName = ""
    lsCommunityNumber = ""
    lsCommunityName = ""
    lsMasterCommunity = ""
    lsPlanName = ""
    lsPlanNumber = ""
    lsSpecNumber = ""
    lsSpecAddress = ""
    lsPrice = ""
    lsFirstName = ""
    lsTitle = ""
    lsLastName = ""
    lsEmail = ""
    lsPhone = ""
    lsStreetAddress = ""
    lsCity = ""
    lsState = ""
    lsPostalCode = ""
    lsCountry = ""
    lIndividualID = 0
    lsPrefPriceLow = ""
    lsPrefPricehigh = ""
    lsFinancing = ""
    lsReasonForBuying = ""
    lsMoveInDate = ""
    lsComments = ""
    lsCountElements = ""
    lsLeadDetails = ""
End Sub


Private Function ProcessData() As String
  Dim cErr As String
  Dim cError As String
  On Error GoTo ProcessDataError
  cErr = ReadINI
  If cErr <> "" Then Err.Raise 100, , cErr
  cErr = GetEleadsData
  cError = cError & cErr
    
    mrsDefaults.Filter = "Category = 'LeadTypeFilter'"
             
     Do While mrsDefaults.EOF = False
        If mrsDefaults("DefaultValue") = lsLeadType Then GoTo ProcessDataExit
        mrsDefaults.MoveNext
     Loop
    
    mrsDefaults.Filter = ""
  
  'Find the Subdivision and the User
  cErr = Regulate
  cError = cError & cErr
  cErr = InsertClient
  cError = cError & cErr
    If mlClientID <> 0 Then
      cErr = ProcessDemos(False)
      cError = cError & cErr
      
      cErr = SendToSalesAgent()
      cError = cError & cErr
    End If
   
  
  
ProcessDataExit:
  ProcessData = cError
  Exit Function
ProcessDataError:
  mlErrorCode = 226 'General Error Durring Process Data
  cError = "ReadINI: ON Builder:" & lsBuilderName & " Client Builder:" & lsBuilderName & "Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & " " & cErr & vbCrLf
  Resume ProcessDataExit
End Function

Private Function ReadINI() As String
  
  Dim cValue As String
  Dim cVirtDir As String
  Dim BuilderID As Long
  Dim DivisionID As Long
  Dim cError As String
  Dim lReturnSize As Long
  Dim cReturnValue As String
  Dim cfname As String
  Dim lPos As Long
  Dim cPath As String
  Const lSize As Long = 512
  On Error GoTo ReadINIError
  
' Retrieve the directory settings for the XML transfers. These are being stored
  ' in an INI file.
  ' Set the return buffer size...
  cReturnValue = String(lSize, 0)
  ' Set the .INI file name. This will use the current Windows directory and the name
  ' of the .INI file.
  mcVirtualPath = ""
  lReturnSize = GetWindowsDirectory(cReturnValue, lSize)
  cfname = Left$(cReturnValue, lReturnSize)
  If Right$(cfname, 1) <> "\" Then cfname = cfname & "\"
  cfname = cfname & "ELeads.ini"
  ' Ensure the .ini file exists.
  If Dir$(cfname) = "" Then
    MsgBox "Could not find '" & cfname & "'.", vbCritical + vbOKOnly, App.EXEName
  Else
    cVirtDir = ""
    cReturnValue = String$(lSize, 0)
    lReturnSize = GetPrivateProfileString(lsBuilderName, "VirtDir", "", cReturnValue, lSize, cfname)
    If lReturnSize > 0 Then
      mcVirtualPath = Left$(cReturnValue, lReturnSize)
    Else
      mlErrorCode = 227 'BuilderID Not Found Nothing  was done
      Err.Raise 100, , "Builder: " & lsBuilderName & " was not found"
    End If
    
    mlBuilderID = 0
    cReturnValue = String$(lSize, 0)
    lReturnSize = GetPrivateProfileString(lsBuilderName, "BuilderID", "", cReturnValue, lSize, cfname)
    If lReturnSize > 0 Then
      cValue = Left$(cReturnValue, lReturnSize)
      If IsNumeric(cValue) Then
        mlBuilderID = CLng(cValue)
      End If
    End If
    
    mlDivisionID = 0
    cReturnValue = String$(lSize, 0)
    lReturnSize = GetPrivateProfileString(lsBuilderName, "DivisionID", "", cReturnValue, lSize, cfname)
    If lReturnSize > 0 Then
      cValue = Left$(cReturnValue, lReturnSize)
      If IsNumeric(cValue) Then
        mlDivisionID = CLng(cValue)
      End If
    End If
  End If
  Set mobjDal = New SSnet_DAL.Functions
  mobjDal.BuilderID = mlBuilderID
  mobjDal.UserID = -10
  cError = mobjDal.SetINIFile(mcVirtualPath)

ReadINIExit:
  On Error Resume Next
  ReadINI = cError
  Exit Function
  
ReadINIError:
   mlErrorCode = 228 'General Read INI Error Nothing was Done
  cError = "ReadINI: ON Builder:" & lsBuilderName & " Client Builder:" & lsBuilderName & "Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & " :  BuilderName is not Recognized... " & vbCrLf
  GoTo ReadINIExit

End Function

Private Function Regulate() As String
  Dim rsDivision As ADODB.Recordset
  Dim vFields As Variant
  Dim cError As String
  Dim vFilter As Variant
  'Look for the Community in the tbElead Database under TransCom  Category.
  'If the Community Does not exist in tbElead then Look for a Division in
  'SS with the Division Number = Master Community if that one is not found
  'Then get the DivisionID from the INI File and get its Default Subdivision
  'If that is now found then Get the Default Division and get that division's
  'Default Subdivision

  On Error GoTo RegulateError
  
  mrsDefaults.Filter = ""
  mrsDefaults.Filter = "Category = 'TransComm' and  DefaultValue = '" & mobjDal.CorrectSQLText(lsCommunityName) & "'"
  
  If mrsDefaults.RecordCount > 0 Then
    'try to get Subdiv from XML
    mlSubdivID = mrsDefaults("ControllingID")
    'mlUserID = GetUserID(mlSubdivID) 'Removed because if its a be back don't rotate
    
    
'
'    mrsDefaults.Filter = ""
'    mrsDefaults.Filter = "Category = 'User' and  ControllingID = " & mlSubdivID
'    If mrsDefaults.RecordCount > 0 Then
'    mlUserID = mrsDefaults("DefaultID")
'    Else
'
'    If mlUserID = 0 Then
'      mlErrorCode = 229 'General Error No default User selected for Subdivision
'      Err.Raise 100, , "No default User selected for Subdivision"
'    End If
  Else
    'Try to get DivisionNum = Master Community  from XML
    Set rsDivision = New ADODB.Recordset
    rsDivision.CursorLocation = adUseClient
    rsDivision.LockType = adLockBatchOptimistic
    vFields = Array("ID")
    vFilter = Array("DivisionNum='" & mobjDal.CorrectSQLText(lsMasterCommunity) & "'")
    cError = mobjDal.GetData("Division", rsDivision, vFields, vFilter)
    If cError <> "" Then
      mlErrorCode = 230 'Error Occured While Retriving Division INformation in the Regulate Function
      Err.Raise 100, , cError
    End If
    If rsDivision.EOF = False Then
      mrsDefaults.Filter = ""
      mrsDefaults.Filter = "Category = 'Subdivision' and  ControllingID = " & rsDivision("ID")
      If mrsDefaults.RecordCount > 0 Then
        mlSubdivID = mrsDefaults("DefaultID")
        
        'mlMasterReportID = mrsDefaults("MasterReportID")
        'mcAttachments = mrsDefaults("Attachments")
        'mbSendEmail = mrsDefaults("SendEmail")
        'mcRespCCAddr = mrsDefaults("CC")
        'mcRespBCCAddr = mrsDefaults("BCC")
        'mcRespSubj = mrsDefaults("Subject")
        'RBR Added Rotator Code 6/22/2011
'        mlUserID = GetUserID(mlSubdivID)
'        If mlUserID = 0 Then
'          mlErrorCode = 229 'General Error No default User selected for Subdivision
'          Err.Raise 100, , "No default User selected for Subdivision"
'        End If
'        mrsDefaults.Filter = ""
'        mrsDefaults.Filter = "Category = 'User' and  ControllingID = " & mlSubdivID
'        If mrsDefaults.RecordCount > 0 Then
'          mlUserID = mrsDefaults("DefaultID")
'        Else
'          mlErrorCode = 229 'General Error No default User selected for Subdivision
'          Err.Raise 100, , "No default User selected for Subdivision"
'        End If
      Else
        mlErrorCode = 231 'General Error No default Subdivision selected for Division.
        Err.Raise 100, , "No default Subdivision selected for Division."
      End If
    Else
      If mlDivisionID <> 0 Then
      
        mrsDefaults.Filter = ""
        mrsDefaults.Filter = "Category = 'Subdivision' and  ControllingID = " & mlDivisionID
        If mrsDefaults.RecordCount > 0 Then
          mlSubdivID = mrsDefaults("DefaultID")
          'mlMasterReportID = mrsDefaults("MasterReportID")
          'mcAttachments = mrsDefaults("Attachments")
          'mbSendEmail = mrsDefaults("SendEmail")
          'mcRespCCAddr = mrsDefaults("CC")
          'mcRespBCCAddr = mrsDefaults("BCC")
          'mcRespSubj = mrsDefaults("Subject")
          'RBR Added Rotator Code 6/22/2011
'          mlUserID = GetUserID(mlSubdivID)
'          If mlUserID = 0 Then
'            mlErrorCode = 229 'General Error No default User selected for Subdivision
'            Err.Raise 100, , "No default User selected for Subdivision"
'          End If
'          mrsDefaults.Filter = ""
'          mrsDefaults.Filter = "Category = 'User' and  ControllingID = " & mlSubdivID
'          If mrsDefaults.RecordCount > 0 Then
'            mlUserID = mrsDefaults("DefaultID")
'          Else
'            mlErrorCode = 229 'General Error No default User selected for Subdivision
'            Err.Raise 100, , "No default User selected for Subdivision"
'          End If
        Else
          mlErrorCode = 231 'General Error No default Subdivision selected for Division.
          Err.Raise 100, , "No default Subdivision selected for Division."
        End If
      Else
    
        mrsDefaults.Filter = ""
        mrsDefaults.Filter = "Category = 'Division' and  ControllingID = " & mlBuilderID
        If mrsDefaults.RecordCount > 0 Then
          mlDivisionID = mrsDefaults("DefaultID")
          mrsDefaults.Filter = ""
          mrsDefaults.Filter = "Category = 'Subdivision' and  ControllingID = " & mlDivisionID
          If mrsDefaults.RecordCount > 0 Then
            mlSubdivID = mrsDefaults("DefaultID")
            mcDefDivRank = mrsDefaults("DefaultValue")
            
           ' mlMasterReportID = mrsDefaults("MasterReportID")
           ' mcAttachments = mrsDefaults("Attachments")
           ' mbSendEmail = mrsDefaults("SendEmail")
           ' mcRespCCAddr = mrsDefaults("CC")
           ' mcRespBCCAddr = mrsDefaults("BCC")
           ' mcRespSubj = mrsDefaults("Subject")
          'RBR Added Rotator Code 6/22/2011
'          mlUserID = GetUserID(mlSubdivID)
'          If mlUserID = 0 Then
'            mlErrorCode = 229 'General Error No default User selected for Subdivision
'            Err.Raise 100, , "No default User selected for Subdivision"
'          End If
'            mrsDefaults.Filter = ""
'            mrsDefaults.Filter = "Category = 'User' and  ControllingID = " & mlSubdivID
'            If mrsDefaults.RecordCount > 0 Then
'              mlUserID = mrsDefaults("DefaultID")
'            Else
'              mlErrorCode = 229 'General Error No default User selected for Subdivision
'             Err.Raise 100, , "No default User selected for Subdivision"
'            End If
          Else
            mlErrorCode = 231 'General Error No default Subdivision selected for Division.
            Err.Raise 100, , "No default Subdivision selected for Division."
          End If
        Else
          mlErrorCode = 232 'General Error No default Division selected for Builder.
          Err.Raise 100, , "No default Division selected for Builder."
        End If
      End If
    End If
    
  End If
  If mlSubdivID <> 0 Then
      Set rsDivision = New ADODB.Recordset
      rsDivision.CursorLocation = adUseClient
      rsDivision.LockType = adLockBatchOptimistic
      vFields = Array("DivisionID")
      vFilter = Array("ID=" & mlSubdivID)
      cError = mobjDal.GetData("Subdivision", rsDivision, vFields, vFilter)
      If rsDivision.EOF = False Then
        mlDivisionID = rsDivision("DivisionID").Value
         
          mrsDefaults.Filter = "Category = 'Subdivision' and  ControllingID = " & mlDivisionID
          If mrsDefaults.RecordCount > 0 Then
            mcDefDivRank = mrsDefaults("DefaultValue")
          End If
          
      End If
  End If


RegulateExit:
  On Error Resume Next
  Regulate = cError
  Set rsDivision = Nothing
  Exit Function

RegulateError:

  cError = "Regulate:  ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  GoTo RegulateExit

End Function

Private Function ProcessDemos(CheckNotes As Boolean) As String
  Dim x As Integer
  Dim Name As String
  Dim Value As String
  Dim dateCompare As Boolean
  Dim NumCompare As Boolean
  Dim tempDate As Date
  Dim tempNum As Long
  Dim tempNum2 As Long
  Dim vFields As Variant
  Dim vFilter As Variant
  Dim rsDatacodes As ADODB.Recordset
  Dim rsClient As ADODB.Recordset
  Dim rsNotes As ADODB.Recordset
  Dim cError As String
  Dim lRecID As Long
  On Error GoTo ProcessDemosError


  If mlBeBackID = 0 And mbIsUpdate = False Then
    cError = ProcessDefaultDemos
    If cError <> "" Then Err.Raise 100, , cError
  End If
  
  
  Set rsClient = New ADODB.Recordset
  rsClient.CursorLocation = adUseClient
  rsClient.LockType = adLockBatchOptimistic
  vFields = Array("*")
  vFilter = Array("ID = " & mlClientID)
  cError = mobjDal.GetData("Client", rsClient, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsClient.EOF = False Then
    Dim BeBackDemo As Boolean
    
    If Not QualNode Is Nothing Then
    For x = 0 To QualNode.Attributes.length - 1
      BeBackDemo = False
      
      Name = QualNode.Attributes.Item(x).nodeName
      Value = QualNode.Attributes.Item(x).nodeValue
      mrsDefaults.Filter = ""
      mrsDefaults.Filter = "Category = '" & Name & "'"
      If mrsDefaults.RecordCount > 0 Then
        If mrsDefaults.RecordCount > 1 Then
          If IsNumeric(Value) = True Then
            tempNum = CLng(Value)
            If mrsDefaults("NumCompare") = True And Not CheckNotes Then
              mrsDefaults.Filter = ""
              mrsDefaults.Filter = "Category = '" & Name & "' and MinRangeValue <=" & tempNum & " and MaxRangeValue >= " & tempNum
              If mrsDefaults.RecordCount > 0 Then
                Set rsDatacodes = New ADODB.Recordset
                rsDatacodes.CursorLocation = adUseClient
                rsDatacodes.LockType = adLockBatchOptimistic
                vFields = Array("SubdivisionDatacodeID", "BeBackUse")
                vFilter = Array("BuilderDatacodeID=" & CStr(mrsDefaults("DefaultID")), "SubdivisionID = " & mlSubdivID)
                cError = mobjDal.GetData("vuEleadsGetDatacodeCascadesBuilder", rsDatacodes, vFields, vFilter)
                If cError <> "" Then Err.Raise 100, , cError
                If rsDatacodes.EOF = False Then
                    If IsNull(rsDatacodes("BeBackUse")) = False Then
                        If rsDatacodes("BeBackUse") = 1 Then
                            BeBackDemo = True
                        End If
                    End If
                
                
                    If mlBeBackID = 0 Then
                      If mbIsUpdate = False Or (mbDontUpdateBeBackDC = False Or BeBackDemo = False) Then
                        If rsDatacodes.EOF = False Then
                          rsClient(mrsDefaults("DefaultValue") & "ID") = rsDatacodes("SubdivisionDatacodeID")
                        End If
                      End If
                      
                    Else
                      If IsNull(rsDatacodes("BeBackUse")) = False Then
                        Call AddBeBackDemo(rsDatacodes("BeBackUse"), rsDatacodes("SubdivisionDatacodeID"))
                      End If
                    End If
                End If
              End If
            End If
          ElseIf mrsDefaults("Note") = True Then
            If mrsDefaults("Category") <> "XMLInNote" And CheckNotes Then
              Set rsNotes = New ADODB.Recordset
              rsNotes.CursorLocation = adUseClient
              rsNotes.LockType = adLockBatchOptimistic
              vFields = Array("ID", "ClientID", "DateCreated", "CreatedByID", "Subject", "Note")
              vFilter = Array("ClientID = " & mlClientID, "Subject='" & mobjDal.CorrectSQLText(mrsDefaults("DefaultValue")) & "'", "Note like '" & mobjDal.CorrectSQLText(Value) & "'")
              cError = mobjDal.GetData("ClientNotes", rsNotes, vFields, vFilter)
              If cError <> "" Then Err.Raise 100, , cError
              If rsNotes.EOF = True Then
                rsNotes.AddNew
                rsNotes("ClientID").Value = mlClientID
                rsNotes("DateCreated").Value = ConvertDateTimeToDate(Now)
                rsNotes("CreatedByID").Value = mlUserID
                'rsNotes("Subject").Value = mrsDefaults("DefaultValue")
                cError = SetFieldValue(rsNotes("Subject"), CStr(mrsDefaults("DefaultValue").Value))
                If cError <> "" Then Err.Raise 100, , cError
                rsNotes("Note").Value = "" & Value
                rsNotes.Update
              ' Save this.
                cError = mobjDal.CreateData("ClientNotes", rsNotes, rsNotes.Bookmark, lRecID)
                If cError <> "" Then Err.Raise 100, , cError
              End If
            End If
          ElseIf mrsDefaults("DateCompare") = True And Not CheckNotes Then
            If IsDate(Value) = True Then
              tempDate = CDate(Value)
              tempDate = tempDate - Now
              tempNum = CLng(tempDate)
              mrsDefaults.Filter = ""
              mrsDefaults.Filter = "Category = '" & Name & "' and MinRangeValue <=" & tempNum & " and MaxRangeValue >= " & tempNum
              If mrsDefaults.RecordCount > 0 Then
                Set rsDatacodes = New ADODB.Recordset
                rsDatacodes.CursorLocation = adUseClient
                rsDatacodes.LockType = adLockBatchOptimistic
                vFields = Array("SubdivisionDatacodeID", "BeBackUse")
                vFilter = Array("BuilderDatacodeID=" & CStr(mrsDefaults("DefaultID")), "SubdivisionID = " & mlSubdivID)
                cError = mobjDal.GetData("vuEleadsGetDatacodeCascadesBuilder", rsDatacodes, vFields, vFilter)
                If cError <> "" Then Err.Raise 100, , cError
                If rsDatacodes.EOF = False Then
                    If IsNull(rsDatacodes("BeBackUse")) = False Then
                         If rsDatacodes("BeBackUse") = 1 Then
                             BeBackDemo = True
                         End If
                     End If
                     
                
                    If mlBeBackID = 0 Then
                      If mbIsUpdate = False Or (mbDontUpdateBeBackDC = False Or BeBackDemo = False) Then
                        If rsDatacodes.EOF = False Then
                          rsClient(mrsDefaults("DefaultValue") & "ID") = rsDatacodes("SubdivisionDatacodeID")
                        End If
                      End If
                      
                    Else
                      If IsNull(rsDatacodes("BeBackUse")) = False Then
                        Call AddBeBackDemo(rsDatacodes("BeBackUse"), rsDatacodes("SubdivisionDatacodeID"))
                      End If
                    End If
                End If
              End If
            End If
          ElseIf Right(mrsDefaults("DefaultValue"), 3) = "Txt" Then
            rsClient(mrsDefaults("DefaultValue").Value) = Value
          
          End If
           ElseIf Right(mrsDefaults("DefaultValue"), 3) = "Txt" Then
            rsClient(mrsDefaults("DefaultValue").Value) = Value
        Else
          If mrsDefaults("Note") = True Then
            If mrsDefaults("Category") <> "XMLInNote" And CheckNotes Then
              Set rsNotes = New ADODB.Recordset
              rsNotes.CursorLocation = adUseClient
              rsNotes.LockType = adLockBatchOptimistic
              vFields = Array("ID", "ClientID", "DateCreated", "CreatedByID", "Subject", "Note")
              vFilter = Array("ClientID = " & mlClientID, "Subject='" & mobjDal.CorrectSQLText(mrsDefaults("DefaultValue")) & "'", "Note like '" & mobjDal.CorrectSQLText(Value) & "'")
              cError = mobjDal.GetData("ClientNotes", rsNotes, vFields, vFilter)
              If cError <> "" Then Err.Raise 100, , cError
              If rsNotes.EOF = True Then
                rsNotes.AddNew
                rsNotes("ClientID").Value = mlClientID
                rsNotes("DateCreated").Value = ConvertDateTimeToDate(Now)
                rsNotes("CreatedByID").Value = mlUserID
                'rsNotes("Subject").Value = mrsDefaults("DefaultValue")
                cError = SetFieldValue(rsNotes("Subject"), mrsDefaults("DefaultValue"))
                If cError <> "" Then Err.Raise 100, , cError
                rsNotes("Note").Value = "" & Value
                rsNotes.Update
              ' Save this.
                cError = mobjDal.CreateData("ClientNotes", rsNotes, rsNotes.Bookmark, lRecID)
                If cError <> "" Then Err.Raise 100, , cError
              End If
            End If
          Else
            If Not CheckNotes Then
                Set rsDatacodes = New ADODB.Recordset
                rsDatacodes.CursorLocation = adUseClient
                rsDatacodes.LockType = adLockBatchOptimistic
                vFields = Array("SubdivisionDatacodeID", "BeBackUse")
                vFilter = Array("SystemName='" & mobjDal.CorrectSQLText(CStr(mrsDefaults("DefaultValue"))) & "'", "SubdivisionID = " & mlSubdivID, "Description = '" & mobjDal.CorrectSQLText(Value) & "'")
                cError = mobjDal.GetData("vuEleadsGetDatacodeCascadesBuilder", rsDatacodes, vFields, vFilter)
                If cError <> "" Then Err.Raise 100, , cError
                    If rsDatacodes.EOF = False Then
                        If IsNull(rsDatacodes("BeBackUse")) = False Then
                                 BeBackDemo = True
                         End If
                    
                        
                        If mlBeBackID = 0 Then
                          If mbIsUpdate = False Or (mbDontUpdateBeBackDC = False Or BeBackDemo = False) Then
                            If rsDatacodes.EOF = False Then
                              rsClient(mrsDefaults("DefaultValue") & "ID") = rsDatacodes("SubdivisionDatacodeID")
                            End If
                          End If
                          
                        Else
                          If IsNull(rsDatacodes("BeBackUse")) = False Then
                            Call AddBeBackDemo(rsDatacodes("BeBackUse"), rsDatacodes("SubdivisionDatacodeID"))
                          End If
                        End If
                    End If
                End If
            End If
        End If
      End If
      rsClient.Update
    Next x
    End If
  End If
  
  cError = mobjDal.UpdateData("Client", rsClient, rsClient.Bookmark)
  
  mrsDefaults.Filter = ""
  mrsDefaults.Filter = "Category = 'XMLInNote'"
  If mrsDefaults.RecordCount > 0 Then
    If mrsDefaults("Note") = True And gHasNotRunXMLInNote Then
        gHasNotRunXMLInNote = False
    
      Set rsNotes = New ADODB.Recordset
      rsNotes.CursorLocation = adUseClient
      rsNotes.LockType = adLockBatchOptimistic
      vFields = Array("ID", "ClientID", "DateCreated", "CreatedByID", "Subject", "Note")
      vFilter = Array("ID=0")
      cError = mobjDal.GetData("ClientNotes", rsNotes, vFields, vFilter)
      If cError <> "" Then Err.Raise 100, , cError
      
      rsNotes.AddNew
      rsNotes("ClientID").Value = mlClientID
      rsNotes("DateCreated").Value = ConvertDateTimeToDate(Now)
      rsNotes("CreatedByID").Value = mlUserID
      'rsNotes("Subject").Value = mrsDefaults("DefaultValue")
      cError = SetFieldValue(rsNotes("Subject"), mrsDefaults("DefaultValue"))
      If cError <> "" Then Err.Raise 100, , cError
      rsNotes("Note").Value = "" & mcLeadText
      rsNotes.Update
      ' Save this.
      cError = mobjDal.CreateData("ClientNotes", rsNotes, rsNotes.Bookmark, lRecID)
      If cError <> "" Then Err.Raise 100, , cError
    End If
  End If
  Call ProcessCommunityDemo
  
ProcessDemosExit:
  On Error Resume Next
  ProcessDemos = cError
  Set rsDatacodes = Nothing
  Set rsClient = Nothing
  Set rsNotes = Nothing
  Exit Function

ProcessDemosError:
mlErrorCode = 233 'General Error Processing Demographic Data
cError = "ProcessDemos:  ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
GoTo ProcessDemosExit

End Function

Private Function DumpLeadFile(Lead As MSXML2.IXMLDOMNode) As String
  Dim cError As String
  Dim LeadText As String
  Dim XMLChild As MSXML2.IXMLDOMNode
  Dim XMLAttrib As MSXML2.IXMLDOMAttribute
  
  For Each XMLAttrib In Lead.Attributes
    GoSub GetAttribValues
  Next
  
  For Each XMLChild In Lead.childNodes
    For Each XMLAttrib In XMLChild.Attributes
      GoSub GetAttribValues
    Next
  Next
  
DumpLeadFileExit:
  On Error Resume Next
  DumpLeadFile = LeadText
  Set XMLChild = Nothing
  Set XMLAttrib = Nothing
  Exit Function
  
DumpLeadFileError:
  mlErrorCode = 234 'General Error occured while generating a tab delimited XML string to put into a Note.
  cError = "DumpLeadFile:  " & Err.Description & vbCrLf
  GoTo DumpLeadFileExit
  
GetAttribValues:
   LeadText = LeadText & vbTab & XMLAttrib.nodeValue
  Return


End Function

Private Function AddBeBackDemo(BeBackUse As Integer, SubdivDatacodeID As Long) As String
  Dim rsBeBack As ADODB.Recordset
  Dim cError As String
  Dim vFields As Variant
  Dim vFilter As Variant
  
  On Error GoTo AddBeBackDemoError
  
  Set rsBeBack = New ADODB.Recordset
  rsBeBack.CursorLocation = adUseClient
  rsBeBack.LockType = adLockBatchOptimistic
  vFields = Array("ID", "ClientID", "VisitDate", "BeBackDemo1ID", "BeBackDemo2ID", "BeBackDemo3ID")
  vFilter = Array("ID=" & CStr(mlBeBackID))
  cError = mobjDal.GetData("tbBeBack", rsBeBack, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError

  rsBeBack("BeBackDemo" & CStr(BeBackUse) & "ID").Value = SubdivDatacodeID
  rsBeBack.Update
  ' Save this.
  cError = mobjDal.UpdateData("tbBeBack", rsBeBack, rsBeBack.Bookmark)
  If cError <> "" Then Err.Raise 100, , cError
  
AddBeBackDemoExit:
  On Error Resume Next
  Set rsBeBack = Nothing
  Exit Function
  
AddBeBackDemoError:
  cError = "AddBeBackDemo:  " & Err.Description & vbCrLf
  GoTo AddBeBackDemoExit

End Function


Private Function ProcessDefaultDemos() As String
  Dim cError As String
  Dim rsDatacodes As ADODB.Recordset
  Dim rsClient As ADODB.Recordset
  Dim vFields As Variant
  Dim vFilter As Variant
On Error GoTo ProcessDefaultDemosErr
  Set rsClient = New ADODB.Recordset
  rsClient.CursorLocation = adUseClient
  rsClient.LockType = adLockBatchOptimistic
  vFields = Array("ID", "Demo1ID", "Demo2ID", "Demo3ID", "Demo4ID", "Demo5ID", "Demo6ID", "Demo7ID", "Demo8ID", _
          "Demo9ID", "Demo10ID", "Demo11ID", "Demo12ID", "Demo13ID", "Demo14ID", "Demo15ID", "Demo16ID", "Demo17ID", "Demo18ID", "Demo19ID", "Demo20ID")
  vFilter = Array("ID = " & mlClientID)
  cError = mobjDal.GetData("Client", rsClient, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsClient.EOF = False Then

    mrsDefaults.Filter = ""
    mrsDefaults.Filter = "Category = 'DefaultDemo'"
    If mrsDefaults.RecordCount > 0 Then
      Do While mrsDefaults.EOF = False
        Set rsDatacodes = New ADODB.Recordset
        rsDatacodes.CursorLocation = adUseClient
        rsDatacodes.LockType = adLockBatchOptimistic
        vFields = Array("SubdivisionDatacodeID")
        vFilter = Array("BuilderDatacodeID=" & CStr(mrsDefaults("DefaultID")), "SubdivisionID = " & mlSubdivID)
        cError = mobjDal.GetData("vuEleadsGetDatacodeCascadesBuilder", rsDatacodes, vFields, vFilter)
        If cError <> "" Then Err.Raise 100, , cError
        If rsDatacodes.EOF = False Then
          rsClient(mrsDefaults("DefaultValue") & "ID") = rsDatacodes("SubdivisionDatacodeID")
          rsClient.Update
        End If
        mrsDefaults.MoveNext
      Loop
    End If
  End If
  cError = mobjDal.UpdateData("Client", rsClient, rsClient.Bookmark)
  If cError <> "" Then Err.Raise 100, , cError
ProcessDefaultDemosExit:
  On Error Resume Next
  ProcessDefaultDemos = cError
  Set rsDatacodes = Nothing
  Set rsClient = Nothing
Exit Function

ProcessDefaultDemosErr:
  mlErrorCode = 235 'General Error occured while populating Default Demographics.
  cError = "ProcessDefaultDemos:  ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
  Resume ProcessDefaultDemosExit

End Function

Private Function ProcessCommunityDemo() As String
On Error GoTo ProcessCommunityDemoErr
  Dim cError As String
  Dim rsDatacodes As ADODB.Recordset
  Dim rsClient As ADODB.Recordset
  Dim vFields As Variant
  Dim vFilter As Variant

  Set rsClient = New ADODB.Recordset
  rsClient.CursorLocation = adUseClient
  rsClient.LockType = adLockBatchOptimistic
  vFields = Array("ID", "Demo1ID", "Demo2ID", "Demo3ID", "Demo4ID", "Demo5ID", "Demo6ID", "Demo7ID", "Demo8ID", _
          "Demo9ID", "Demo10ID", "Demo11ID", "Demo12ID", "Demo13ID", "Demo14ID", "Demo15ID", "Demo16ID", "Demo17ID", "Demo18ID", "Demo19ID", "Demo20ID", "Demo21ID", "Demo22ID", "Demo23ID", "Demo24ID", "Demo25ID", "Demo26ID", "Demo27ID", "Demo28ID", "Demo29ID", "Demo30ID", "Demo31ID", "Demo32ID", "Demo33ID", "Demo34ID", "Demo35ID", "Demo36ID", "Demo37ID", "Demo38ID")
  vFilter = Array("ID = " & mlClientID)
  cError = mobjDal.GetData("Client", rsClient, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsClient.EOF = False Then

    mrsDefaults.Filter = ""
    mrsDefaults.Filter = "Category = 'CommunityNameDemo'"
    If mrsDefaults.RecordCount > 0 Then
      Do While mrsDefaults.EOF = False
        Set rsDatacodes = New ADODB.Recordset
        rsDatacodes.CursorLocation = adUseClient
        rsDatacodes.LockType = adLockBatchOptimistic
        vFields = Array("SubdivisionDatacodeID", "BeBackUse", "SystemName")
        vFilter = Array("LabelID=" & CStr(mrsDefaults("DefaultID")), "Description='" & mobjDal.CorrectSQLText(CStr(lsCommunityName)) & "'", "SubdivisionID = " & mlSubdivID)
        cError = mobjDal.GetData("vuEleadsGetDatacodeCascadesBuilder", rsDatacodes, vFields, vFilter)
        If cError <> "" Then Err.Raise 100, , cError
        If mlBeBackID = 0 Then
          If rsDatacodes.EOF = False Then
            rsClient(rsDatacodes("SystemName") & "ID") = rsDatacodes("SubdivisionDatacodeID")
            rsClient.Update
          End If
        Else
          If IsNull(rsDatacodes("BeBackUse")) = False Then
            Call AddBeBackDemo(rsDatacodes("BeBackUse"), rsDatacodes("SubdivisionDatacodeID"))
          End If
        End If
        mrsDefaults.MoveNext
      Loop
    End If
  End If
  cError = mobjDal.UpdateData("Client", rsClient, rsClient.Bookmark)
  If cError <> "" Then Err.Raise 100, , cError
ProcessCommunityDemoExit:
  ProcessCommunityDemo = cError
  On Error Resume Next
  Set rsDatacodes = Nothing
  Set rsClient = Nothing
Exit Function

ProcessCommunityDemoErr:
 cError = "ProcessCommunityDemo:  ON Builder:" & lsBuilderName & " Client Firstname:" & lsFirstName & " LastName: " & lsLastName & " Email: " & lsEmail & "  " & Err.Description & vbCrLf
 Resume ProcessCommunityDemoExit
  
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  Set mobjKernel = Nothing
  Set mobjDal = Nothing
  Set mrsDefaults = Nothing
  Set QualNode = Nothing
  mlBuilderID = 0
End Sub
Public Function SendResponse(ByVal cToAddress As String) As String
  Dim rsMasterReport As ADODB.Recordset
  Dim rsReportDirectory As ADODB.Recordset
  Dim rsReportRecordset As ADODB.Recordset
  Dim rsMasterPlan As ADODB.Recordset
  Dim rsResponseInfo As ADODB.Recordset
  Dim lMasterPlanID As Long
  Dim lMasterReportID As Long
  Dim cProcessViewName As String
  Dim cProcessFileName As String
  Dim cReportDirectory As String
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cMailContent As String
  Dim cFromAddress As String
  Dim cError As String

  GoSub GetMasterPlanID
  GoSub GetResponseInfo
  
  If mbSendEmail = True Then
    GoSub GetMasterReportInfo
    GoSub GetMasterReportRecordset
    cMailContent = MergeEmailContent(rsReportRecordset, cProcessFileName)
    cFromAddress = GetFromAddress()
    If Trim(cFromAddress) = "" Then Err.Raise 100, , "No email address was found to serve as the ‘From’ address.  No email was sent."
    cError = SendEmail(cToAddress, cFromAddress, mcRespSubj, mcRespCCAddr, mcRespBCCAddr, cMailContent, 0, mcAttachments, cReportDirectory)
  End If

SendResponseExit:
  SendResponse = cError
  Exit Function
  Return
  
  
GetMasterPlanID:
  Set rsMasterPlan = New ADODB.Recordset
  rsMasterPlan.CursorLocation = adUseClient
  rsMasterPlan.LockType = adLockBatchOptimistic
  vFields = Array("ID")
  vFilter = Array("DivisionID = " & mlDivisionID, "PlanModelNum='" & mobjDal.CorrectSQLText(lsPlanNumber) & "'")
  cError = mobjDal.GetData("MasterPlan", rsMasterPlan, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsMasterPlan.EOF = False Then
    lMasterPlanID = "" & rsMasterPlan("ID")
  Else
    lMasterPlanID = 0
  End If
  Return
  
  
GetResponseInfo:


  Set rsResponseInfo = New ADODB.Recordset
  rsResponseInfo.CursorLocation = adUseClient
  rsResponseInfo.LockType = adLockBatchOptimistic
  vFields = Array("SendEmail", "Attachments", "CC", "BCC", "Subject", "MasterReportID")
  If lsStatusName = "Lead" Then
    vFilter = Array("ResponseType='L'", "SubdivisionID=" & mlSubdivID)
  Else
    vFilter = Array("ResponseType='P'", "SubdivisionID=" & mlSubdivID)
  End If
  
  cError = mobjDal.GetData("tbEleadAutoResponse", rsResponseInfo, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsResponseInfo.EOF = False Then
    mbSendEmail = "" & rsResponseInfo("SendEmail").Value
    mcAttachments = "" & rsResponseInfo("Attachments").Value
    mcRespCCAddr = "" & rsResponseInfo("CC").Value
    mcRespBCCAddr = "" & rsResponseInfo("BCC").Value
    mcRespSubj = "" & rsResponseInfo("Subject").Value
    mlMasterReportID = "" & rsResponseInfo("MasterReportID").Value
  Else
    mbSendEmail = False
  End If
  
  
  
  If mbSendEmail = True Then
    Set rsResponseInfo = New ADODB.Recordset
    rsResponseInfo.CursorLocation = adUseClient
    rsResponseInfo.LockType = adLockBatchOptimistic
    vFields = Array("Attachments")
    vFilter = Array("SubdivisionID=" & mlSubdivID, "MasterPlanID = " & lMasterPlanID)
    cError = mobjDal.GetData("tbEleadAutoResponsePlanAttachments", rsResponseInfo, vFields, vFilter)
    If cError <> "" Then Err.Raise 100, , cError
    If rsResponseInfo.EOF = False Then
      mcAttachments = mcAttachments & "|" & rsResponseInfo("Attachments").Value
    End If
  End If
  Return

GetMasterReportInfo:
  Set rsMasterReport = New ADODB.Recordset
  rsMasterReport.CursorLocation = adUseClient
  rsMasterReport.LockType = adLockBatchOptimistic
  
  vFields = Array("FileName", "DataGroup")
  vFilter = Array("ID=" & CStr(mlMasterReportID))
  cError = mobjDal.GetData("MasterReports", rsMasterReport, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsMasterReport.EOF = False Then
    cProcessViewName = "" & rsMasterReport.Fields("DataGroup").Value
   
  End If
  
  Set rsReportDirectory = New ADODB.Recordset
  rsReportDirectory.CursorLocation = adUseClient
  rsReportDirectory.LockType = adLockBatchOptimistic
  vFields = Array("ReportDirectory")
  vFilter = Array("ID=" & mlBuilderID)
  cError = mobjDal.GetData("Builder", rsReportDirectory, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  
  cReportDirectory = ""
  If rsReportDirectory.EOF = False Then
    cReportDirectory = "" & rsReportDirectory.Fields("ReportDirectory").Value
  End If
  
  If rsMasterReport.EOF = False Then
    cProcessViewName = "" & rsMasterReport.Fields("DataGroup").Value
    cProcessFileName = "" & rsMasterReport.Fields("FileName").Value
    If Right(Trim(cReportDirectory), 1) <> "\" Then
      cReportDirectory = cReportDirectory & "\"
    End If
    cProcessFileName = cReportDirectory & cProcessFileName
  Else
    cProcessViewName = ""
    cProcessFileName = ""
  End If
  Return

GetMasterReportRecordset:
  Set rsReportRecordset = New ADODB.Recordset
  rsReportRecordset.CursorLocation = adUseClient
  rsReportRecordset.LockType = adLockBatchOptimistic
  vFields = Array("*")
  vFilter = Array("ClientID =" & mlClientID, "UserID=" & mlUserID)
  
  cError = mobjDal.GetData(cProcessViewName, rsReportRecordset, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  Return
  
SendResponseError:
  cError = Err.Description
  Resume SendResponseExit

End Function
Public Function MergeEmailContent(rsMailMerge As ADODB.Recordset, _
                                  cfname As String, Optional cWorkflowLink As String) As String
                           
  Dim irsIndex As Integer
  Dim iEmailFile As Integer
  Dim cReportTemplate  As String
  Dim cEmailFileStr As String
  Dim NextLine As String

  On Error GoTo MergeEmailContentError
  
  'Send the email out
  
  iEmailFile = FreeFile
  Open cfname For Input As iEmailFile
  Do Until EOF(iEmailFile)
    Line Input #iEmailFile, NextLine
    cEmailFileStr = cEmailFileStr + NextLine + vbCrLf
  Loop
  Close iEmailFile
          
  cReportTemplate = cEmailFileStr
  irsIndex = 0
  
  If rsMailMerge.EOF = False Then
    For irsIndex = 0 To rsMailMerge.Fields.count - 1
      If IsNull(rsMailMerge.Fields(irsIndex).Value) = True Then
        cReportTemplate = Replace(cReportTemplate, "{" & rsMailMerge.Fields(irsIndex).Name & "}", "")
      Else
        cReportTemplate = Replace(cReportTemplate, "{" & rsMailMerge.Fields(irsIndex).Name & "}", rsMailMerge.Fields(irsIndex).Value)
      End If
    Next
  End If
  If IsMissing(cWorkflowLink) = True Then cWorkflowLink = ""
  cReportTemplate = Replace(cReportTemplate, "{WorkflowLink}", cWorkflowLink)
    
    
MergeEmailContentExit:
  On Error Resume Next
  MergeEmailContent = cReportTemplate
  Exit Function
  
MergeEmailContentError:
  cReportTemplate = ""
  Resume MergeEmailContentExit

End Function
Private Function SendEmail(ByVal cToAddress As String, _
                           ByVal cFromAddress As String, _
                           ByVal cSubject As String, _
                           ByVal cCC As String, _
                           ByVal cBCC As String, _
                           ByVal cContent As String, _
                           ByVal iMsgType As Integer, _
                           Optional cAttachmentFiles As String, Optional ReportDirectory As String) As String
                           
                           
  Dim cError As String
  Dim objMail As Object
  Dim iPos1 As Integer
  Dim iPos2 As Integer
  Dim cAttachmentFile As String
  
  
  On Error GoTo SendError
      
  Set objMail = CreateObject("CDO.Message")
  With objMail
    iPos1 = 1
    iPos2 = InStr(iPos1, cAttachmentFiles, "|")
    Do While iPos2 > 0
        cAttachmentFile = (Mid$(cAttachmentFiles, iPos1, iPos2 - iPos1))
        If cAttachmentFile <> "" And Right(cAttachmentFile, 1) <> "\" Then
          .AddAttachment (ReportDirectory & "\attachments\" & cAttachmentFile)
        End If
        iPos1 = iPos2 + 1
        iPos2 = InStr(iPos1, cAttachmentFiles, "|")
    Loop
    
    If cFromAddress <> "" Then
      .From = cFromAddress
    End If
    .To = cToAddress
    If cCC <> "" Then
      .Cc = cCC
    End If
    If cBCC <> "" Then
      .Bcc = cBCC
    End If
    If cSubject <> "" Then
      .Subject = cSubject
    End If
    If InStr(1, UCase(cContent), UCase("<HTML")) Then
      .HTMLBody = cContent
    Else
      .TextBody = cContent
    End If
    .send
  End With
  
SendExit:
  On Error Resume Next
  Set objMail = Nothing
  SendEmail = cError
  Exit Function

SendError:
  cError = Err.Description
  Resume SendExit
  
End Function
Private Function GetFromAddress() As String
'This function checks the UserID for an email address if it doesn't have one
'then it check the Subdiv record if it doesn't have one
'then it check the Division record if it doesn't have one
'then it check the Builder record

  Dim rsUser As ADODB.Recordset
  Dim rsSubdiv As ADODB.Recordset
  Dim rsDivision As ADODB.Recordset
  Dim rsBuilder As ADODB.Recordset
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim Email As String
  Dim BuilderID As Long
  Dim DivisionID As Long
  Dim cError As String
  
  GoSub GetUserEmail
  If Trim(Email) = "" Then
    GoSub GetSubdivEmail
    If Trim(Email) = "" Then
      GoSub GetDivisionEmail
      If Trim(Email) = "" Then
        GoSub GetBuilderEmail
      End If
    End If
  End If

GetFromAddressExit:
  GetFromAddress = Email
  Exit Function

GetUserEmail:
  Email = ""
  Set rsUser = New ADODB.Recordset
  rsUser.CursorLocation = adUseClient
  rsUser.LockType = adLockBatchOptimistic
  vFields = Array("Email")
  vFilter = Array("ID= " & mlUserID)
  cError = mobjDal.GetData("User", rsUser, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsUser.EOF = False Then
    If "" & rsUser("Email") <> "" Then
      Email = rsUser("Email")
    End If
  End If
  Return

GetSubdivEmail:
  Email = ""
  Set rsSubdiv = New ADODB.Recordset
  rsSubdiv.CursorLocation = adUseClient
  rsSubdiv.LockType = adLockBatchOptimistic
  vFields = Array("Email")
  vFilter = Array("ID= " & mlSubdivID)
  cError = mobjDal.GetData("Subdivision", rsSubdiv, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsSubdiv.EOF = False Then
    If "" & rsSubdiv("Email") <> "" Then
      Email = rsSubdiv("Email")
    End If
  End If
  Return
  
GetDivisionEmail:
  Email = ""
  Set rsDivision = New ADODB.Recordset
  rsDivision.CursorLocation = adUseClient
  rsDivision.LockType = adLockBatchOptimistic
  vFields = Array("Email")
  vFilter = Array("ID= " & mlDivisionID)
  cError = mobjDal.GetData("Division", rsDivision, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsDivision.EOF = False Then
    If "" & rsDivision("Email") <> "" Then
      Email = rsDivision("Email")
    End If
  End If
  Return

GetBuilderEmail:
  Email = ""
  Set rsBuilder = New ADODB.Recordset
  rsBuilder.CursorLocation = adUseClient
  rsBuilder.LockType = adLockBatchOptimistic
  vFields = Array("Email")
  vFilter = Array("ID= " & mlBuilderID)
  cError = mobjDal.GetData("Builder", rsBuilder, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsBuilder.EOF = False Then
    If "" & rsBuilder("Email") <> "" Then
      Email = rsBuilder("Email")
    End If
  End If
  Return
  
End Function
Private Function GetSpecLot(ByRef LotID As Long) As String
  Dim rsLot As ADODB.Recordset
  Dim lLotID As Long
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cError As String
  
  
  
  LotID = 0
  Set rsLot = New ADODB.Recordset
  rsLot.CursorLocation = adUseClient
  rsLot.LockType = adLockBatchOptimistic
  vFields = Array("ID", "PhaseID", "LotStatusID", "LotUnitNum", "BlockNum", "PhasePlanID", _
                "Address1", "LegalAddress1", "Premium", "JobUnitNum", "ClientID")
  vFilter = Array("LotUnitNum = '" & mobjDal.CorrectSQLText(lsSpecNumber) & "'", "Address1 = '" & mobjDal.CorrectSQLText(lsSpecAddress) & "'", "SubdivisionID = " & mlSubdivID, "StatusName = 'Spec'")
  cError = mobjDal.GetData("vuGetLotsBySubdivision", rsLot, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsLot.EOF = False Then
    LotID = rsLot("ID")
  End If
  
  
  
GetSpecLotExit:
GetSpecLot = cError
Exit Function

GetSpecLotErr:
  cError = Err.Description
  GoTo GetSpecLotExit
  
End Function
Private Function SetLot(lLotID As Long, ByVal lClientID As Long) As String
'If Spec, then check tbLot.PhasePlanID, GarageOrientationID, ColorSchemeID, HouseBase, TrimMasonry, RoofTile
'If any of these contain data (the PhasePlanID and the GarageOrientationID SHOULD) then
'Write the LotID to tbClient, and write any/all XXX.ID’s from tbLot to the corresponding fields in tbClient.
'Check tbBuiltOptions for records with the same LotID where MostRecent = True and Removed = False
'COPY any matches to tbSelectedOption for the ClientID.
'Check tbAdjustments for any records with the same LotID.  COPY any matches to tbAdjustments with the ClientID.
'If tbLot.PhasePlanID/ GarageOrientationID are null for some reason, then write the LotID to tbClient and send
'an error email.  Do not bother looking up the plan, if there is one in the XML file.

 


  Dim rsLot As ADODB.Recordset
  Dim rsBuiltOptions As ADODB.Recordset
  Dim rsSelectedOptions As ADODB.Recordset
  Dim rsAdjustments As ADODB.Recordset
  Dim rsClientAdjustments As ADODB.Recordset
  Dim rsLotAdjustments As ADODB.Recordset
  Dim rsClient As ADODB.Recordset
  Dim iAdjustmentType As Integer
  Dim vOrder As Variant
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cError As String
  Dim dTotalAdjustments As Double
  
  
  'There was, so see if the lot we're changing to is a Spec lot.
  Set rsLot = New ADODB.Recordset
  rsLot.CursorLocation = adUseClient
  rsLot.LockType = adLockBatchOptimistic
  vFields = Array("ID", "StatusName", "GarageOrientationID", "PhasePlanID", "ColorSchemeID", "HouseBase", "TrimMasonry", "RoofTile")
  vFilter = Array("ID=" & CStr(lLotID))
  cError = mobjDal.GetData("GetLotsBySubdivision", rsLot, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsLot("StatusName").Value = gcLOTSTATUS_SPEC Then
    Set rsClient = New ADODB.Recordset
    rsClient.CursorLocation = adUseClient
    rsClient.LockType = adLockBatchOptimistic
    vFields = Array("ID", "LotID", "GarageOrientationID", "PlanID", "ColorSchemeID", "HouseBase", "TrimMasonry", "RoofTile")
    vFilter = Array("ID=" & CStr(lClientID))
    cError = mobjDal.GetData("tbClient", rsClient, vFields, vFilter)
    If cError <> "" Then Err.Raise 100, , cError
    If rsClient.EOF = False Then
      rsClient("LotID").Value = rsLot("ID").Value
      rsClient("PlanID").Value = rsLot("PhasePlanID").Value
      rsClient("ColorSchemeID").Value = rsLot("ColorSchemeID").Value
      rsClient("HouseBase").Value = rsLot("HouseBase").Value
      rsClient("TrimMasonry").Value = rsLot("TrimMasonry").Value
      rsClient("RoofTile").Value = rsLot("RoofTile").Value
      rsClient("GarageOrientationID").Value = rsLot("GarageOrientationID").Value
      
      rsClient.Update
      cError = mobjDal.UpdateData("Client", rsClient, rsClient.Bookmark)
      If cError <> "" Then Err.Raise 100, , cError
    End If
      
    ' This is a Spec lot, so get the Built Options.
    Set rsBuiltOptions = New ADODB.Recordset
    rsBuiltOptions.CursorLocation = adUseClient
    rsBuiltOptions.LockType = adLockBatchOptimistic
    vFields = Array("ID", "AvailableOptionID", "OptionTypeID", "OptionGroupID", "Qty", _
                    "Price", "OptionDesc", "CustomerDesc", "Elevation", "BuilderApproved")
    ' 12/12/02 DSP Changed to use ONLY the current active built options.
    vFilter = Array("LotID=" & CStr(lLotID), "MostRecent=TRUE", "Removed=FALSE")
    cError = mobjDal.GetData("BuiltOptions", rsBuiltOptions, vFields, vFilter)
    If cError <> "" Then Err.Raise 100, , cError
    
    ' See if there are any Built Options.
    If rsBuiltOptions.RecordCount > 0 Then
      ' There are, so get the Selected Options.
      Set rsSelectedOptions = New ADODB.Recordset
      rsSelectedOptions.CursorLocation = adUseClient
      rsSelectedOptions.LockType = adLockBatchOptimistic
      vFields = Array("ID", "AvailableOptionID", "OptionTypeID", "OptionGroupID", "Qty", _
                      "Price", "OptionDesc", "CustomerDesc", "Elevation", "ClientID", "BuilderApproved", "CustomerApproved")
      vFilter = Array("ClientID=" & CStr(lClientID))
      cError = mobjDal.GetData("SelectedOption", rsSelectedOptions, vFields, vFilter)
      If cError <> "" Then Err.Raise 100, , cError
      
      With rsSelectedOptions
        ' If any records are returned, they will need to be deleted first.
        If .RecordCount > 0 Then
          vFilter = Array("ClientID=" & CStr(lClientID))
          cError = mobjDal.deleteData("SelectedOption", vFilter)
          If cError <> "" Then Err.Raise 100, , cError
          ' Clear all records out of the local set so that we can use
          ' the recordset for adding records.
          .MoveFirst
          Do
            .Delete
            .MoveNext
          Loop While .EOF = False
        End If
        
        ' Now copy all of the Built Options to the Selected Options.
        rsBuiltOptions.MoveFirst
        ' Cycle through all of the records.
        Do
          .AddNew
          .Fields("AvailableOptionID").Value = rsBuiltOptions("AvailableOptionID").Value
          .Fields("OptionTypeID").Value = rsBuiltOptions("OptionTypeID").Value
          .Fields("OptionGroupID").Value = rsBuiltOptions("OptionGroupID").Value
          .Fields("Qty").Value = rsBuiltOptions("Qty").Value
          .Fields("Price").Value = rsBuiltOptions("Price").Value
          .Fields("OptionDesc").Value = rsBuiltOptions("OptionDesc").Value
          .Fields("CustomerDesc").Value = rsBuiltOptions("CustomerDesc").Value
          .Fields("Elevation").Value = rsBuiltOptions("Elevation").Value
          .Fields("ClientID").Value = lClientID
          If IsNull(rsBuiltOptions("AvailableOptionID").Value) = True Then
            .Fields("BuilderApproved").Value = rsBuiltOptions("BuilderApproved").Value
            .Fields("CustomerApproved").Value = giOPTION_IN_PROCESS
          End If
          .Update
          rsBuiltOptions.MoveNext
        Loop While rsBuiltOptions.EOF = False
      End With
      
      ' Now save all of these new Selected Option records.
      cError = mobjDal.CreateDataBatch("SelectedOption", rsSelectedOptions)
      If cError <> "" Then Err.Raise 100, , cError
    End If
    
    
    
    Set rsLotAdjustments = New ADODB.Recordset
    rsLotAdjustments.CursorLocation = adUseClient
    rsLotAdjustments.LockType = adLockBatchOptimistic
    vFields = Array("ClientID", "LotID", "Description", "Adjustment", "AdjustmentDate", "Type", "UserID")
    ' 12/12/02 DSP Changed to use ONLY the current active built options.
    vFilter = Array("LotID=" & CStr(lLotID))
    cError = mobjDal.GetData("tbAdjustments", rsLotAdjustments, vFields, vFilter)
    If cError <> "" Then Err.Raise 100, , cError
    
    ' See if there are any Built Options.
    If rsLotAdjustments.RecordCount > 0 Then
      ' There are, so get the Selected Options.
      Set rsClientAdjustments = New ADODB.Recordset
      rsClientAdjustments.CursorLocation = adUseClient
      rsClientAdjustments.LockType = adLockBatchOptimistic
      vFields = Array("ClientID", "LotID", "Description", "Adjustment", "AdjustmentDate", "Type", "UserID")
      vFilter = Array("ClientID=" & CStr(lClientID))
      cError = mobjDal.GetData("tbAdjustments", rsClientAdjustments, vFields, vFilter)
      If cError <> "" Then Err.Raise 100, , cError
      
      With rsClientAdjustments
        ' If any records are returned, they will need to be deleted first.
        If .RecordCount > 0 Then
          vFilter = Array("ClientID=" & CStr(lClientID))
          cError = mobjDal.deleteData("tbAdjustments", vFilter)
          If cError <> "" Then Err.Raise 100, , cError
          ' Clear all records out of the local set so that we can use
          ' the recordset for adding records.
          .MoveFirst
          Do
            .Delete
            .MoveNext
          Loop While .EOF = False
        End If
        
        ' Now copy all of the Built Options to the Selected Options.
        rsLotAdjustments.MoveFirst
        ' Cycle through all of the records.
        Do
          .AddNew
          .Fields("Description").Value = rsLotAdjustments("Description").Value
          .Fields("Adjustment").Value = rsLotAdjustments("Adjustment").Value
          .Fields("AdjustmentDate").Value = rsLotAdjustments("AdjustmentDate").Value
          .Fields("Type").Value = rsLotAdjustments("Type").Value
          .Fields("ClientID").Value = lClientID
          .Fields("UserID").Value = rsLotAdjustments("UserID").Value
          .Update
          rsLotAdjustments.MoveNext
        Loop While rsLotAdjustments.EOF = False
      End With
      
      ' Now save all of these new Selected Option records.
      cError = mobjDal.CreateDataBatch("tbAdjustments", rsClientAdjustments)
      If cError <> "" Then Err.Raise 100, , cError
      
      'The following lines will re-calculate the client change values
      iAdjustmentType = giCUSTOMER_LOT_PREMIUM
        GoSub CalcAndSaveAdjustments
      iAdjustmentType = giCUSTOMER_LOT_INCENTIVE
        GoSub CalcAndSaveAdjustments
      iAdjustmentType = giCUSTOMER_OPTION_INCENTIVE
        GoSub CalcAndSaveAdjustments
      iAdjustmentType = giCUSTOMER_UPFRONT_INCENTIVE
        GoSub CalcAndSaveAdjustments
      iAdjustmentType = giCUSTOMER_PRICE_CHANGE
        GoSub CalcAndSaveAdjustments
      iAdjustmentType = giCUSTOMER_INCENTIVE_CHANGE
        GoSub CalcAndSaveAdjustments
    End If
  End If
    
  
  
GetSpecLotExit:
SetLot = cError
Exit Function


GetAdjustments:
  ' Retrieve adjustment information.
  Set rsAdjustments = New ADODB.Recordset
  rsAdjustments.CursorLocation = adUseClient
  rsAdjustments.LockType = adLockBatchOptimistic
  vFields = Array("ID", "ClientID", "LotID", "Description", _
            "Adjustment", "AdjustmentDate", "Type", "UserID", "FirstName", "LastName")
  ' Filter for the type of adjustment we're retrieving.
  Select Case iAdjustmentType
    Case giCUSTOMER_LOT_PREMIUM
      vFilter = Array("ClientID=" & CStr(lClientID), "Type='Premium Change'")
    Case giCUSTOMER_LOT_INCENTIVE
      vFilter = Array("ClientID=" & CStr(lClientID), "Type='Premium Incentive'")
    Case giCUSTOMER_OPTION_INCENTIVE
      vFilter = Array("ClientID=" & CStr(lClientID), "Type='Option Incentive'")
    Case giCUSTOMER_UPFRONT_INCENTIVE
      vFilter = Array("ClientID=" & CStr(lClientID), "Type='Upfront Incentive'")
    Case giCUSTOMER_PRICE_CHANGE
      vFilter = Array("ClientID=" & CStr(lClientID), "Type='Price Change'")
    Case giCUSTOMER_INCENTIVE_CHANGE
      vFilter = Array("ClientID=" & CStr(lClientID), "Type='Price Incentive'")
    Case giSPEC_PRICE_CHANGE
      vFilter = Array("LotID=" & CStr(lLotID), "Type='Price Change'")
    Case giSPEC_LOT_PREMIUM
      vFilter = Array("LotID=" & CStr(lLotID), "Type='Premium Change'")
    Case giSPEC_OPTION_INCENTIVE
      vFilter = Array("LotID=" & CStr(lLotID), "Type='Option Incentive'")
  End Select
  vOrder = Array("AdjustmentDate")
  cError = mobjDal.GetData("GetAdjustments", rsAdjustments, vFields, vFilter, vOrder)
  If cError <> "" Then Err.Raise 100, , cError
  Return
  
GetClientInfo:
  ' Get the adjustment fields from the Client record.
  Set rsClient = New ADODB.Recordset
  rsClient.CursorLocation = adUseClient
  rsClient.LockType = adLockBatchOptimistic
  vFields = Array("ID", "LotPremiumChange", "LotPremiumIncentive", "OptionIncentives", _
                  "OtherIncentives", "BasePrice", "BasePriceChange", "PriceIncentive")
  vFilter = Array("ID=" & CStr(lClientID))
  cError = mobjDal.GetData("Client", rsClient, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsClient.RecordCount <> 1 Then
    Err.Raise 100, , "Could not find Client information."
  End If
  Return
 
    
CalcAndSaveAdjustments:
  ' Add up all the adjustments for the total amount.
  GoSub GetAdjustments
  dTotalAdjustments = 0
  If rsAdjustments.RecordCount > 0 Then
    Do
      dTotalAdjustments = dTotalAdjustments + rsAdjustments("Adjustment").Value
      rsAdjustments.MoveNext
    Loop While rsAdjustments.EOF = False
    dTotalAdjustments = Round(dTotalAdjustments, 0)
  
    GoSub GetClientInfo
    
    ' Save the new adjusted total to the appropriate location.
    If rsClient.RecordCount > 0 Then
      Select Case iAdjustmentType
        Case giCUSTOMER_LOT_PREMIUM
          rsClient("LotPremiumChange").Value = dTotalAdjustments
        Case giCUSTOMER_LOT_INCENTIVE
          rsClient("LotPremiumIncentive").Value = dTotalAdjustments
        Case giCUSTOMER_OPTION_INCENTIVE
          rsClient("OptionIncentives").Value = dTotalAdjustments
        Case giCUSTOMER_UPFRONT_INCENTIVE
          rsClient("OtherIncentives").Value = dTotalAdjustments
        Case giCUSTOMER_PRICE_CHANGE
          ' Update the change amount.
          rsClient("BasePriceChange").Value = dTotalAdjustments
        Case giCUSTOMER_INCENTIVE_CHANGE
          rsClient("PriceIncentive").Value = dTotalAdjustments
      End Select
      rsClient.Update
      cError = mobjDal.UpdateData("Client", rsClient, rsClient.Bookmark)
      If cError <> "" Then Err.Raise 100, , cError
    End If
  End If
  Return

GetSpecLotErr:
  cError = Err.Description
  Resume GetSpecLotExit
  
End Function
Private Function SetUsed(ID As Long, Recieved As Long) As String


  Dim rsRotator As ADODB.Recordset
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cError As String
  
  Set rsRotator = New ADODB.Recordset
  rsRotator.CursorLocation = adUseClient
  rsRotator.LockType = adLockBatchOptimistic
  vFields = Array("*")
  vFilter = Array("ID=" & CStr(ID))
  cError = mobjDal.GetData("tbEleadRotator", rsRotator, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsRotator.EOF = False Then
    rsRotator.Fields("Recieved").Value = Recieved
    rsRotator.Update
    cError = mobjDal.UpdateData("tbEleadRotator", rsRotator, rsRotator.Bookmark)
  End If
      
End Function
Private Function GetUserID(SubdivisionID As Long) As Long
'RBR Added Rotator Code 6/22/2011
'rotator will get all active users for this subdivisoni that are included in the rotator.
'Find the next one in the list that hasn't recieved a Client if everyone has recieved one it will
'Clear all the recieved bits and start again.

  Dim rsRotator As ADODB.Recordset
  Dim vOrder As Variant
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cError As String
  Dim x As String
  
  mlUserID = 0
  
  If lsUserName <> "" Then
    Set rsRotator = New ADODB.Recordset
    rsRotator.CursorLocation = adUseClient
    rsRotator.LockType = adLockBatchOptimistic
    vFields = Array("ID")
    vFilter = Array("BuilderID=" & mlBuilderID, "UserName='" & lsUserName & "'", "Active = 1")
    
    cError = mobjDal.GetData("tbUser", rsRotator, vFields, vFilter)
    
    If cError <> "" Then Err.Raise 100, , cError
    If rsRotator.EOF = False Then
      mlUserID = rsRotator.Fields("ID").Value
    Else
      mlUserID = 0
    End If
    
    
        
  
  End If
  
  If mlUserID = 0 Then
    Set rsRotator = New ADODB.Recordset
    rsRotator.CursorLocation = adUseClient
    rsRotator.LockType = adLockBatchOptimistic
    
    vFields = Array("*")
    vFilter = Array("Recieved=0", "SubdivisionID=" & CStr(SubdivisionID), "Active = 1")
    vOrder = Array("ID asc")
    
    cError = mobjDal.GetData("vuGetEleadsRotatorForActiveUser", rsRotator, vFields, vFilter, vOrder)
    If cError <> "" Then Err.Raise 100, , cError
    
    If rsRotator.EOF = False Then
      mlUserID = rsRotator.Fields("UserID").Value
      x = SetUsed(rsRotator.Fields("ID"), 1)
    Else
      Set rsRotator = New ADODB.Recordset
      rsRotator.CursorLocation = adUseClient
      rsRotator.LockType = adLockBatchOptimistic
      vFields = Array("*")
      vFilter = Array("Recieved = 1", "SubdivisionID=" & CStr(SubdivisionID), "Active = 1")
      vOrder = Array("ID asc")
      cError = mobjDal.GetData("vuGetEleadsRotatorForActiveUser", rsRotator, vFields, vFilter, vOrder)
      
      If cError <> "" Then Err.Raise 100, , cError
      Do While rsRotator.EOF = False
          x = SetUsed(rsRotator.Fields("ID"), 0)
          rsRotator.MoveNext
      Loop
      
      Set rsRotator = New ADODB.Recordset
      rsRotator.CursorLocation = adUseClient
      rsRotator.LockType = adLockBatchOptimistic
      vFields = Array("*")
      vFilter = Array("Recieved =0", "SubdivisionID=" & CStr(SubdivisionID), "Active = 1")
      vOrder = Array("ID asc")
      cError = mobjDal.GetData("vuGetEleadsRotatorForActiveUser", rsRotator, vFields, vFilter, vOrder)
      
      If cError <> "" Then Err.Raise 100, , cError
      If rsRotator.EOF = False Then
        mlUserID = rsRotator.Fields("UserID").Value
        x = SetUsed(rsRotator.Fields("ID"), 1)
      End If
    End If
  End If

  
  
  
'
'  If rsRotator.EOF = False Then
'    rsRotator.Filter = "Recieved = 0"
'    If rsRotator.RecordCount > 0 Then 'found Our winner
'
'    Else
'      rsRotator.Filter = ""
'      Do While rsRotator.EOF = False
'        x = SetUsed(rsRotator.Fields("ID"), 0)
'        rsRotator.MoveNext
'      Loop
'
'     rsRotator.Filter = "Recieved=0"
'     If rsRotator.RecordCount > 0 Then 'found Our winner
'        rsRotator.MoveFirst
'        mlUserID = rsRotator.Fields("UserID").Value
'        x = SetUsed(rsRotator.Fields("ID"), 1)
'     End If
'    End If
'
'  End If
  GetUserID = mlUserID
  
End Function

Private Function GetScenario(IndividualID As Long) As String
  Dim rsFLI As ADODB.Recordset
  Dim rsScenario As ADODB.Recordset
  Dim vOrder As Variant
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cError As String
  Dim scen As String
  Dim count As String
  
  Set rsFLI = New ADODB.Recordset
  rsFLI.CursorLocation = adUseClient
  rsFLI.LockType = adLockBatchOptimistic
  vFields = Array("*")
  vFilter = Array("ScreenName='CustTPB' and SystemName='Scenario'", "SubdivisionID=" & CStr(mlSubdivID))
  cError = mobjDal.GetData("vuGetFieldLabelsBySubdivision", rsFLI, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsFLI.EOF = False Then
  
   scen = rsFLI.Fields("Display").Value
  
  End If
  
  
  Set rsScenario = New ADODB.Recordset
  rsScenario.CursorLocation = adUseClient
  rsScenario.LockType = adLockBatchOptimistic
  vFields = Array("COUNT(*) as cnt")
  vFilter = Array("IndividualID =" & CStr(IndividualID), "MemberTypeID = 1")
  cError = mobjDal.GetData("tbMember", rsScenario, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsScenario.EOF = False Then
  
    count = rsScenario.Fields("cnt").Value
    
  End If
  
  GetScenario = scen & " " & count + 1
  
End Function

Private Function GetCountryID(Country As String) As Integer

  Dim rsCountry As ADODB.Recordset
  Dim vOrder As Variant
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cError As String
  Dim scen As String
  Dim count As String
  
  Set rsCountry = New ADODB.Recordset
  rsCountry.CursorLocation = adUseClient
  rsCountry.LockType = adLockBatchOptimistic
  vFields = Array("*")
  vFilter = Array("Country='" & CStr(Country) & "'")
  cError = mobjDal.GetData("tbCountry", rsCountry, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsCountry.EOF = False Then
  
   GetCountryID = rsCountry.Fields("ID").Value
   
   Else
    GetCountryID = 0
  
  End If
  
  
  
End Function

Private Function GetStateID(State As String, lCountryID As Long) As Integer

  Dim rsState As ADODB.Recordset
  Dim vOrder As Variant
  Dim vFilter As Variant
  Dim vFields As Variant
  Dim cError As String
 
  Set rsState = New ADODB.Recordset
  rsState.CursorLocation = adUseClient
  rsState.LockType = adLockBatchOptimistic
  vFields = Array("*")
  If lCountryID > 0 Then
    vFilter = Array("Code='" & CStr(State) & "'", "CountryID = " & lCountryID)
  Else
    vFilter = Array("Code='" & CStr(State) & "'")
  End If
  
  
  cError = mobjDal.GetData("tbState", rsState, vFields, vFilter)
  If cError <> "" Then Err.Raise 100, , cError
  If rsState.EOF = False Then
  
   GetStateID = rsState.Fields("ID").Value
   
   Else
    GetStateID = 0
  End If
End Function

Private Function CheckIndividualUnique(lsFirstName As String, _
                                        lsLastName As String, _
                                        lsStreetAddress As String, _
                                        lsEmail As String, _
                                        ByRef lIndividualID As Long, _
                                        ByRef bMatchFound As Boolean)
                                        
    Dim cError As String
    Dim rsUniqueIndividual As ADODB.Recordset
    Dim vFilter As Variant
    Dim vFields As Variant

  
   
    
    If bValidateByEmailInsteadOfaddress = False Then
        cError = mobjKernel.CheckIndividualUnique(mobjDal, _
                                        lsFirstName, _
                                        lsLastName, _
                                        lsStreetAddress, _
                                        lIndividualID, _
                                        bMatchFound)
                                        
    Else
        Set rsUniqueIndividual = New ADODB.Recordset
        rsUniqueIndividual.CursorLocation = adUseClient
        rsUniqueIndividual.LockType = adLockBatchOptimistic
        
        vFields = Array("*")
        vFilter = Array("Firstname='" & lsFirstName & "'", "LastName='" & lsLastName & "'", "Email='" & lsEmail & "'")
        
        cError = mobjDal.GetData("tbIndividual", rsUniqueIndividual, vFields, vFilter)
        If cError <> "" Then Err.Raise 100, , cError
        
        If rsUniqueIndividual.EOF = False Then
            bMatchFound = True
            lIndividualID = rsUniqueIndividual("ID")
        End If
    End If


End Function


