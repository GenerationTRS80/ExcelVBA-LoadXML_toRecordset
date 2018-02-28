Public Function acOpenRecordset_from_ExcelWorksheet(xlWrkBk_SP As Excel.Workbook, Optional bClose_SP_Workbook As Boolean = True) As Boolean

'----------------------------------------------------------------------------------------------------------------------
'   This sub will take a Cost Model that is open on the ITO Deal Site and copy the CMI tab into a recordset object
'

'Objects
 Dim xDocCMI As Object
 Dim xDocWorksheet_Tab_DBUpload As Object
 'Dim msXMLDoc As MSXML2.DOMDocument 'this object is just for testing not used in production
 Dim rsCMI As ADODB.Recordset
 Dim rsWorksheet_Tab_DBUpload As ADODB.Recordset
 Dim rsFilter As ADODB.Recordset

'Variables
 Dim xlWrkSht_CMI As Excel.Worksheet
 Dim xlWrkSht_Tab_DBUpload As Excel.Worksheet
 Dim rngCMI As Excel.Range
 Dim rngWorksheet_Tab_DBUpload As Excel.Range
 Dim sFilterString As String
 
'Set CopyRecordset to Spreadsheet to TRUE
 acOpenRecordset_from_ExcelWorksheet = True
 
 On Error GoTo ProcErr
 
'Instantiate objects for CMI tab and DBUpload tab
 Set rsCMI = New ADODB.Recordset
 Set xDocCMI = CreateObject("MSXML2.DOMDocument")
 
 Set rsWorksheet_Tab_DBUpload = New ADODB.Recordset
 Set xDocWorksheet_Tab_DBUpload = CreateObject("MSXML2.DOMDocument")
 
'Get Corp Model Input worksheet and DBUpload Worksheet
 Set xlWrkSht_CMI = xlWrkBk_SP.Worksheets("Corp Model Input")
 Set xlWrkSht_Tab_DBUpload = xlWrkBk_SP.Worksheets("DBUpload")


'Get cells C1 through BH397 from CMI tab
'*** NOTE: using these references instead of a name range allow to select historical cost models without a name range created in them ***
 Set rngCMI = xlWrkSht_CMI.Range(xlWrkSht_CMI.Cells(PULL_WORKSHEET_START_ROWNUM, PULL_WORKSHEET_START_COLUMNNUM), xlWrkSht_CMI.Cells(PULL_WORKSHEET_ROWCOUNT, PULL_WORKSHEET_COLUMNCOUNT))
 Set rngWorksheet_Tab_DBUpload = xlWrkSht_Tab_DBUpload.Range(xlWrkSht_Tab_DBUpload.Cells(1, 1), xlWrkSht_Tab_DBUpload.Cells(16, 2))
 
 
'Load range into XML object
 xDocCMI.LoadXML rngCMI.Value(xlRangeValueMSPersistXML)
 xDocWorksheet_Tab_DBUpload.LoadXML rngWorksheet_Tab_DBUpload.Value(xlRangeValueMSPersistXML)
 
'Open recordset from XML
 rsCMI.Open xDocCMI, , adOpenStatic, adLockBatchOptimistic
 rsWorksheet_Tab_DBUpload.Open xDocWorksheet_Tab_DBUpload, , adOpenStatic, adLockBatchOptimistic
 
'Instantiate Public Recordsets used in the subroutine
 Set rsPUBLIC_CMIworksheet = New ADODB.Recordset
 Set rsPUBLIC_Worksheet_Tab_DBUpload = New ADODB.Recordset
 
'Disconnect the Public Recordsets
 rsPUBLIC_CMIworksheet.CursorLocation = adUseClient
 rsPUBLIC_Worksheet_Tab_DBUpload.CursorLocation = adUseClient
 
'Populate PUBLIC Recordsets with clone method
 Set rsPUBLIC_CMIworksheet = rsCMI.Clone
 Set rsPUBLIC_Worksheet_Tab_DBUpload = rsWorksheet_Tab_DBUpload.Clone

  
'----------->>> COPY DBUpload Tab into Spreadsheet <<<-----------
 '-----------------------Find DealNumber-----------
  Set rsFilter = rsWorksheet_Tab_DBUpload.Clone
  
  sFilterString = rsFilter.Fields(0).Name & "='DealNumber'"
  rsFilter.Filter = sFilterString


 '<<Set to Public variable Pub_DBUpload_DealNumber
  Pub_DBUpload_DealNumber = rsFilter.Fields(1).Value
  
  
 '-------------------------Find Tower Name------------------
  Set rsFilter = rsWorksheet_Tab_DBUpload.Clone
  
  sFilterString = rsFilter.Fields(0).Name & "='TowerName'"
  rsFilter.Filter = sFilterString
  
 '<<Set to Public variable Pub_DBUpload_TowerName
  Pub_DBUpload_TowerName = rsFilter.Fields(1).Value

  
  
 '-----------------------Find TemplateNumber-----------
  Set rsFilter = rsWorksheet_Tab_DBUpload.Clone
  
  sFilterString = rsFilter.Fields(0).Name & "='TemplateNumber'"
  rsFilter.Filter = sFilterString
 
 '<<Set to Public variable Pub_DBUpload_TemplateNumber<<
  Pub_DBUpload_TemplateNumber = rsFilter.Fields(1).Value
  
  
 '-----------------------Find FileName Number-----------
  Set rsFilter = rsWorksheet_Tab_DBUpload.Clone
  
  sFilterString = rsFilter.Fields(0).Name & "='FileName'"
  rsFilter.Filter = sFilterString
 
 '<<Set to Public variable Pub_DBUpload_TemplateNumber<<
  Pub_DBUpload_FileName = rsFilter.Fields(1).Value

    
ProcExit:

'Close the Sharepoint Excel Worksheet and do not save it
 If bClose_SP_Workbook = True Then
 
 'Close workbook
  xlWrkBk_SP.Close False
  
 End If

'Clear xDoc objects
 Set xDocCMI = Nothing
 Set xDocWorksheet_Tab_DBUpload = Nothing
 
'Close Recordset
 rsCMI.Close
 Set rsCMI = Nothing
 
 rsWorksheet_Tab_DBUpload.Close
 Set rsWorksheet_Tab_DBUpload = Nothing
 
 rsFilter.Close
 Set rsFilter = Nothing

 Exit Function

ProcErr:

  Select Case Err.Number
  
  Case 9 'Description Subscript out of range
    acOpenRecordset_from_ExcelWorksheet = False
    bClose_SP_Workbook = False
    MsgBox " This Cost Model does not have a Corp Model Input tab." & vbCrLf & vbCrLf & "Copy the data directly from the model into the appropriate section of the Export Corp tab!", vbInformation + vbOKOnly, "Corp Model Input NOT in this Cost Model"
    xlWrkBk_SP.Activate
    
    Resume ProcExit

  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case 3704 'Recordset is already closed
    Resume Next
    
  Case -2147467259 'Steam Object can't be read because it is empty
    acOpenRecordset_from_ExcelWorksheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
    Resume ProcExit

  Case Else
    acOpenRecordset_from_ExcelWorksheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit

End Function