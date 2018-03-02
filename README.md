# ExcelVBA-LoadXML_toRecordset
 This function OpenRecordset_fromXML VBscript will parse an Excel workboook that is open in a web browser into an ADODB recordset using XML DOM.

 The application as whole will open an Excel workbook in a browser that resides on an SharePoint document library. The the web based Excel workbook contents are accessed using XML DOM (document object). The contents of the web page are parsed into an XML DOM object using the MSXML api within VBA. The library used for the MSXML is the Microsoft XML v6.0

 Once, the XML DOM object is created. It is then used by an ADODB recordset object as a data source. The recordset then is copied into the spreadsheet with the data from the Excel Workbook open in the browser.

 The XML DOM object is created in VBA through the create object method (see below example from code) 
 xDocCMI = CreateObject("MSXML2.DOMDocument"). 

 The MSXML2.DOMDocument object then can be read by the ADODB recordset (rsCMI is the recordset object) using the open recordset method. (see below code example)

 rsCMI.Open xDocCMI, , adOpenStatic, adLockBatchOptimistic

 Note: Make sure in Excel you have these libraries set in references tools -> references ->  Microsoft Scripting Runtime, Microsoft XML v6.0
