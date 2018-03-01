# ExcelVBA-LoadXML_toRecordset
 This function OpenRecordset_fromXML VBscript will read and import Excel in a webpage format into an ADODB recordset using XML DOM.

 The application as whole parses a web based Excel workbook such as Excel Online or other web based forms of an Excel spreadsheet. The XML DOM (document object) is parsed using the MSXML api within VBA. The library used for the MSXML is the Microsoft XML v6.0

 The XML DOM object is created in VBA through the create object method (see below example from code) 
 xDocCMI = CreateObject("MSXML2.DOMDocument"). 

 The MSXML2.DOMDocument object then can be read by the ADODB recordset (rsCMI is the recordset object) using the open recordset method. (see below code example)

 rsCMI.Open xDocCMI, , adOpenStatic, adLockBatchOptimistic

 Note: Make sure in Excel you have these libraries set in references tools -> references ->  Microsoft Scripting Runtime, Microsoft XML v6.0
