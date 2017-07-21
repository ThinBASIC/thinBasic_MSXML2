' ========================================================================================
' Demonstrates the use of createProcessingInstruction.
' The following example specifies the target string "xml" and data string
' "version = '1.0'" to generate the processing instruction <?XML version="1.0"?>.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pPri AS IXMLDOMProcessingInstruction
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pItem AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML "<root><child/></root>"
      pPri = pXmlDoc.createProcessingInstruction("xml", "version=""1.0""")
      pNodeList = pXmlDoc.childNodes
      pItem = pNodeList.item(0)
      pXmlDoc.insertBefore pPri, pItem
      AfxShowMsg pXmlDoc.xml
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
