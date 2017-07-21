' ========================================================================================
' Demonstrates the use of createProcessingInstruction.
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

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pPri = pXmlDoc.createProcessingInstruction("xml", "version=""1.0""")
   AfxShowMsg pPri.xml

END FUNCTION
' ========================================================================================
