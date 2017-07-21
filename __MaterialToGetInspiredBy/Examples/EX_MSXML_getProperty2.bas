' ========================================================================================
' Demonstrates the use of the getProperty (IXMLDOMDocument2) method.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument2
   LOCAL vValue AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pXmlDoc.setProperty "SelectionLanguage", "XPath"
   vValue = pXmlDoc.getProperty("SelectionLanguage")
   AfxShowMsg VARIANT$$(vValue)

END FUNCTION
' ========================================================================================
