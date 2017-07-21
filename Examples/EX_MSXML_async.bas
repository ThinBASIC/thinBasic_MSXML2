' ========================================================================================
' Demonstrates the use of the async property.
' The following example sets the async property of a DOMDocument object to false before
' loading books.xml.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
'#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   pXmlDoc.load "books.xml"
   IF pXmlDoc.parseError.errorCode THEN
      AfxShowMsg "You have error " & pXmlDoc.parseError.reason
   ELSE
      AfxShowMsg pXmlDoc.xml
   END IF

END FUNCTION
' ========================================================================================
