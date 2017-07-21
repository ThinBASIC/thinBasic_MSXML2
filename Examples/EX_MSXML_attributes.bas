' ========================================================================================
' Demonstrates the use of the attributes property.
' The following example creates an IXMLDOMNamedNodeMap object from a document's attributes
' property, and then displays the number of nodes in the object.
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
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   pXmlDoc.load "books.xml"
   IF pXmlDoc.parseError.errorCode THEN
      AfxShowMsg "You have error " & pXmlDoc.parseError.reason
   ELSE
      pNamedNodeMap = pXmlDoc.documentElement.firstChild.attributes
      AfxShowMsg "Length = " & FORMAT$(pNamedNodeMap.length)
   END IF

END FUNCTION
' ========================================================================================
