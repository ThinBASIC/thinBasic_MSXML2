' ========================================================================================
' Demonstrates the use of the implementation property.
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
   LOCAL pImplementation AS IXMLDOMImplementation

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISFALSE ISOBJECT(pXmlDoc) THEN EXIT FUNCTION

   pImplementation = pXmlDoc.implementation
   AfxShowMsg "Implementation reference pointer = " & STR$(OBJPTR(pImplementation))

END FUNCTION
' ========================================================================================
