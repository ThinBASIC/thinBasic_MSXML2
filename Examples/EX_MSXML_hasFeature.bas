' ========================================================================================
' Demonstrates the use of the hasFeature method.
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
   LOCAL iFeature AS INTEGER

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   iFeature = pXmlDoc.implementation.hasFeature("DOM", "1.0")
   AfxShowMsg "Has feature: " & FORMAT$(iFeature)

END FUNCTION
' ========================================================================================
