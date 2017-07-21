' ========================================================================================
' Demonstrates the use of the NewParser property.
' The following example sets the ProhibitDTD and the validateOnParse properties to false
' to allow to use MSXML 6.0 with the new parser
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

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.setProperty "NewParser", %VARIANT_TRUE
      pXmlDoc.setProperty "ProhibitDTD", %VARIANT_FALSE
      pXmlDoc.validateOnParse = %VARIANT_FALSE
      pXmlDoc.loadXML "<?xml version='1.0' ?><!DOCTYPE root [<!ELEMENT root (#PCDATA)>]><root a='bc'>abc</root>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         AfxShowMsg pXmlDoc.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
