' ========================================================================================
' Demonstrates the use of the getNamedItem method.
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
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pNodeBook AS IXMLDOMNode
   LOCAL pNodeId AS IXMLDOMNode
   LOCAL vIdValue AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.Load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeBook = pXmlDoc.selectSingleNode("//book")
         pNodeId = pNodeBook.attributes.getNamedItem("id")
         vIdValue = pNodeId.nodeValue
         AfxShowMsg VARIANT$$(vIdValue)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
