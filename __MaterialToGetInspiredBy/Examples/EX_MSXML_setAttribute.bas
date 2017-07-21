' ========================================================================================
' Demonstrates the use of the setAttribute method.
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
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL pDOMElement AS IXMLDOMElement
   LOCAL vValue AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDOMNode = pXmlDoc.selectSingleNode("//book")
         pDOMELEment = pDOMNode
         vValue = DATE$
         pDOMElEment.setAttribute "PublishDate", vValue
         vValue = pDOMElement.getAttribute("PublishDate")
         AfxShowMsg VARIANT$$(vValue)
         AfxShowMsg pDOMNode.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
