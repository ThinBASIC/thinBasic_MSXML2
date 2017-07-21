' ========================================================================================
' Demonstrates the use of the setAttributeNode method.
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
   LOCAL pNodeBook AS IXMLDOMNode
   LOCAL pDOMElement AS IXMLDOMElement
   LOCAL pNodePublishDate AS IXMLDOMAttribute
   LOCAL vValue AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodePublishDate = pXmlDoc.createAttribute("PublishDate")
         pNodePublishDate.nodeValue = DATE$
         pNodeBook = pXmlDoc.selectSingleNode("//book")
         pDOMELement = pNodeBook
         pDOMElement.setAttributeNode pNodePublishDate
         vValue = pDOMElement.getAttribute("PublishDate")
         AfxShowMsg VARIANT$$(vValue)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
