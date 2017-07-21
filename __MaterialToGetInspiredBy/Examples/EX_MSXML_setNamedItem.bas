' =========================================================================================
' Demonstrates the use of the setNamedItem method.
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument2
   LOCAL pNodePublishDate AS IXMLDOMAttribute
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pNodeBook AS IXMLDOMNode
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
         pNodePublishDate = pXmlDoc.createAttribute("PublishDate")
         pNodePublishDate.value = DATE$
         pNodeBook = pXmLDoc.selectSingleNode("//book")
         pNamedNodeMap = pNodeBook.attributes
         pNamedNodeMap.setNamedItem pNodePublishDate
         pDOMElement = pNodeBook
         vValue = pDOMElement.getAttribute("PublishDate")
         AfxShowMsg VARIANT$$(vValue)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
