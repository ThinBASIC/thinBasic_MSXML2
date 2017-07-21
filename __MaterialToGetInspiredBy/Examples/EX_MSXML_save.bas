' =========================================================================================
' Demonstrates the use of the save method.
' The following example creates a DomDocument object from a string, then saves the
' document to a file in the application folder. If you look at the resulting file you will
' see that, instead of one continuous line of text, after each tag or data string. That is
' because of the $LF constant inserted in the string at the appropriate locations.
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.validateOnParse = %VARIANT_FALSE
      pXmlDoc.loadXML _
         "<?xml version='1.0'?>" + $LF + _
         "<doc title='test'>" + $LF + _
         "   <page num='1'>" + $LF + _
         "      <para title='Saved at last'>" + $LF + _
         "          This XML data is finally saved." + $LF + _
         "      </para>" + $LF + _
         "   </page>" + $LF + _
         "   <page num='2'>" + $LF + _
         "      <para>" + $LF + _
         "          This page is intentionally left blank." + $LF + _
         "      </para>" + $LF + _
         "   </page>" + $LF + _
         "</doc>" + $LF
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pXmlDoc.save "saved.xml"
         AfxShowMsg "Saved."
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
