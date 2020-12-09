  $DOMDOC_LoadXML                 = "LOADXML"
  $DOMDOC_Load                    = "LOAD"
  $DOMDOC_TEXT                    = "TEXT"
  $DOMDOC_XML                     = "XML"
  $DOMDOC_COUNT                   = "COUNT"
  $DOMDOC_SetProperty             = "SETPROPERTY"
  $DOMDOC_GetProperty             = "GETPROPERTY"

  $DOMDOC_ParseError              = "PARSEERROR"
  $DOMDOC_ParseError_ErrorCode    = "ERRORCODE"
  $DOMDOC_ParseError_Reason       = "REASON"
  $DOMDOC_ParseError_SrcText      = "SRCTEXT"
  $DOMDOC_ParseError_Line         = "LINE"
  $DOMDOC_ParseError_LinePos      = "LINEPOS"
  $DOMDOC_ParseError_FilePos      = "FILEPOS"
  $DOMDOC_ParseError_Url          = "URL"

  $DOMDOC_GetElementsByTagName    = "GETELEMENTSBYTAGNAME"
  $DOMDOC_SelectSingleNode        = "SELECTSINGLENODE"
  $DOMDOC_SelectNodes             = "SELECTNODES"
  $DOMDOC_ChildNodes              = "CHILDNODES"

  '----------------------------------------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  ' Class
  '----------------------------------------------------------------------------
  Class cMsxml2_DOMDocument

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    Instance oDOMDoc    As IXMLDOMDocument3'Dispatch

    Instance vResult    As Variant
    Instance lResult    As Long

    '----------------------------------------------------------------------------
    '---Define here any Class Method (if needed)
    '---Class methods are visible only from inside the class
    '---so they can be called only from inside other methods.
    '----------------------------------------------------------------------------
    Class Method Create()
      Set oDOMDoc = newcom $PROGID_MSXML2_DOMDocument600'"Msxml2.DOMDocument.6.0"

      If Not IsObject(oDOMDoc) Then
        thinBasic_RunTimeError(%ERR__MODULE_SPECIFIC, "It was not possible to create a DOMDocument object (" & $PROGID_MSXML2_DOMDocument600 & ")")
      End If
    End Method

    Class Method Destroy()
      oDOMDoc = Nothing
    End Method

    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iMsxml2_DOMDocument

      Inherit IUnknown

      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method Msxml2_DOMDocument_GetClassObject() As Dword
        Method = VarPtr(oDOMDoc)
      End Method

      '----------------------------------------------------------------------------
      ' Creator Direct
      '----------------------------------------------------------------------------
      Method Msxml2_DOMDocument_Create_Direct(ByVal ptrToDirectRecordSet As Dword) As Dword
        '---Here is the magic: we store at the memory address of pRecordSet (that is a ADORecordset COM pointer)
        '---The address passed by the script. The address is usually returned by another Class method or property
        '---Instructed to return an OBJPTR to an ADORecordset
        Poke Dword, VarPtr(oDOMDoc), ptrToDirectRecordSet
      End Method

'      '----------------------------------------------------------------------------
'      '
'      '----------------------------------------------------------------------------
'      Method LoadXML(ByVal sXml As String) As Long
'        Local vXml As variant
'
'        vXml = sXml
'        Object Call oDOMDoc.LoadXml(vXML) To vResult
'
'      End Method
'
'      '----------------------------------------------------------------------------
'      '
'      '----------------------------------------------------------------------------
'      Method Text() As String
'
'        Object Call oDOMDoc.Text To vResult
'
'        Method = Variant$(vResult)
'      End Method

      '----------------------------------------------------------------------------
      '
      '----------------------------------------------------------------------------
'      Method GetElementsByTagName(ByVal sTag As String) As String
'        register nItem as long
'
'        Local vTag    As variant
'        Local ofields As Dispatch
'        Local sResult As String
'        Local vResult As Variant
'        Local nItems  As Long
'        Local v1      As variant
'
'        Set vTag = sTag
'        Object Call oDOMDoc.GetElementsByTagName(vTag) To vResult
'
'        Set oFields = vResult
'        sResult = Variant$$(vResult)
'
''MsgBox FuncName$ + ":" + sResult
'
'        Object Get oFields.length To vResult
'        nItems = Variant#(vresult)
'
'        For nItem = 0 To nItems -1
'            v1 = nItem
'            Object Get oFields.Item(v1).Text To vResult
'            If VariantVT(vResult) <> 0 Then
'              If nItem = 0 Then
'                sResult += Variant$(vResult)
'              Else
'                sResult += $CrLf & Variant$(vResult)
'              End If
'            End If
'        Next
'        oFields = Nothing
'
'        Method = sResult
'      End Method

    End Interface

  End Class

  '--------------------------------------------------------------------------------------
  ' We need to create a constructor wrapper function
  ' Constructor function will be executed by thinBasic when the class will be instantiated
  ' in script using the NEW keyword
  '--------------------------------------------------------------------------------------
  ' IMPORTANT:
  '   ALL wrapper functions MUST HAVE a "BYVAL pObject AS LONG" parameter.
  '   pObject will contain a pointer to the object of our class
  '   This pointer is initialized by thinBasic engine to hold object reference
  '--------------------------------------------------------------------------------------
  Sub Msxml2_DOMDocument_Create( ByVal pObject As Long )
    Local pp      As Byte

    pp = thinBasic_CheckOpenParens_Optional
    If pp Then thinBasic_CheckCloseParens_Mandatory

    If thinBasic_ErrorFree Then
      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject
      oObject(1&) = Class "cMsxml2_DOMDocument"
    End If

  End Sub

  '----------------------------------------------------------------------------
  Sub Msxml2_DOMDocument_Create_Direct(ByVal pObject As Dword, ByVal pDirectClass As Dword)
  '----------------------------------------------------------------------------
    Local pp          As Byte
    Local pRecordSet  As Ext

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cMsxml2_DOMDocument"
      '---We call sSet in order to assign initial value
      'oObject(1&).cExcel_Workbook_Create(pRecordSet)
      oObject(1&).Msxml2_DOMDocument_Create_Direct(pDirectClass)

    'else
      '---emit some error
    End If
  End Sub

  '--------------------------------------------------------------------------------------
   Sub Msxml2_DOMDocument_Destroy( ByVal pObject As Long )
  '--------------------------------------------------------------------------------------
    If pObject Then
      ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject
      Set oObject(1&) = Nothing
    End If

  End Sub

  '----------------------------------------------------------------------------
  Function Msxml2_DOMDocument_GetClassObject(ByVal pObject As Long) As Dword
  '----------------------------------------------------------------------------

    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject
      Function = oObject(1&).Msxml2_DOMDocument_GetClassObject()
    End If
  End Function
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------


  '--------------------------------------------------------------------------------------
  Function Msxml2_DOMDocument_ClassHandling ( _
                                                        ByVal pObject       As Long       , _
                                                        ByVal lAction       As Long       , _
                                                        ByVal sPreviousName As String     , _
                                              Optional  ByVal RetValNum     As Ext Ptr    , _
                                                        ByVal RetValStr     As String Ptr , _
                                                        ByVal lRecursive    As Long       , _
                                                        ByVal sParam1       As String     , _
                                                        ByVal sParam2       As String     , _
                                                        ByVal eParam1       As Ext        , _
                                                        ByVal eParam2       As Ext          _
                                            ) As Long
  '--------------------------------------------------------------------------------------
  ' IMPORTANT: function must return main return type of called method/property
  '--------------------------------------------------------------------------------------
    Local pp            As Long
    Local sName         As String
    Local sName2        As String
    Local sName3        As String

    Local sXml          As String
    Local sTagName      As String
    Local vResult       As Variant

    Local Counter       As Long
    Local sBuffer       As String
    
    Local lIdx          As Long
    
    Local sPropName     As String
    Local sPropValue    As String

    Static pNode        As IXMLDOMNode
    Static pNodeList    As IXMLDOMNodeList

    '---As usual we map our instance (whatever it will be) to a pObject pointer
    'ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject

    ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject
    ReDim pDomDocument(1& To 1&) As IXMLDOMDocument3 At oObject(1&).Msxml2_DOMDocument_GetClassObject

    '---We need to parse next token and we need to use special fucntion: thinBasic_GetUnknownToken
    '---This function is special because it instruct parser (thinCore) to not check for token validity
    '---but just parse something and return whetever it will find in terms of token string

    '   ATTENTION: thinBasic_GetUnknownToken always return UCASE tokens in order to simplify SELECT/CASE
    sName = thinBasic_GetUnknownToken

    '---Now we have to check parsed token using token content
    Select Case Const$ sName      '<<<---IMPORTANT: use CONST$ to speed up Select Case operations

      '-------------------------------------------------------
      '---Methods
      '---------------
      Case $DOMDOC_LoadXML
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sXML
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            @RetValNum = pDomDocument(1&).LoadXml(sXml)
          End If
        End If

      Case $DOMDOC_Load
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sXML
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            @RetValNum = pDomDocument(1&).Load(sXml)
          End If
        End If

      Case $DOMDOC_TEXT
        Function = %thinBasic_ReturnString
        @RetValStr = pDomDocument(1&).Text

      Case $DOMDOC_XML
        Function = %thinBasic_ReturnString
        @RetValStr = pDomDocument(1&).XML
      
      Case $DOMDOC_ParseError
        '---Do net iteration
        
        Case $DOMDOC_ParseError_ErrorCode
          Select Case Const$ sPreviousName
            Case $DOMDOC_ParseError
              Function = %thinBasic_ReturnNumber            
              @RetValNum = pDomDocument(1&).ParseError.errorCode
          End Select
        Case $DOMDOC_ParseError_Reason
          Select Case Const$ sPreviousName
            Case $DOMDOC_ParseError
              Function = %thinBasic_ReturnString
              @RetValStr = pDomDocument(1&).ParseError.Reason
          End Select
        Case $DOMDOC_ParseError_SrcText
          Select Case Const$ sPreviousName
            Case $DOMDOC_ParseError
              Function = %thinBasic_ReturnString
              @RetValStr = pDomDocument(1&).ParseError.srcText
          End Select
        Case $DOMDOC_ParseError_Line
          Select Case Const$ sPreviousName
            Case $DOMDOC_ParseError
              Function = %thinBasic_ReturnNumber
              @RetValNum = pDomDocument(1&).ParseError.Line
          End Select
        Case $DOMDOC_ParseError_LinePos
          Select Case Const$ sPreviousName
            Case $DOMDOC_ParseError
              Function = %thinBasic_ReturnNumber
              @RetValNum = pDomDocument(1&).ParseError.LinePos
          End Select
        Case $DOMDOC_ParseError_FilePos
          Select Case Const$ sPreviousName
            Case $DOMDOC_ParseError
              Function = %thinBasic_ReturnNumber
              @RetValNum = pDomDocument(1&).ParseError.FilePos
          End Select
        Case $DOMDOC_ParseError_Url
          Select Case Const$ sPreviousName
            Case $DOMDOC_ParseError
              Function = %thinBasic_ReturnString
              @RetValStr = pDomDocument(1&).ParseError.Url
          End Select

      Case $DOMDOC_SetProperty
        Function = %thinBasic_ReturnNumber
        pp = thinBasic_CheckOpenParens_Optional
          thinBasic_ParseString sPropName 
          if thinBasic_CheckComma_Mandatory then
            thinBasic_ParseString sPropValue
          End If
        If pp Then thinBasic_CheckCloseParens_Mandatory
        
        If thinBasic_ErrorFree Then
          If lRecursive Then
          Else
            pDomDocument(1&).SetProperty(sPropName, sPropValue)
          End If
        End If
      Case $DOMDOC_GetProperty
        Function = %thinBasic_ReturnString
        pp = thinBasic_CheckOpenParens_Optional
          thinBasic_ParseString sPropName 
        If pp Then thinBasic_CheckCloseParens_Mandatory
        
        If thinBasic_ErrorFree Then
          If lRecursive Then
          Else
            sBuffer = Variant$$(pDomDocument(1&).GetProperty(sPropName))
          End If
          @RetValStr = sBuffer
        End If
                
'      Case "LEN"
'        Function = %thinBasic_ReturnNumber
'        @RetValNum = Len(@RetValStr)

      Case $DOMDOC_GetElementsByTagName
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sTagName
        If pp Then thinBasic_CheckCloseParens_Mandatory

        If thinBasic_CheckOpenParens_Optional Then
          thinBasic_ParseLong lIdx
          thinBasic_CheckCloseParens_Mandatory
        End If  

        If thinBasic_CheckPoint_Optional Then
          sName2 = thinBasic_GetUnknownToken
        End If
        
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else 
'MsgBox "OK"
            '@RetValNum = pDomDocument(1&).GetElementsByTagName(sTagName).Length
'MsgBox "OK1:"

              pNodeList = pDomDocument(1&).GetElementsByTagName(sTagName)
              'For Counter = 0 To pNodeList.length - 1
              '  sBuffer += pNodeList.Item(Counter).nodeName + ":" & pNodeList.Item(Counter).Text & $CrLf
              'Next
'MsgBox "OK2:"
              If IsObject(pNodeList) Then
                If pNodeList.Length > 0 Then
'MsgBox "OK3:"    
                  lIdx = lIdx - 1
                  If lIdx <= 0 Or lIdx > pNodeList.Length + 1 Then lIdx = 0
                  
                  Select Case Const$ sName2
                    Case $DOMDOC_XML 
                      Function = %thinBasic_ReturnString
                      sBuffer = pNodeList.Item(lIdx).xml
                      @RetValStr = sBuffer
                    Case $DOMDOC_TEXT
                      Function = %thinBasic_ReturnString
                      sBuffer = pNodeList.Item(lIdx).Text
                      @RetValStr = sBuffer
                    Case $DOMDOC_COUNT
                      Function = %thinBasic_ReturnNumber
                      @RetValNum = pNodeList.Length
                    Case Else 
                      Function = %thinBasic_ReturnNumber
                      @RetValNum = VarPtr(pNodeList)
                  End Select
                End If
'MsgBox "OK4:"
              End If
          End If
        End If

      Case $DOMDOC_SelectSingleNode
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnString
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sTagName
        If pp Then thinBasic_CheckCloseParens_Mandatory

        If thinBasic_CheckPoint_Optional Then
          sName2 = thinBasic_GetUnknownToken
        End If
        
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else 
'MsgBox "OK1:" & sTagName
            '@RetValNum = pDomDocument(1&).GetElementsByTagName(sTagName).Length

            pNode = pDomDocument(1&).SelectSingleNode(sTagName)
              'For Counter = 0 To pNodeList.length - 1
              '  sBuffer += pNodeList.Item(Counter).nodeName + ":" & pNodeList.Item(Counter).Text & $CrLf
              'Next
              If IsObject(pNode) Then
'MsgBox "OK2"
                Select Case Const$ sName2
                  Case $DOMDOC_XML 
                    sBuffer = pNode.xml
                  Case $DOMDOC_TEXT
                    sBuffer = pNode.Text
                  Case Else 
                    Function = %thinBasic_ReturnNumber
                    @RetValNum = VarPtr(pNode)
                End Select
              End If
              @RetValStr = sBuffer
          End If
        End If

      Case $DOMDOC_ChildNodes
        Function = %thinBasic_ReturnNumber
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else 
'MsgBox "OK1:" & sTagName
            '@RetValNum = pDomDocument(1&).GetElementsByTagName(sTagName).Length

            pNodeList = pDomDocument(1&).ChildNodes
            @RetValNum = VarPtr(pNodeList)
          End If
        End If
      
      Case $DOMDOC_SelectNodes
        Function = %thinBasic_ReturnNumber
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sTagName
        If pp Then thinBasic_CheckCloseParens_Mandatory
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else 
'MsgBox "OK1:" & sTagName
            '@RetValNum = pDomDocument(1&).GetElementsByTagName(sTagName).Length

            pNodeList = pDomDocument(1&).SelectNodes(sTagName)
            @RetValNum = VarPtr(pNodeList)
          End If
        End If
      
              
      Case Else
        thinBasic_RunTimeError(%ERR__CLASS_METHODPROPERTY_NOTFOUND)
    End Select

    '-------------------------------------------------------
    '---Compound methods
    '---------------
    '   In case we have compound methods (defined as point after a method or property
    '   call this function recursively.
    '---------------
    If thinBasic_CheckPoint_Optional Then
      Function = Msxml2_DOMDocument_ClassHandling ( _
                                                  pObject   , _
                                                  lAction   , _
                                                  sName     , _
                                                  RetValNum , _
                                                  RetValStr , _
                                                  %TRUE       _
                                                )
    End If

  End Function

