  '----------------------------------------------------------------------------
  ' Class cMsxml2_XMLHTTP implements Microsoft XMLHTTP class
  '----------------------------------------------------------------------------
  Class cMsxml2_ServerXMLHTTP

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    Instance pHTTPServer  As IServerXMLHTTPRequest2

    Instance vResult  As Variant
    Instance lResult  As Long

    '----------------------------------------------------------------------------
    '---Define here any Class Method (if needed)
    '---Class methods are visible only from inside the class
    '---so they can be called only from inside other methods.
    '----------------------------------------------------------------------------
    Class Method Create()
      pHTTPServer     = ANYCOM $PROGID_MSXML2_ServerXMLHTTP600

      If not IsObject(pHTTPServer) Then
        thinBasic_RunTimeError(%ERR__MODULE_SPECIFIC, "It was not possible to create a ServerXMLHTTP object " & $PROGID_MSXML2_ServerXMLHTTP600 & ")")
      End If
    End Method

    Class Method Destroy()
      pHTTPServer = Nothing
    End Method

    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iMsxml2_ServerXMLHTTP

      Inherit IUnknown

      '----------------------------------------------------------------------------
      ' Creator
      '----------------------------------------------------------------------------
      Method cServerXMLHTTP_Create() As Dword
        pHTTPServer = ANYCOM $PROGID_MSXML2_ServerXMLHTTP600
        If IsNothing(pHTTPServer) Then Exit Method
      End Method

      '----------------------------------------------------------------------------
      ' Destructor
      '----------------------------------------------------------------------------
      Method cServerXMLHTTP_Destroy() As Dword
        Set pHTTPServer = Nothing
      End Method

      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method cServerXMLHTTP_GetClassObject() As Dword
        Method = VarPtr(pHTTPServer)
      End Method

      '----------------------------------------------------------------------------
      ' IsNothing
      '----------------------------------------------------------------------------
      Method cServerXMLHTTP_IsNothing() As Dword
        Method = IsNothing(pHTTPServer)
      End Method

      '----------------------------------------------------------------------------
      ' IsObject
      '----------------------------------------------------------------------------
      Method cServerXMLHTTP_IsObject() As Dword
        Method = IsObject(pHTTPServer)
      End Method

    End Interface

  End Class
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------


  '--------------------------------------------------------------------------------------
  '---WRAPPERS to simulate an object using a PTR
  '--------------------------------------------------------------------------------------
  ' Here we are. In order to let thinBasic know about a class, we need to wrap
  ' some, or all of the class interface methods defined above.
  ' This will also let us choose any syntax of our method
  '--------------------------------------------------------------------------------------

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
  Sub Msxml2_ServerXMLHTTP_Create( ByVal pObject As Long )
    Local pp      As Byte

    '---Remember: programmer is responsible of syntax so wrapper method MUST parse syntax

    '---Optionally checks if there is a (
    pp = thinBasic_CheckOpenParens_Optional
    If pp Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iMsxml2_ServerXMLHTTP At pObject

      oObject(1&) = Class "cMsxml2_ServerXMLHTTP"
      '---We call internal creator
      oObject(1&).cServerXMLHTTP_Create()

    End If

  End Sub


  '--------------------------------------------------------------------------------------
   Sub Msxml2_ServerXMLHTTP_Destroy( ByVal pObject As Long )
  '--------------------------------------------------------------------------------------

    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iMsxml2_ServerXMLHTTP At pObject
      oObject(1&).cServerXMLHTTP_Destroy()
    End If

  End Sub

  '----------------------------------------------------------------------------
  Function Msxml2_ServerXMLHTTP_GetClassObject(ByVal pObject As Long) As Dword
  '----------------------------------------------------------------------------
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iMsxml2_ServerXMLHTTP At pObject
      Function = oObject(1&).cServerXMLHTTP_GetClassObject()
    End If
  End Function


  '--------------------------------------------------------------------------------------
  Function Msxml2_ServerXMLHTTP_ClassHandling ( _
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

    Local sMethod       As String
    Local sUrl          As String

    Local sHeader       As String
    
    Local eAsync        As Ext
    Local sUser         As String
    Local sPassword     As String
      Local eAsync_Present     As Byte
      Local sUser_Present     As Byte
      Local sPassword_Present As Byte

    Local sValue        As String
    Local sBody         As String
    Local eresolveTimeout  As Ext
    Local econnectTimeout  As Ext
    Local esendTimeout     As Ext
    Local ereceiveTimeout  As Ext
    Local eWaitSec      As Ext
    Local vWaitSec      As variant


    '---We map our instance (whatever it will be) to a pObject pointer
    ReDim oObject(1& To 1&) As iMsxml2_ServerXMLHTTP At pObject
    '---Here we get internal real class object and we will use it directly from now on
    ReDim lServerXMLHTTP(1& To 1&) As IServerXMLHTTPRequest2 At oObject(1&).cServerXMLHTTP_GetClassObject

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
      Case "STATUS"
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory

        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---If all went fine ...  
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else
            @RetValNum = lServerXMLHTTP(1&).Status
          End If
        End If

      Case "STATUSTEXT"
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnString
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else
            @RetValStr = lServerXMLHTTP(1&).StatusText
          End If
        End If

      Case "READYSTATE"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory
        Function = %thinBasic_ReturnNumber
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else
            @RetValNum = lServerXMLHTTP(1&).ReadyState
          End If
        End If

      Case "OPEN"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sMethod
        If thinBasic_CheckComma_Mandatory Then
          thinBasic_ParseString sUrl

          eAsync = %FALSE
          If thinBasic_CheckComma_optional Then
            eAsync_Present = %TRUE
            thinBasic_ParseNumber eAsync
            If thinBasic_CheckComma_Optional Then
              sUser_Present = %TRUE
              thinBasic_ParseString sUser
              If thinBasic_CheckComma_Optional Then
                sPassword_Present = %TRUE
                thinBasic_ParseString sPassword
              End If
            End If
          End If

        End If

        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            If sPassword_Present = %TRUE Then
              lServerXMLHTTP(1&).Open(sMethod, sUrl, eAsync, sUser, sPassword)
            ElseIf sUser_Present = %TRUE Then           
              lServerXMLHTTP(1&).Open(sMethod, sUrl, eAsync, sUser)
            ElseIf eAsync_Present = %TRUE Then
              lServerXMLHTTP(1&).Open(sMethod, sUrl, eAsync)
            Else
              lServerXMLHTTP(1&).Open(sMethod, sUrl)
            End If
                          
            '    Method open <1> (Byval bstrMethod As WString, Byval bstrUrl As WString, Opt Byval varAsync As Variant, Opt Byval _
            '        bstrUser As Variant, Opt ByVal bstrPassword As Variant)
          End If
        End If

      Case "SETREQUESTHEADER"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sHeader
        If thinBasic_CheckComma_Mandatory Then
          thinBasic_ParseString sValue
        End If
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            '@RetValNum =
            lServerXMLHTTP(1&).SetRequestHeader(sHeader, sValue)
          End If
        End If

      Case "SEND"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sBody
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            '@RetValNum =
            lServerXMLHTTP(1&).Send(sBody)
          End If
        End If

      Case "ABORT"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            lServerXMLHTTP(1&).Abort
          End If
        End If

      Case "RESPONSETEXT"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnString
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else
            @RetValStr = lServerXMLHTTP(1&).ResponseText
          End If
        End If

      Case "SETTIMEOUTS"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        thinBasic_Parse4Numbers(eresolveTimeout, econnectTimeout, esendTimeout, ereceiveTimeout)
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            '@RetValNum =
            lServerXMLHTTP(1&).setTimeouts(eresolveTimeout, econnectTimeout, esendTimeout, ereceiveTimeout)
          End If
        End If

      Case "GETALLRESPONSEHEADERS" ' UCase$("getAllResponseHeaders")
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory
        Function = %thinBasic_ReturnString
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else
            @RetValStr = lServerXMLHTTP(1&).getAllResponseHeaders
          End If
        End If

      Case "GETRESPONSEHEADER" ' UCase$("getResponseHeader")
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnString
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sBody
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else
            @RetValStr = lServerXMLHTTP(1&).getResponseHeader(sBody)
          End If
        End If


      Case "ISOBJECT"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory
        Function = %thinBasic_ReturnNumber
        '---parse Params
        If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory
        @RetValNum = oObject(1&).cServerXMLHTTP_IsObject

      Case "ISNOTHING"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory
        Function = %thinBasic_ReturnNumber
        '---parse Params
        If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory
        @RetValNum = oObject(1&).cServerXMLHTTP_IsNothing

      Case "WAITFORRESPONSE"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseNumber eWaitSec
        Let vWaitSec = eWaitSec As Long
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            @RetValNum = lServerXMLHTTP(1&).WaitForResponse(eWaitSec)
'MsgBox "WaitForResponse" & Str$(eWaitSec)
          End If
        End If

      Case "GETOPTION"
        Local lGetOption As Long

        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseLong lGetOption
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else                     
            Select Case lGetOption
              Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_URL
                Function = %thinBasic_ReturnString
                @RetValStr = Variant$(lServerXMLHTTP(1&).getOption(lGetOption))
                
              Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_URL_CODEPAGE
                Function = %thinBasic_ReturnNumber
                @RetValNum = Variant#(lServerXMLHTTP(1&).getOption(lGetOption))

              Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_ESCAPE_PERCENT_IN_URL
                Function = %thinBasic_ReturnNumber
                @RetValNum = Variant#(lServerXMLHTTP(1&).getOption(lGetOption))

              Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
                Function = %thinBasic_ReturnNumber
                @RetValNum = Variant#(lServerXMLHTTP(1&).getOption(lGetOption))

              Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_SELECT_CLIENT_SSL_CERT
                Function = %thinBasic_ReturnString
                @RetValStr = Variant$(lServerXMLHTTP(1&).getOption(lGetOption))

            End Select
          End If
        End If

      Case "SETOPTION"
        Local lSetOption        As Long
        Local sSetOptionValue   As String
        Local lSetOptionValue   As Long
        Local vSetOptionVariant As variant

        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseLong lSetOption
        thinBasic_CheckComma_Mandatory

          '---If all went fine ...
          If thinBasic_ErrorFree Then
            If lRecursive Then
              '@RetValStr = Mid$(@RetValStr, eStart, eLen)
            Else                     
              Select Case lSetOption
                'Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_URL
                '  Function = %thinBasic_ReturnString
                '  @RetValStr = Variant$(lServerXMLHTTP(1&).getOption(lGetOption))
                  
                Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_URL_CODEPAGE
                  Function = %thinBasic_ReturnNumber
                  thinBasic_ParseLong lSetOptionValue
                  vSetOptionVariant = Variant#(lSetOptionValue)
                  lServerXMLHTTP(1&).SetOption(lSetOption, vSetOptionVariant)

                Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_ESCAPE_PERCENT_IN_URL
                  Function = %thinBasic_ReturnNumber
                  thinBasic_ParseLong lSetOptionValue
                  vSetOptionVariant = Variant#(lSetOptionValue)
                  lServerXMLHTTP(1&).SetOption(lSetOption, vSetOptionVariant)
  
                Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
                  Function = %thinBasic_ReturnNumber
                  thinBasic_ParseLong lSetOptionValue
                  vSetOptionVariant = Variant#(lSetOptionValue)
                  lServerXMLHTTP(1&).SetOption(lSetOption, vSetOptionVariant)

                Case %Enum__SERVERXMLHTTP_OPTION.SXH_OPTION_SELECT_CLIENT_SSL_CERT
                  Function = %thinBasic_ReturnNumber
                  thinBasic_ParseString sSetOptionValue
                  vSetOptionVariant = Variant$(sSetOptionValue)
                  lServerXMLHTTP(1&).SetOption(lSetOption, vSetOptionVariant)
  
              End Select
            End If
          End If
        If pp Then thinBasic_CheckCloseParens_Mandatory


'%SXH_OPTION_ESCAPE_PERCENT_IN_URL
'%SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
'%SXH_OPTION_SELECT_CLIENT_SSL_CERT
'%SXH_OPTION_URL
'%SXH_OPTION_URL_CODEPAGE

'%SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
'%SXH_SERVER_CERT_IGNORE_CERT_CN_INVALID
'%SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID
'%SXH_SERVER_CERT_IGNORE_UNKNOWN_CA
'%SXH_SERVER_CERT_IGNORE_WRONG_USAGE




'      '-------------------------------------------------------
'      '---Properties
'      '---------------
'      Case "VALUE"  '---Support GET and SET. Programmer is free to support whatever event type
'
'        '---In all "value" cases we will return a string, so we can assign return type here
'        Function = %thinBasic_ReturnString
'
'        Select Case Long lAction
'          Case %Class_Action_Get
'            @RetValStr = oObject(1&).sGet
'
'          Case %Class_Action_Set
'            lEqualType = thinBasic_CheckEqualType_Mandatory
'            If thinBasic_ErrorFree Then
'              Select Case Long lEqualType
'
'                '---Assignment
'                Case  %Equal_EQ
'                      oObject(1&).sSet(thinBasic_Parse1String)
'
'                '---Assignment with concatenation
'                Case  %Equal_PLUSEQUAL, _
'                      %Equal_CONCEQUAL, _
'                      %Equal_POINTEQUAL
'
'                      oObject(1&).sSet(oObject(1&).sGet() & thinBasic_Parse1String)
'                Case Else
'                  '---Not supperted for strings
'                  'MsgBox FuncName$ & Str$(lEqualType)
'              End Select
'            End If
'
'        End Select

      Case Else
        thinBasic_RunTimeError(%ERR__CLASS_METHODPROPERTY_NOTFOUND)
    End Select

    '-------------------------------------------------------
    '---Compound methods
    '---------------
    '   In case we have compound methods (defined as point after a method or property
    '   call this function recursively.
    '   Example: myObject.MyMethod.MyCompoundSubMethod
    '---------------
    If thinBasic_CheckPoint_Optional Then
      Function = Msxml2_ServerXMLHTTP_ClassHandling ( _
                                                  pObject   , _
                                                  lAction   , _
                                                  sName     , _
                                                  RetValNum , _
                                                  RetValStr , _
                                                  %TRUE       _
                                                )
    End If

  End Function
