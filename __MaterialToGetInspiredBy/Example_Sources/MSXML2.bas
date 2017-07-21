' Small XHR taster by Chris Holbrook
' May 25 2012
' To show how XMLHttpRequest makes light work of handling XML from web servers
' compiles with PBWin V9 or PBWIn V10
' Using Dispatch interfaces.
'
#Compile Exe
#Dim All

'$url = "http://www.w3schools.com/xml/plant_catalog.xml"

''-----------------------------------------------------------------------------------
'' Get the text for ech tag. Ignores child tags, just extracts all the text for tag
'' and its children.
''
'Function GetTextFromXML (sxml As String,  _ In XML To Parse
'                         snodeids As String _ In CSV list Of node ids Of interest
'                         ) As String
'    Local v1, v2, v3 As variant
'    Local vnodename, vresult As variant
'    Local oDOM, ofields, ofield As Dispatch
'    Local i, j, n As Long
'    Local stag As String
'    Local sresult As String
'    Set oDOM = newcom "Microsoft.XMLDOM"'.1.0"
'
'    v1 = sxml
'    Object Call oDOM.LoadXML(v1)
'    '
'    For j = 1 To ParseCount(snodeids)
'        stag = Parse$(snodeids,j)
'        v1 = stag
'        Object Call oDOM.getElementsByTagName(v1) To vresult
'        Set oFields = vresult
'        '
'        Object Get oFields.length To vResult
'        n = Variant#(vresult)
'        '
'        For i = 0 To n -1
'            v1 = i
'            Object Get oFields.Item(v1).Text To vresult
'            If VariantVT(vresult) <> 0 Then
'                sresult += $CrLf + stag + "," + Variant$(vResult)
'            End If
'        Next
'    Next
'    Function = sresult
'    oDOM = Nothing
'    Exit Function
'End Function


Function GetAuth (sxml As String) As String
    Local v1, v2, v3 As variant
    Local vnodename, vresult As variant
    Local oDOM, ofields, ofield As Dispatch
    Local i, j, n As Long
    Local stag As String
    Local sresult As String 
    
    Set oDOM = newcom "Microsoft.XMLDOM"'.1.0"

    v1 = sxml
    Object Call oDOM.LoadXML(v1)
    '
    'For j = 1 To ParseCount(snodeids)
        stag = "authenticationToken"'Parse$(snodeids,j)
        v1 = stag
        Object Call oDOM.getElementsByTagName(v1) To vresult
        Set oFields = vresult
        sresult = Variant$(vresult)
        ''
        Object Get oFields.length To vResult
        n = Variant#(vresult)
        '
        For i = 0 To n -1
            v1 = i
            Object Get oFields.Item(v1).Text To vresult
            If VariantVT(vresult) <> 0 Then
                'sresult += $CrLf + stag + "," + Variant$(vResult)
                sresult = Variant$(vResult)
            End If
        Next 
        oFields = Nothing
    'Next
    Function = sresult
    oDOM = Nothing
    'Exit Function
End Function



'---------------------------------------------------------------------
Function PBMain As Long
    Local v1, v2, v3 As variant
    Local vresult As variant
    Local vresponse As variant
    Local lresult As Long
    Local s As String
    Local sHttp_Buffer As String
    Local sAuth As String

    Local oHTTP As Dispatch
    '
    Set oHTTP = newcom "Msxml2.XMLHTTP"
    '

    v1 = "POST"
    v2 = "http://erbomnet01.erbolario.erbolario.net/magonet/loginmanager/loginmanager.asmx"
    v3 = 0
    'v4 = $USERID
    'v5 = $PASSWORD
    Object Call oHTTP.Open(v1, v2, v3) To vresult
    v1 = "Content-Type"
    v2 = "text/xml; charset=utf-8"
    Object Call oHTTP.setRequestHeader (v1, v2) To vresult
    v1 = "SOAPAction"
    v2 = """http://microarea.it/LoginManager/Login"""
    Object Call oHTTP.setRequestHeader (v1, v2) To vresult
    '
    sHttp_Buffer  =          "<?xml version=""1.0"" encoding=""utf-8""?>"
    sHttp_Buffer += $CrLf &  "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
    sHttp_Buffer += $CrLf &  "  <soap:Body>"
    sHttp_Buffer += $CrLf &  "    <Login xmlns=""http://microarea.it/LoginManager/"">"
    sHttp_Buffer += $CrLf &  "      <companyName>MNTestPartizione</companyName>"
    sHttp_Buffer += $CrLf &  "      <userName>sa</userName>"
    sHttp_Buffer += $CrLf &  "      <password>Service</password>"
    sHttp_Buffer += $CrLf &  "      <askingProcess>EROSTEST</askingProcess>"
    sHttp_Buffer += $CrLf &  "      <overWriteLogin>1</overWriteLogin>"
    sHttp_Buffer += $CrLf &  "    </Login>"
    sHttp_Buffer += $CrLf &  "  </soap:Body>"
    sHttp_Buffer += $CrLf &  "</soap:Envelope>"

    v1 = sHttp_Buffer
    Object Call oHTTP.Send(v1) To vresult
    Object Get oHTTP.Responsetext To vresponse
    '
    '? Variant$(vresponse)
    oHTTP      = Nothing
    s = Variant$(vresponse)
    ? s

    
    sAuth = GetAuth(s)
    ? sAuth
    
    


    Set oHTTP = newcom "Msxml2.XMLHTTP"

    v1 = "POST"
    v2 = "http://erbomnet01.erbolario.erbolario.net/magonet/loginmanager/loginmanager.asmx"
    v3 = 0
    'v4 = $USERID
    'v5 = $PASSWORD
    Object Call oHTTP.Open(v1, v2, v3) To vresult
    v1 = "Content-Type"
    v2 = "text/xml; charset=utf-8"
    Object Call oHTTP.setRequestHeader (v1, v2) To vresult
    v1 = "SOAPAction"
    v2 = """http://microarea.it/LoginManager/LogOff"""
    Object Call oHTTP.setRequestHeader (v1, v2) To vresult
    '
    sHttp_Buffer  =          "<?xml version=""1.0"" encoding=""utf-8""?>"
    sHttp_Buffer += $CrLf &  "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
    sHttp_Buffer += $CrLf &  "  <soap:Body>"
    sHttp_Buffer += $CrLf &  "    <LogOff xmlns=""http://microarea.it/LoginManager/"">"
    sHttp_Buffer += $CrLf &  "      <authenticationToken>"   & sAuth  & "</authenticationToken>"
    sHttp_Buffer += $CrLf &  "    </LogOff>"
    sHttp_Buffer += $CrLf &  "  </soap:Body>"
    sHttp_Buffer += $CrLf &  "</soap:Envelope>"

    v1 = sHttp_Buffer
    Object Call oHTTP.Send(v1) To vresult
    Object Get oHTTP.Responsetext To vresponse
    '
    '? Variant$(vresponse)
    oHTTP      = Nothing
    s = Variant$(vresponse)
    ? s
    
    
    
    
    
    '? getTextfromXML(s, "PLANT,COMMON,BOTANICAL")
    Exit Function
End Function
