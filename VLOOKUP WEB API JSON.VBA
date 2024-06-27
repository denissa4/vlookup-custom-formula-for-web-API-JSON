    '------------------------------------------------------------------------
    'This function is property of NLSQL Limited,
    'NLSQL is B2B SaaS to empower employees with intuitive text interface to
    'poorly-accessible coporate data to inform and speed business decisions with significant benefits for enterprises.
    'Below code is 100% open-source under MIT Licence, so you can feel free to use it for
    'Commercial use, Modification, Distribution or private use. Except if you are Microsoft =)
    'NLSQL Limited don't have any Liability or Warranties related to your code usage
    'subscribe to our LinkedIn https://www.linkedin.com/company/nlsql-com
    'or Youtube https://www.youtube.com/channel/UC8KtzeNHxhLGVwiOCwvRBkg?sub_confirmation=1
    'in order to have even more great open-source tools, absolutely free
    
    'By Denis
    'Date: 06/09/2021
    'info@nlsql.com
    'https://nlsql.com/////
    '------------------------------------------------------------------------

'Function works similar to standard Excel VLOOKUP function,
'so it finds the FIRST value based on API connection and field name from API responce

Function VLOOKUPWEB(ByVal cell As Range, ByVal txt As String, Optional ByVal pause, _
                    Optional ByVal header As String, Optional ByVal postcell As Range) As String ' As Integer
    On Error Resume Next
    Dim GetResponse As String
    FuncDesc = "Function created for Vlookup JSON to excel in a most simpliest way same as using VLOOKUP"
    URL = cell.Value    'URL is required to have correct API link including https or http. HTTPS works better for the data security
    GetResponse = ""
        If URL Like "*http*" Then
        
            If Not postcell Is Nothing Then
                pst = postcell.Value
                zapr = "POST"   'support of POST request. Body (5th) Postcell argument is required
            Else
                pst = ""
                zapr = "GET"   'support of GET request. All parameters expected to be in (1st) Cell argument
            End If
            
            Set xmlhttp = CreateObject("MSXML2.XMLHTTP") 'MSXML2.XMLHTTP
             xmlhttp.Open zapr, URL, False: DoEvents
             xmlhttp.setRequestHeader "Authorization", header   'you can change Authorization to api-key, etc. as required by API provider
             xmlhttp.setRequestHeader "Content-Type", "application/json"   'you can add as much headers as needed
             xmlhttp.send pst  'send request both POST and GET supported
             GetResponse = xmlhttp.responseText
             'MsgBox GetResponse
             Set xmlhttp = Nothing
        End If

        If pause > 0 Then
            t = Timer '3rd argument - Pause in seconds for time-out
            Do
                DoEvents
                If t > Timer Then Exit Do
            Loop Until Timer - t > pause
        End If
          
    s = Split(GetResponse, txt)

        If UBound(s) > 0 Then
           
            ps = Split(s(1), """")
            VLOOKUPWEB = ps(2)
    
        Else:
            VLOOKUPWEB = "not available":
    
        End If
        
      
   
End Function
