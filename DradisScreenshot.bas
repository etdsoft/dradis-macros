Attribute VB_Name = "DradisScreenshot"
'
'The MIT License (MIT)
'Copyright (c) 2011 Daniel Martin <etd[- [at-]]nomejortu.com>
'
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
' http://www.holmessoft.co.uk/homepage/WininetVB.htm
'

Private Declare PtrSafe Function InternetOpen _
    Lib "wininet.dll" _
        Alias "InternetOpenA" _
            (ByVal lpszAgent As String, _
            ByVal dwAccessType As Long, _
            ByVal lpszProxyName As String, _
            ByVal lpszProxyBypass As String, _
            ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function InternetConnect _
    Lib "wininet.dll" _
        Alias "InternetConnectA" _
            (ByVal hInternetSession As Long, _
            ByVal lpszServerName As String, _
            ByVal nServerPort As Integer, _
            ByVal lpszUsername As String, _
            ByVal lpszPassword As String, _
            ByVal dwService As Long, _
            ByVal dwFlags As Long, _
            ByVal dwContext As Long) As Long
            
Private Declare PtrSafe Function HttpOpenRequest _
    Lib "wininet.dll" _
        Alias "HttpOpenRequestA" _
            (ByVal hHttpSession As Long, _
            ByVal lpszVerb As String, _
            ByVal lpszObjectName As String, _
            ByVal lpszVersion As String, _
            ByVal lpszReferer As String, _
            ByVal lpszAcceptTypes As String, _
            ByVal dwFlags As Long, _
            ByVal dwContext As Long) As Long
    
Private Declare PtrSafe Function HttpSendRequest _
    Lib "wininet.dll" _
        Alias "HttpSendRequestA" _
            (ByVal hHttpRequest As Long, _
            ByVal lpszHeaders As String, _
            ByVal dwHeadersLength As Long, _
            ByVal lpOptional As String, _
            ByVal dwOptionalLength As Long) As Boolean
    
Private Declare PtrSafe Function InternetSetOption _
    Lib "wininet.dll" _
        Alias "InternetSetOptionA" _
            (ByVal hInternet As Long, _
            ByVal dwOption As Long, _
            ByRef lpBuffer As Any, _
            ByVal dwBufferLength As Long) As Long
            
Private Declare PtrSafe Function InternetQueryOption _
    Lib "wininet.dll" _
        Alias "InternetQueryOptionA" _
            (ByVal hInternet As Long, _
            ByVal lOption As Long, _
            ByRef sBuffer As Any, _
            ByRef lBufferLength As Long) As Long
    
Private Declare PtrSafe Function InternetReadFile _
    Lib "wininet.dll" _
        (ByVal hFile As Long, _
        ByVal lpBuffer As String, _
        ByVal dwNumberOfBytesToRead As Long, _
        ByRef lpNumberOfBytesRead As Long) As Boolean
    
Private Declare PtrSafe Function InternetCloseHandle _
    Lib "wininet.dll" _
        (ByVal hInet As Long) As Integer
        
Private Const ERROR_INTERNET_INVALID_CA = 12045
        
Private Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
Private Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Private Const INTERNET_FLAG_NO_COOKIES = &H80000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_NO_UI = &H200
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_SECURE = &H800000

Private Const INTERNET_OPEN_TYPE_DIRECT = 1

Private Const INTERNET_OPTION_SECURITY_FLAGS = 31

Private Const INTERNET_SERVICE_HTTP = 3

Private Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100
Private Const SECURITY_FLAG_IGNORE_REVOCATION = &H80
'

' This method uses the win32 API to save a remote file to the local disk.
' I didn't include a lot of error-checking to keep the code short. If you run
' into issues, trace through this function, if hInternet, hConnect or hRequest
' become NULL, you're in trouble. Use the following snippet of code just after
' the call that returned 0 to get the corresponding error code and good luck 
' with MSDN!
'
' If Err.LastDllError <> 0 Then
'   MsgBox "wininet.dll error #" & Err.LastDllError
'   GoTo exitfunc
' End If
'
Public Sub WininetVB(sServer As String, iPort As Integer, sUser As String, sPassword As String, sURL As String, sFile As String)

Dim hInternet, hConnect, hRequest As Long
Dim lFlags As Long
Dim bRes As Boolean

Dim sBuffer As String * 1
Dim lBytesRead As Long


hInternet = InternetOpen("DradisMacro3", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
hConnect = InternetConnect(hInternet, sServer, iPort, sUser, sPassword, INTERNET_SERVICE_HTTP, 0, 0)

lFlags = INTERNET_FLAG_NO_COOKIES
lFlags = lFlags Or _
            INTERNET_FLAG_NO_CACHE_WRITE Or _
            INTERNET_FLAG_RELOAD
            
If (iPort = 443) Or (iPort = 3004) Then
    lFlags = lFlags Or _
            INTERNET_FLAG_SECURE Or _
            INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
            INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
End If

hRequest = HttpOpenRequest(hConnect, "GET", sURL, "HTTP/1.0", vbNullString, vbNullString, lFlags, 0)

If (iPort = 443) Or (iPort = 3004) Then
    bRet = InternetQueryOption(hRequest, INTERNET_OPTION_SECURITY_FLAGS, lFlags, Len(lFlags))
    lFlags = lFlags Or SECURITY_FLAG_IGNORE_UNKNOWN_CA Or SECURITY_FLAG_IGNORE_REVOCATION
    bRet = InternetSetOption(hRequest, INTERNET_OPTION_SECURITY_FLAGS, lFlags, Len(lFlags))
End If

bRes = HttpSendRequest(hRequest, vbNullString, 0, vbNullString, 0)

iFile = FreeFile()
Open sFile For Binary Access Write As iFile

Do
    bRes = InternetReadFile(hRequest, sBuffer, Len(sBuffer), lBytesRead)
    If lBytesRead > 0 Then
        Put iFile, , sBuffer
    End If
Loop While lBytesRead > 0

Close iFile

exitfunc:
    If hRequest <> 0 Then InternetCloseHandle (hRequest)
    If hConnect <> 0 Then InternetCloseHandle (hConnect)
    If hInternet <> 0 Then InternetCloseHandle (hInternet)
End Sub

' This method breaks a URL into its components
Sub ParseURL(sURL As String, ByRef sSchema As String, ByRef sDomain As String, ByRef iPort As Integer, ByRef sQuery As String)
    Dim urlParts
    Dim domainParts
    
    If InStr(sURL, "//") > 0 Then
        urlParts = Split(sURL, "/")
        sSchema = urlParts(0)
        sDomain = urlParts(2)
        sQuery = Right(sURL, Len(sURL) - (Len(sSchema) + 2) - Len(sDomain))
        
        If InStr(sDomain, ":") Then
            domainParts = Split(sDomain, ":")
            sDomain = domainParts(0)
            iPort = Val(domainParts(1))
        Else
            If (sSchema = "http:") Then
                iPort = 80
            Else
                iPort = 443
            End If
        End If
    Else
        MsgBox "Invalid URL", vbCritical
    End If
End Sub

' This is the main macro entry point
Sub DradisScreenshot()
    Dim sTmpDir As String
    Dim sURL As String
    Dim sServer As String
    Dim sPath As String
    Dim sFile As String
    Dim iPort As Integer
    Dim sDradisUser As String
    Dim sDradisPassword As String
    
    ' Adjust this path:
    sTmpDir = "C:\Users\nobody\Desktop\dradis\"

    ' Consider hard-coding the following if you always use the same user/pwd
    sDradisUser = InputBox("Your dradis user:  ")
    sDradisPassword = InputBox("The server password:  ")
    

    'First search the main document using the Selection
    With Selection
        With .Find
            .Text = "\!(<[fhpts]{3,5}://*>)\!"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Do While .Find.Execute
        
          sURL = Mid(Selection.Range.Text, 2, Len(Selection.Range.Text) - 2)
          
          ' Hopefully this method will be robust enough
          ParseURL sURL, vbNull, sServer, iPort, sPath
          
          sFile = sTmpDir & "\" & Mid(sPath, InStrRev(sPath, "/") + 1)
          
          ' Pull the file from the server and into sTmpDir
          WininetVB sServer, iPort, sDradisUser, sDradisPassword, sPath, sFile
          
          ' Now that the file is on disk, we can use the standard disalog to include it
          With Dialogs(wdDialogInsertPicture)
             .Name = sFile
             .Execute
          End With
        Loop
    End With
End Sub
