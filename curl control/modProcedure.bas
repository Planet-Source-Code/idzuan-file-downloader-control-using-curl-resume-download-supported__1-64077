Attribute VB_Name = "modProcedure"
' $Id: EasyGet.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' Modified on 2006/01/17 00:05:27 by AliasMohdIdzuan
' Demonstrate Easy Get Capability
' CAUTION : USES CURL LIBRARY

' Make sure reference to VBVM6Lib.tlb, vblibcurl.tlb exists
' Put libcurl.dll, vblibcurl.dll in current dir path
' For more info, visit http://curl.haxx.se/libcurl/vb/

' Plz vote for me :)

Public actfilesize As Double
Public bar As Variant
Public buf As String
Public c As Integer
Public filename As String
Public filesize As Double
Public bytes() As Variant
Public stopcurl As Boolean


Public m_itemlabel As String
Public m_file_url As String
Public m_savefileas As String
Public m_proxy_url As String
Public m_use_post As Integer
Public m_post_fields As String
Public m_progress As Integer
Public m_kbps As Double

Private easy As Long

Public Sub EasyGet()
    
    Dim ret As CURLcode
    Dim buf As New Buffer

    c = 0
    Erase bytes()
    easy = vbcurl_easy_init()
    If easy = 0 Then Exit Sub   ' if handle to easy=0(no internet connection), abort file download
    filesize = CheckFile(m_savefileas)
    'vbcurl_easy_setopt easy, CURLOPT_FRESH_CONNECT, True   'Set to use fresh connection every time
    'vbcurl_easy_setopt easy, CURLOPT_CONNECTTIMEOUT, 1     'Time out in second
    If filesize > 0 Then
        vbcurl_easy_setopt easy, CURLOPT_RESUME_FROM, filesize  'Resume file download start chunk from position XXX bytes
    End If
    vbcurl_easy_setopt easy, CURLOPT_USERAGENT, "libcurl-agent/1.0"
    vbcurl_easy_setopt easy, CURLOPT_BUFFERSIZE, 500        'File buffer size
    vbcurl_easy_setopt easy, CURLOPT_LOW_SPEED_LIMIT, 10    'Time lower limit
    vbcurl_easy_setopt easy, CURLOPT_LOW_SPEED_TIME, 5      'Download lower speed limit
    If m_proxy_url <> "" Then
        vbcurl_easy_setopt easy, CURLOPT_PROXY, m_proxy_url 'If using proxy server, format http://proxyserver:portno
    End If
    vbcurl_easy_setopt easy, CURLOPT_AUTOREFERER, True      'Set to follow referer location
    vbcurl_easy_setopt easy, CURLOPT_FOLLOWLOCATION, True   'Set to follow location
    vbcurl_easy_setopt easy, CURLOPT_URL, m_file_url        'File download URL
    If m_use_post = 1 Then
        vbcurl_easy_setopt easy, CURLOPT_POST, 1            'Set to 0 if no POST field append
        vbcurl_easy_setopt easy, CURLOPT_POSTFIELDS, m_post_fields  'Your field submit format "a=1&b=2&c=3"
    End If
                                                
    vbcurl_easy_setopt easy, CURLOPT_WRITEDATA, ObjPtr(buf) 'Get return value here...
    vbcurl_easy_setopt easy, CURLOPT_WRITEFUNCTION, _
        AddressOf WriteFunction
    'vbcurl_easy_setopt Easy, CURLOPT_DEBUGFUNCTION, _
    '    AddressOf DebugFunction
    'vbcurl_easy_setopt Easy, CURLOPT_VERBOSE, True
    vbcurl_easy_setopt easy, CURLOPT_NOPROGRESS, False      'Get file transfer progress here...
    vbcurl_easy_setopt easy, CURLOPT_PROGRESSFUNCTION, _
        AddressOf ProgressFunction
    vbcurl_easy_setopt easy, CURLOPT_PROGRESSDATA, bar      'Get progress data
    
    ret = vbcurl_easy_perform(easy)
    vbcurl_easy_cleanup easy
        
    If CheckFile(m_savefileas) > 0 Then
    AppendFileByte
        Else
    AddFileByte
    End If
    
End Sub

' This function illustrates a couple of key concepts in libcurl.vb.
' First, the data passed in rawBytes is an actual memory address
' from libcurl. Hence, the data is read using the MemByte() function
' found in the VBVM6Lib.tlb type library. Second, the extra parameter
' is passed as a raw long (via ObjPtr(buf)) in Sub EasyGet()), and
' we use the AsObject() function in VBVM6Lib.tlb to get back at it.
Public Function WriteFunction(ByVal rawBytes As Long, _
    ByVal sz As Long, ByVal nmemb As Long, _
    ByVal extra As Long) As Long
    
    Dim totalBytes As Long, i As Long
    Dim obj As Object, buf As Buffer
    
    totalBytes = sz * nmemb
    
    Set obj = AsObject(extra)
    Set buf = obj
      
    DoEvents
    ' append the binary characters to the HTML string
    For i = 0 To totalBytes - 1
        ' Append the write data
        buf.stringData = buf.stringData & Chr(MemByte(rawBytes + i)) 'Convert Byte to Char
    Next
       
    c = c + 1
    ReDim Preserve bytes(c)
    bytes(c - 1) = buf.stringData   'Get receive data, store in array
    'Debug.Print buf
    buf.stringData = ""
    
    ' Need this line below since AsObject gets a stolen reference
    ObjectPtr(obj) = 0&

    If stopcurl = True Then totalBytes = 0 'Return 0 if want to stop file transfer
    If c = 100 Then
        AppendFileByte         'Limit memo array up to 100
        c = 0
    End If
    
    ' Return value
    WriteFunction = totalBytes 'Return 0 if want to stop file transfer
End Function

' Again, rawBytes comes straight from libcurl and extra is a
' long, though we're not using it here.
Public Function DebugFunction(ByVal info As curl_infotype, _
    ByVal rawBytes As Long, ByVal numBytes As Long, _
    ByVal extra As Long) As Long
    Dim debugMsg As String
    Dim i As Long
    debugMsg = ""
    For i = 0 To numBytes - 1
        debugMsg = debugMsg & Chr(MemByte(rawBytes + i))
    Next
    Debug.Print "info=" & info & ", debugMsg=" & debugMsg
    DebugFunction = 0
End Function

'Check for file download progress here...
'dlTotal = download file total, dlNow = file download size up to this time
'ulTotal = upload file total, ulNow = total filesize uploaded to this time
Public Function ProgressFunction( _
   ByVal extraData As Object, _
   ByVal dlTotal As Double, _
   ByVal dlNow As Double, _
   ByVal ulTotal As Double, _
   ByVal ulNow As Double _
) As Integer

Dim fileleft As Double, download As Double, minus As Double, progress As Double

On Error Resume Next
vbcurl_easy_getinfo easy, CURLINFO_SPEED_DOWNLOAD, info
speed = Format(info / 1024, "##0.0")
m_kbps = speed
fileleft = CInt(dlTotal / 1024)
download = CInt(dlNow / 1024)
minus = fileleft - download
progress = (((CInt(filesize / 1024) + fileleft) - minus) / (CInt(filesize / 1024) + fileleft)) * 100
Debug.Print fileleft & "kb", download & "kb", minus & "kb", Format(progress, "##0") & "%", speed & " kbps"
m_progress = CInt(progress)
ProgressFunction = 0    'Return 0 if to continue progress update

End Function

'Create new file
Public Function AddFileByte()
Dim res As String

Open m_savefileas For Output As #1
For i = 0 To UBound(bytes) - 1
    Print #1, bytes(i);
Next i
Close #1

End Function

'Append byte to existed file
Public Function AppendFileByte()
Dim res As String

On Error GoTo jump
Open m_savefileas For Append As #1
For i = 0 To UBound(bytes) - 1
    Print #1, bytes(i);
Next i
Close #1
Exit Function
jump:
MsgBox "File already downloaded"

End Function

'Check if file exists
Public Function CheckFile(filepath As String) As Long
On Error GoTo jump
CheckFile = FileLen(filepath)
Exit Function
jump:
CheckFile = 0
End Function


