'
' POST a block of data to the server with it's hash
' Server should return 'OK' if successful
'
Function PostData(ByRef BlockData() As Byte) As String
    
    Dim sBoundary As String * 24
    sBoundary = RandomString(24)
    
    Dim BlockHash As String * 32
    Dim Base64EncodedBlock As String
    Base64EncodedBlock = EncodeBase64(BlockData)
    
    Dim PostBody As String
    
    Dim http As Object
    Set http = CreateObject("winhttp.winhttprequest.5.1")
    http.Open "POST", URL, False
    http.SetRequestHeader "Content-Type", "multipart/form-data" & ";boundary=" & sBoundary
    
    ' build the multipart/form-data... by hand :~(
    PostBody = PostBody & "--" & sBoundary & vbCrLf
    
    PostBody = PostBody & "Content-Disposition: form-data; name=" & Chr(34) & vbCrLf
    PostBody = PostBody & vbCrLf
    PostBody = PostBody & "--" & sBoundary & vbCrLf
    
    PostBody = PostBody & "Content-Type: application/octet-stream" & vbCrLf
    PostBody = PostBody & "Content-Disposition: form-data; name=""file""; filename=" & Chr(34) & "upload.bin" & Chr(34) & vbCrLf
    PostBody = PostBody & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
    PostBody = PostBody & Base64EncodedBlock & vbCrLf & vbCrLf
    PostBody = PostBody & "--" & sBoundary & "--"
   
    http.Send PostBody

End Function

'
' Random string generator for multipart boundaries
'
Function RandomString(length As Integer) As String

    Randomize
    Dim chars As String
    Dim result As String
    result = ""
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    Dim i As Long
    For i = 1 To length
        result = result & Mid$(chars, Int(Rnd() * 62) + 1, 1)
    Next
    RandomString = result

End Function

Private Function EncodeBase64(bytes) As String
  Dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set bytes, get encoded String
  EL.NodeTypedValue = bytes
  EncodeBase64 = EL.Text
End Function
