sUrl = "https://www.pushsafer.com/api"

'Add more optionally pushsafer parameter
'https://www.pushsafer.com/pushapi

sRequest = "k=YourKey&d=DeviceID&t=Title Here&m=Test Message send with VBScript&i=20&s=37&v=3"

HTTPPost sUrl, sRequest

Function HTTPPost(sUrl, sRequest)
  set oHTTP = CreateObject("Microsoft.XMLHTTP")
  oHTTP.open "POST", sUrl,false
  oHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  oHTTP.setRequestHeader "Content-Length", Len(sRequest)
  oHTTP.send sRequest
  HTTPPost = oHTTP.responseText
 End Function
