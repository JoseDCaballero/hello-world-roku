sub init()
    m.top.functionName = "loadData"
    m.apiUrl = GetApiUrl()
end sub

sub loadData()
    request = CreateObject("roUrlTransfer")
    port = CreateObject("roMessagePort")
    request.SetMessagePort(port)
    request.SetCertificatesFile("common:/certs/ca-bundle.crt")
    request.InitClientCertificates()
    request.SetUrl(m.apiUrl)

    if not request.AsyncGetToString() then
        m.top.error = "NO SE PUDO INICIAR API"
        return
    end if

    msg = wait(10000, port)

    if msg = invalid then
        request.AsyncCancel()
        m.top.error = "API TIMEOUT"
        return
    end if

    statusCode = msg.GetResponseCode()
    response = msg.GetString()

    if response = invalid then
        m.top.error = "API SIN RESPUESTA: " + msg.GetFailureReason()
        return
    end if

    if response = "" then
        m.top.error = "API SIN RESPUESTA: " + msg.GetFailureReason()
        return
    end if

    if statusCode < 200 or statusCode >= 300 then
        m.top.error = "API HTTP " + statusCode.ToStr()
        return
    end if

    m.top.response = response
end sub
