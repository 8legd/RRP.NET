﻿Option Explicit On

Public Class Batch

    Private url As System.Uri
    Private clientTimeout As Integer
    Private requestTimeout As Integer
    Private textEncoding As System.Text.Encoding

    Private batch As New ArrayList

    Public Sub New(ByVal clientTimeout As Integer, _
                   ByVal requestTimeout As Integer, _
                   ByVal textEncoding As System.Text.Encoding, _
                   Optional ByVal rrp As String = "http://127.0.0.1:8000")
        Dim url As String = rrp
        If url.EndsWith("/") Then
            url = url.Substring(0, url.Length - 1)
        End If
        url = url & "/batch/multipartmixed"
        Me.url = New System.Uri(url)
        Me.clientTimeout = clientTimeout
        Me.requestTimeout = requestTimeout
        Me.textEncoding = textEncoding
    End Sub

    'Inner class for managing an individual request/response pair in the batch
    Protected Class requestResponse

        Public RequestUrl As System.Uri
        Public RequestMethod As String
        Public RequestContentType As String
        Public RequestBodyText As String
        Public RequestAdditionalHeaders As System.Collections.Specialized.NameValueCollection


        Public ResponseStatusOK As Boolean
        Public ResponseStatusLine As String
        Public ResponseHeaders As System.Net.WebHeaderCollection
        Public ResponseBodyText As String


        Sub New(ByVal requestUrl As System.Uri, _
                ByVal requestMethod As String, _
                ByVal requestContentType As String, _
                ByVal requestBodyText As String, _
                Optional ByVal requestAdditionalHeaders As System.Collections.Specialized.NameValueCollection = Nothing)

            Me.RequestUrl = requestUrl
            Me.RequestMethod = requestMethod
            Me.RequestContentType = requestContentType
            Me.RequestBodyText = requestBodyText
            Me.RequestAdditionalHeaders = requestAdditionalHeaders

        End Sub

        Sub New(ByVal httpApplicationContent As String)

            'Set some defaults
            Me.ResponseStatusOK = False
            Me.ResponseStatusLine = ""
            Me.ResponseHeaders = New System.Net.WebHeaderCollection()
            Me.ResponseBodyText = ""

            'Try and build the response
            Dim splitOnHTTPParts As String() = httpApplicationContent.Split(New String() {vbCrLf & vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim splitOnCRLF As String()
            Dim splitOnSpace As String()
            Dim splitOnColon As String()
            Dim length = UBound(splitOnHTTPParts)
            If length > 0 Then
                splitOnCRLF = splitOnHTTPParts(1).Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                For Each line As String In splitOnCRLF
                    If Me.ResponseStatusLine = "" Then 'First line is the status line
                        Me.ResponseStatusLine = line
                        splitOnSpace = Me.ResponseStatusLine.Split(New String() {" "}, StringSplitOptions.RemoveEmptyEntries)
                        If UBound(splitOnSpace) > 1 Then
                            Me.ResponseStatusOK = (splitOnSpace(1) = "200")
                        End If
                    Else
                        'The rest are headers
                        splitOnColon = line.Split(New String() {": "}, StringSplitOptions.RemoveEmptyEntries)
                        If UBound(splitOnColon) > 0 Then
                            Me.ResponseHeaders.Add(splitOnColon(0), splitOnColon(1))
                        End If
                    End If
                Next
            End If
            If length > 1 Then
                Me.ResponseBodyText = splitOnHTTPParts(2)
            End If

        End Sub

        Function toApplicationHttp(ByVal textEncoding As System.Text.Encoding) As String
            '	Method<SPACE>Request-URI<SPACE>HTTP-Version<CRLF>
            '	Headers<CRLF>
            '	CRLF
            '	Body<CRLF>
            '	CRLF
            Dim result As String = "Content-Type: application/http" & vbCrLf & vbCrLf
            result = result & "POST " & Me.RequestUrl.PathAndQuery & " HTTP/1.1" & vbCrLf
            result = result & "Host: " & Me.RequestUrl.Host & vbCrLf
            result = result & "Content-Type: " & Me.RequestContentType & vbCrLf
            If Not Me.RequestAdditionalHeaders Is Nothing Then
                For Each key In Me.RequestAdditionalHeaders.AllKeys
                    result = result & key & ": " & Me.RequestAdditionalHeaders(key) & vbCrLf
                Next key
            End If
            result = result & "Content-Length: " & textEncoding.GetBytes(Me.RequestBodyText).Length.ToString() & vbCrLf
            result = result & "Forwarded: proto=" & Me.RequestUrl.Scheme & vbCrLf
            result = result & vbCrLf & Me.RequestBodyText & vbCrLf & vbCrLf
            Return result
        End Function

    End Class

    'Returns an index representing the position in the batch for this request
    Public Function AddRequest(ByVal requestUrl As System.Uri, _
                                ByVal requestMethod As String, _
                                ByVal requestContentType As String, _
                                ByVal requestBodyText As String, _
                                Optional ByVal requestAdditionalHeaders As System.Collections.Specialized.NameValueCollection = Nothing) As Integer
        Me.batch.Add(New requestResponse(requestUrl, requestMethod, requestContentType, requestBodyText, requestAdditionalHeaders))
        Return Me.batch.Count - 1
    End Function


    Public Sub Process()

        'Build a multipart boundary
        Dim multipartBoundary As String = System.Guid.NewGuid.ToString
        While Len(multipartBoundary) < 36
            multipartBoundary = "-" & multipartBoundary
        End While
        multipartBoundary = "----" & multipartBoundary

        'Build the batch request
        Dim batchRequestResponse As requestResponse
        Dim batchRequestContent As String = "--" & multipartBoundary
        For Each rr As requestResponse In batch
            batchRequestContent = batchRequestContent & vbCrLf & rr.toApplicationHttp(Me.textEncoding) & "--" & multipartBoundary
        Next
        batchRequestContent = batchRequestContent & "--" & vbCrLf

        Dim additionalHeaders As New System.Collections.Specialized.NameValueCollection
        additionalHeaders.Add("x-rrp-timeout", Me.requestTimeout)
        batchRequestResponse = New requestResponse(Me.url, "POST", "multipart/mixed; boundary=" & multipartBoundary, batchRequestContent, additionalHeaders)

        Me.transport(Me.clientTimeout, batchRequestResponse)

        'Collect the individual responses
        Dim rawResponses As New ArrayList
        If batchRequestResponse.ResponseStatusOK Then
            Dim mimeContentType As System.Net.Mime.ContentType = New System.Net.Mime.ContentType(batchRequestResponse.ResponseHeaders.Get("Content-Type"))
            If mimeContentType.MediaType = "multipart/mixed" AndAlso mimeContentType.Boundary <> "" Then
                Dim splitOnBoundary As String() = batchRequestResponse.ResponseBodyText.Split(New String() {"--" & mimeContentType.Boundary}, StringSplitOptions.RemoveEmptyEntries)
                For Each part As String In splitOnBoundary
                    If part.Trim() <> "" AndAlso part.Trim() <> "--" Then
                        rawResponses.Add(New requestResponse(part))
                    End If
                Next part
            Else
                Throw New Exception("Invalid Content-Type in response, expected `multipart/mixed` received " & batchRequestResponse.ResponseHeaders.Get("Content-Type"))
            End If
        Else
            Throw New Exception("Error response received " & batchRequestResponse.ResponseStatusLine)
        End If

        If rawResponses.Count <> batch.Count Then
            Throw New Exception("Error received " & rawResponses.Count & " responses but expected " & batch.Count)
        End If

        'Update the original batch response
        For i As Integer = 0 To (batch.Count - 1)
            batch(i).ResponseStatusOK = rawResponses(i).ResponseStatusOK
            batch(i).ResponseStatusLine = rawResponses(i).ResponseStatusLine
            batch(i).ResponseHeaders = rawResponses(i).ResponseHeaders
            batch(i).ResponseBodyText = rawResponses(i).ResponseBodyText
        Next


    End Sub

    'Useful for benchmarking
    Public Sub ProcessWithoutRRP()
        For Each rr As requestResponse In batch
            Me.transport(Me.requestTimeout, rr)
        Next
    End Sub


    Public Sub ReadResponse(ByVal requestIndex As Integer, _
                             ByRef responseStatusOK As Boolean, _
                             ByRef responseStatusLine As String, _
                             ByRef responseHeaders As System.Net.WebHeaderCollection, _
                             ByRef responseBodyText As String)
        Dim rr = batch.Item(requestIndex)
        responseStatusOK = rr.ResponseStatusOK
        responseStatusLine = rr.ResponseStatusLine
        responseHeaders = rr.ResponseHeaders
        responseBodyText = rr.ResponseBodyText
    End Sub


    Private Sub transport(ByVal timeout As Integer, _
                          ByRef reqres As requestResponse)

        Dim req As System.Net.HttpWebRequest = DirectCast(System.Net.WebRequest.CreateDefault(reqres.RequestUrl), System.Net.HttpWebRequest)
        req.Method = reqres.RequestMethod
        req.Timeout = timeout * 1000 'timeout is specified in seconds
        req.ContentType = reqres.RequestContentType

        If Not reqres.RequestAdditionalHeaders Is Nothing Then
            For Each key In reqres.RequestAdditionalHeaders.AllKeys
                req.Headers.Add(key, reqres.RequestAdditionalHeaders(key))
            Next key
        End If

        Dim reqStream As New System.IO.StreamWriter(req.GetRequestStream(), Me.textEncoding)
        reqStream.Write(reqres.RequestBodyText)
        reqStream.Flush()
        reqStream.Close()

        Dim res = CType(req.GetResponse(), System.Net.HttpWebResponse)
        reqres.ResponseStatusOK = (res.StatusCode = System.Net.HttpStatusCode.OK)
        reqres.ResponseStatusLine = "HTTP/" & res.ProtocolVersion.ToString() & " " & res.StatusCode & " " & res.StatusDescription
        reqres.ResponseHeaders = res.Headers

        Dim resStream As System.IO.Stream = res.GetResponseStream()
        Dim reader As New System.IO.StreamReader(resStream)

        reqres.ResponseBodyText = reader.ReadToEnd()

        reader.Close()
        res.Close()

    End Sub

End Class