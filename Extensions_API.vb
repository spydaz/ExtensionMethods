Imports System.IO
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
''' <summary>
''' Manage API CALLS
''' </summary>
<ComClass(ClassAPI_.ClassId, ClassAPI_.InterfaceId, ClassAPI_.EventsId)>
Public Class ClassAPI_
    Public Const ClassId As String = "2828BA90-710A-401C-7AB3-38FF97BC1AC9"
    Public Const EventsId As String = "CDB7A307-F15A-401A-AC6D-3CF8786FD6F1"
    Public Const InterfaceId As String = "8BA745F8-511A-4059-829B-B531310144B5"

    Public Event RESPONSE_RECIEVED(ByRef Sender As ClassAPI_, ByRef Response As String)
    Public Event Error_RECIEVED(ByRef Sender As ClassAPI_, ByRef ErrorText As String)

    ''' <summary>
    ''' API URL
    ''' </summary>
    Public txtApiUrl As String = ""
    Private txtUserName As String = ""
    Private txtApiKey As String = ""


    Private txtEndPoint As String = ""
    ''' <summary>
    ''' Response Data
    ''' </summary>
    Private txtResponse As String = ""
    Public ReadOnly Property Response As String
        Get
            Return txtResponse
        End Get
    End Property
    ''' <summary>
    ''' Post Data
    ''' </summary>
    Private txtPostVars As String = ""
    Public ReadOnly Property PostedData As String
        Get
            Return txtPostVars
        End Get
    End Property
    Public Mode As String = "JSON"
    ''' <summary>
    ''' Return XML / JSON
    ''' </summary>
    ''' <returns></returns>
    Public Function GetMode() As String
        Return Mode
    End Function
    '+--------------------------------------------------------------------------+
    '|																			|
    '|	ValidateServerCertificate   										    |
    '|	=========================   											|
    '|																			|
    '|	Inputs:		Object sender  												|
    '|              X509Certificate certificate                                 |
    '|              X509Chain chain                                             |
    '|              SslPolicyErrors sslPolicyErrors                             |
    '|																			|
    '|	Returns:	bool														|
    '|																			|
    '|	Notes:		Override to allow any host certificate.                     | 
    '+--------------------------------------------------------------------------+	
    Public Shared Function ValidateServerCertificate(ByVal sender As Object, ByVal certificate As X509Certificate, ByVal chain As X509Chain, ByVal sslPolicyErrors As SslPolicyErrors) As Boolean
        Return True
    End Function
    '
    '       +----------------------------------------------------------------------------+
    '       |																			|
    '       |	httpRequest                                							    |
    '       |	===========			                        							|
    '       |																			|
    '       |	Inputs:	    string verb 					            				|
    '       |																			|
    '       |	Returns:	VOID            											|
    '       |																			|
    '       |	Notes:	   Accept 'put', 'post', 'get', or delete from the caller and   |
    '       |               perform the appropriate HTTP communications with the SLAPI   |
    '       |               server.                                                      |
    '       +----------------------------------------------------------------------------+															
    '       
    Private Sub httpRequest(ByVal verb As String)
        Try
            Dim url As String = txtApiUrl + "/" + txtEndPoint
            Dim req As HttpWebRequest = TryCast(WebRequest.Create(New Uri(url)), HttpWebRequest)
            Dim authPair As [String] = txtUserName + ":" + txtApiKey
            authPair = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(authPair))
            ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateServerCertificate)
            req.Headers.Add("Authorization", "Basic " & authPair)
            req.Method = verb
            If GetMode().Equals("XML") Then
                req.ContentType = "text/json"
            Else
                req.ContentType = "text/xml"
            End If
            If (verb.Equals("post")) OrElse (verb.Equals("put")) Then
                Dim content As Byte() = UTF8Encoding.UTF8.GetBytes(txtPostVars.Trim())
                req.ContentLength = content.Length
                Using post As Stream = req.GetRequestStream()
                    post.Write(content, 0, content.Length)
                End Using
            End If
            Dim result As String = Nothing
            Using resp As HttpWebResponse = TryCast(req.GetResponse(), HttpWebResponse)
                Dim reader As New StreamReader(resp.GetResponseStream())
                result = reader.ReadToEnd()
            End Using
            If GetMode().Equals("XML") Then
                txtResponse = FormatXml(result)
                RaiseEvent RESPONSE_RECIEVED(Me, txtResponse)
            Else
                txtResponse = FormatJSON(result)
                RaiseEvent RESPONSE_RECIEVED(Me, txtResponse)
            End If
        Catch [error] As WebException

            Dim extMsg As String = ""

            If [error].Response IsNot Nothing Then

                If [error].Response.ContentLength <> 0 Then
                    Using stream = [error].Response.GetResponseStream()
                        Using reader = New StreamReader(stream)
                            extMsg = reader.ReadToEnd()
                        End Using
                    End Using
                End If

            End If
            System.Windows.Forms.MessageBox.Show([error].Message.ToString() + vbCrLf + extMsg)
            RaiseEvent Error_RECIEVED(Me, [error].Message.ToString() + vbCrLf + extMsg)

        End Try
    End Sub

    '+--------------------------------------------------------------------------+
    '|																			|
    '|	FormatXml      			                							    |
    '|	=========                 	    										|
    '|																			|
    '|	Inputs:		string unformattedXml										|
    '|																			|
    '|	Returns:	string formattedXml											|
    '|																			|
    '|	Notes:		Use this method to perform some basic housekeeping on       |
    '|              XML string before plunking them into a text box.  It will   |
    '|              make the output easier on the eyes.                         |
    '+--------------------------------------------------------------------------+
    Private Function FormatXml(ByVal unformattedXml As String) As String
        Dim xd As New XmlDocument()
        Try
            xd.LoadXml(unformattedXml)
        Catch [error] As Exception
            System.Windows.Forms.MessageBox.Show([error].Message.ToString())
            Return unformattedXml
        End Try
        Dim sb As New StringBuilder()
        Dim sw As New StringWriter(sb)
        Dim xtw As XmlTextWriter = Nothing
        Try
            xtw = New XmlTextWriter(sw)
            xtw.Formatting = Formatting.Indented
            xd.WriteTo(xtw)
        Finally
            If xtw IsNot Nothing Then
                xtw.Close()
            End If
        End Try
        Return sb.ToString()
    End Function
    '       +---------------------------------------------------------------------------+
    '       |																     	    |
    '       |	FormatJson      			                							|
    '       |	=========                 	    										|
    '       |																			|
    '       |	Inputs:		string unformattedJson										|
    '       |																			|
    '       |	Returns:	string formattedJson										|
    '       |																			|
    '       |	Notes:	   Use this method to perform some basic housekeeping on        |
    '       |              JSON string before plunking them into a text box.  It will   |
    '       |              make the output easier on the eyes.                          |
    '       +---------------------------------------------------------------------------+															
    Private Function FormatJSON(ByVal unformattedJSON As String) As String
        Dim sb As New StringBuilder()
        Dim chars As Char() = unformattedJSON.ToCharArray()
        Dim len As Integer = chars.Length
        Dim indent As Integer = 0
        Dim last_char As Char = " "c
        Dim new_line As Boolean = True
        For i As Integer = 0 To len - 1
            If chars(i) = "}"c Then
                new_line = True
                sb.AppendLine()
            ElseIf chars(i) = "{"c Then
                If i > 0 Then
                    sb.AppendLine()
                    For j As Integer = 0 To indent - 1
                        sb.Append(" "c)
                    Next
                End If
            End If
            If last_char = "}"c Then
                indent -= 3
            End If
            If last_char = "{"c Then
                new_line = True
                sb.AppendLine()
                indent += 3
            End If
            If last_char = ","c Then
                sb.AppendLine()
                new_line = True
            End If
            If new_line Then
                For j As Integer = 0 To indent - 1
                    sb.Append(" "c)
                Next
            End If
            sb.Append(chars(i))
            new_line = False
            last_char = chars(i)
        Next
        Return sb.ToString()
    End Function
    ''' <summary>
    ''' After data has been set (SetCredentials) (SetUrl) (SetPostData)
    ''' Call can by Simple Verb
    ''' </summary>
    ''' <param name="Verb"></param>
    Public Sub MakeCall(ByRef Verb As String)
        Select Case Verb
            Case "put"
                httpRequest("put")
            Case "get"
                httpRequest("get")
            Case "post"
                httpRequest("post")
            Case "delete"
                httpRequest("delete")
        End Select
    End Sub
    Public Sub SetCredentials(ByRef Username As String, ByRef ApiKey As String)
        txtUserName = Username
        txtApiKey = ApiKey
    End Sub
    Public Sub SetURL(ByRef Url As String, ByRef Endpoint As String)
        txtApiUrl = Url
        txtPostVars = Endpoint
    End Sub
    Public Sub SetPostData(ByRef PostData As String)
        txtPostVars = PostData
    End Sub
End Class

