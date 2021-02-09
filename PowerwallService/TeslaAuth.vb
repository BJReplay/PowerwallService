Imports System.Net
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web
Imports Newtonsoft.Json.Linq

Namespace TeslaAuth
    Public Class Tokens
        Public Property AccessToken As String
        Public Property RefreshToken As String
    End Class
    Friend Class LoginInfo
        Public Property CodeVerifier As String
        Public Property CodeChallenge As String
        Public Property State As String
        Public Property Cookie As String
        Public Property FormFields As Dictionary(Of String, String)
    End Class
    Module TeslaAuthHelper
        Private Const TESLA_CLIENT_ID As String = "81527cff06843c8634fdc09e8ac0abefb46ac849f38fe1e431c2ef2106796384"
        Private Const TESLA_CLIENT_SECRET As String = "c7257eb71a564034f9419ee651c7d0e5f7aa6bfbd18bafb5c5c033b093bb2fa3"
        Private random As Random = New Random()

        Function RandomString(ByVal length As Integer) As String
            Const chars As String = "abcdefghijklmnopqrstuvwxyz0123456789"
            Return New String(Enumerable.Repeat(chars, length).[Select](Function(s) s(random.[Next](s.Length))).ToArray())
        End Function

        Function ComputeSHA256Hash(ByVal text As String) As String
            Dim hashString As String

            Using sha256 = SHA256Managed.Create()
                Dim hash = sha256.ComputeHash(Encoding.[Default].GetBytes(text))
                hashString = ToHex(hash, False)
            End Using

            Return hashString
        End Function

        Private Function ToHex(ByVal bytes As Byte(), ByVal upperCase As Boolean) As String
            Dim result As StringBuilder = New StringBuilder(bytes.Length * 2)

            For i As Integer = 0 To bytes.Length - 1
                result.Append(bytes(i).ToString(If(upperCase, "X2", "x2")))
            Next

            Return result.ToString()
        End Function

        Function Authenticate(ByVal username As String, ByVal password As String, ByVal Optional mfaCode As String = Nothing) As Tokens
            Dim loginInfo = InitializeLogin()
            Dim code = GetAuthorizationCode(username, password, mfaCode, loginInfo)
            Dim tokens = ExchangeCodeForBearerToken(code, loginInfo)
            Dim accessToken = ExchangeAccessTokenForBearerToken(tokens.AccessToken)
            Return New Tokens With {
                .AccessToken = accessToken,
                .RefreshToken = tokens.RefreshToken
            }
        End Function

        Private Function InitializeLogin() As LoginInfo
            Dim result = New LoginInfo()
            result.CodeVerifier = RandomString(86)
            Dim code_challenge_SHA256 = ComputeSHA256Hash(result.CodeVerifier)
            result.CodeChallenge = Convert.ToBase64String(Encoding.[Default].GetBytes(code_challenge_SHA256))
            result.State = RandomString(20)

            Using client As HttpClient = New HttpClient()
                Dim values As Dictionary(Of String, String) = New Dictionary(Of String, String) From {
                    {"client_id", "ownerapi"},
                    {"code_challenge", result.CodeChallenge},
                    {"code_challenge_method", "S256"},
                    {"redirect_uri", "https://auth.tesla.com/void/callback"},
                    {"response_type", "code"},
                    {"scope", "openid email offline_access"},
                    {"state", result.State}
                }
                Dim b As UriBuilder = New UriBuilder("https://auth.tesla.com/oauth2/v3/authorize")
                b.Port = -1
                Dim q = HttpUtility.ParseQueryString(b.Query)

                For Each v In values
                    q(v.Key) = v.Value
                Next

                b.Query = q.ToString()
                Dim url As String = b.ToString()
                Dim response As HttpResponseMessage = client.GetAsync(url).Result
                Dim resultContent = response.Content.ReadAsStringAsync().Result
                Dim hiddenFields = Regex.Matches(resultContent, "type=\""hidden\"" name=\""(.*?)\"" value=\""(.*?)\""")
                Dim formFields = New Dictionary(Of String, String)()

                For Each match As Match In hiddenFields
                    formFields.Add(match.Groups(1).Value, match.Groups(2).Value)
                Next

                Dim cookies As IEnumerable(Of String) = response.Headers.SingleOrDefault(Function(header) header.Key.ToLowerInvariant() = "set-cookie").Value
                Dim cookie = cookies.ToList()(0)
                cookie = cookie.Substring(0, cookie.IndexOf(" "))
                cookie = cookie.Trim()
                result.Cookie = cookie
                result.FormFields = formFields
                Return result
            End Using
        End Function

        Private Function GetAuthorizationCode(ByVal username As String, ByVal password As String, ByVal mfaCode As String, ByVal loginInfo As LoginInfo) As String
            Dim formFields = loginInfo.FormFields
            formFields.Add("identity", username)
            formFields.Add("credential", password)
            Dim code As String = ""

            Using ch As HttpClientHandler = New HttpClientHandler()
                ch.AllowAutoRedirect = False
                ch.UseCookies = False

                Using client As HttpClient = New HttpClient(ch)
                    client.BaseAddress = New Uri("https://auth.tesla.com")
                    client.DefaultRequestHeaders.Add("Cookie", loginInfo.Cookie)
                    Dim start As DateTime = DateTime.UtcNow

                    Using content As FormUrlEncodedContent = New FormUrlEncodedContent(formFields)
                        Dim b As UriBuilder = New UriBuilder("https://auth.tesla.com/oauth2/v3/authorize")
                        b.Port = -1
                        Dim q = HttpUtility.ParseQueryString(b.Query)
                        q("client_id") = "ownerapi"
                        q("code_challenge") = loginInfo.CodeChallenge
                        q("code_challenge_method") = "S256"
                        q("redirect_uri") = "https://auth.tesla.com/void/callback"
                        q("response_type") = "code"
                        q("scope") = "openid email offline_access"
                        q("state") = loginInfo.State
                        b.Query = q.ToString()
                        Dim url As String = b.ToString()
                        Dim result As HttpResponseMessage = client.PostAsync(url, content).Result
                        Dim resultContent As String = result.Content.ReadAsStringAsync().Result

                        If Not (result.IsSuccessStatusCode Or result.StatusCode = HttpStatusCode.Redirect) Then
                            Throw New Exception(If(String.IsNullOrEmpty(result.ReasonPhrase), result.StatusCode.ToString(), result.ReasonPhrase))
                        End If

                        Dim location As Uri = result.Headers.Location

                        If result.StatusCode <> HttpStatusCode.Redirect Then

                            If result.StatusCode = HttpStatusCode.OK AndAlso resultContent.Contains("passcode") Then

                                If String.IsNullOrEmpty(mfaCode) Then
                                    Throw New Exception("Multi-factor code required to authenticate")
                                End If

                                Return GetAuthorizationCodeWithMfa(mfaCode, loginInfo)
                            Else
                                Throw New Exception("Expected redirect did not occur")
                            End If
                        End If

                        If location Is Nothing Then
                            Throw New Exception("Redirect locaiton not available")
                        End If

                        code = HttpUtility.ParseQueryString(location.Query).[Get]("code")
                        Return code
                    End Using
                End Using
            End Using

            Throw New Exception("Authentication process failed")
        End Function

        Private Function ExchangeCodeForBearerToken(ByVal code As String, ByVal loginInfo As LoginInfo) As Tokens
            Dim body = New JObject()
            body.Add("grant_type", "authorization_code")
            body.Add("client_id", "ownerapi")
            body.Add("code", code)
            body.Add("code_verifier", loginInfo.CodeVerifier)
            body.Add("redirect_uri", "https://auth.tesla.com/void/callback")

            Using client As HttpClient = New HttpClient()
                client.BaseAddress = New Uri("https://auth.tesla.com")

                Using content = New StringContent(body.ToString(), System.Text.Encoding.UTF8, "application/json")
                    Dim result As HttpResponseMessage = client.PostAsync("https://auth.tesla.com/oauth2/v3/token", content).Result
                    Dim resultContent As String = result.Content.ReadAsStringAsync().Result
                    Dim response As JObject = JObject.Parse(resultContent)
                    Dim tokens = New Tokens() With {
                        .AccessToken = response("access_token").Value(Of String)(),
                        .RefreshToken = response("refresh_token").Value(Of String)()
                    }
                    Return tokens
                End Using
            End Using
        End Function

        Private Function ExchangeAccessTokenForBearerToken(ByVal accessToken As String) As String
            Dim body = New JObject()
            body.Add("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer")
            body.Add("client_id", TESLA_CLIENT_ID)
            body.Add("client_secret", TESLA_CLIENT_SECRET)

            Using client As HttpClient = New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(5)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & accessToken)

                Using content = New StringContent(body.ToString(), System.Text.Encoding.UTF8, "application/json")
                    Dim result As HttpResponseMessage = client.PostAsync("https://owner-api.teslamotors.com/oauth/token", content).Result
                    Dim resultContent As String = result.Content.ReadAsStringAsync().Result
                    Dim response As JObject = JObject.Parse(resultContent)
                    Return response("access_token").Value(Of String)()
                End Using
            End Using
        End Function

        Function RefreshToken(ByVal OldRefreshToken As String) As String
            Dim body = New JObject()
            body.Add("grant_type", "refresh_token")
            body.Add("client_id", "ownerapi")
            body.Add("refresh_token", OldRefreshToken)
            body.Add("scope", "openid email offline_access")

            Using client As HttpClient = New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(5)

                Using content = New StringContent(body.ToString(), System.Text.Encoding.UTF8, "application/json")
                    Dim result As HttpResponseMessage = client.PostAsync("https://auth.tesla.com/oauth2/v3/token", content).Result
                    Dim resultContent As String = result.Content.ReadAsStringAsync().Result
                    Dim response As JObject = JObject.Parse(resultContent)
                    Dim accessToken As String = response("access_token").Value(Of String)()
                    Return ExchangeAccessTokenForBearerToken(accessToken)
                End Using
            End Using
        End Function

        Private Function GetAuthorizationCodeWithMfa(ByVal mfaCode As String, ByVal loginInfo As LoginInfo) As String
            Dim mfaFactorId As String = GetMfaFactorId(loginInfo)
            VerifyMfaCode(mfaCode, loginInfo, mfaFactorId)
            Dim code = GetCodeAfterValidMfa(loginInfo)
            Return code
        End Function

        Private Function GetMfaFactorId(ByVal loginInfo As LoginInfo) As String
            Dim resultContent As String

            Using ch As HttpClientHandler = New HttpClientHandler()
                ch.UseCookies = False

                Using client As HttpClient = New HttpClient(ch)
                    client.DefaultRequestHeaders.Add("Cookie", loginInfo.Cookie)
                    Dim b As UriBuilder = New UriBuilder("https://auth.tesla.com/oauth2/v3/authorize/mfa/factors")
                    b.Port = -1
                    Dim q = HttpUtility.ParseQueryString(b.Query)
                    q.Add("transaction_id", loginInfo.FormFields("transaction_id"))
                    b.Query = q.ToString()
                    Dim url As String = b.ToString()
                    Dim result As HttpResponseMessage = client.GetAsync(url).Result
                    resultContent = result.Content.ReadAsStringAsync().Result
                    Dim response = JObject.Parse(resultContent)
                    Return response("data")(0)("id").Value(Of String)()
                End Using
            End Using
        End Function

        Private Sub VerifyMfaCode(ByVal mfaCode As String, ByVal loginInfo As LoginInfo, ByVal factorId As String)
            Using ch As HttpClientHandler = New HttpClientHandler()
                ch.AllowAutoRedirect = False
                ch.UseCookies = False

                Using client As HttpClient = New HttpClient(ch)
                    client.BaseAddress = New Uri("https://auth.tesla.com")
                    client.DefaultRequestHeaders.Add("Cookie", loginInfo.Cookie)
                    Dim body = New JObject()
                    body.Add("factor_id", factorId)
                    body.Add("passcode", mfaCode)
                    body.Add("transaction_id", loginInfo.FormFields("transaction_id"))

                    Using content = New StringContent(body.ToString(), System.Text.Encoding.UTF8, "application/json")
                        Dim result As HttpResponseMessage = client.PostAsync("https://auth.tesla.com/oauth2/v3/authorize/mfa/verify", content).Result
                        Dim resultContent As String = result.Content.ReadAsStringAsync().Result
                        Dim response = JObject.Parse(resultContent)
                        Dim valid As Boolean = response("data")("valid").Value(Of Boolean)()

                        If Not valid Then
                            Throw New Exception("MFA code invalid")
                        End If
                    End Using
                End Using
            End Using
        End Sub

        Private Function GetCodeAfterValidMfa(ByVal loginInfo As LoginInfo) As String
            Using ch As HttpClientHandler = New HttpClientHandler()
                ch.AllowAutoRedirect = False
                ch.UseCookies = False

                Using client As HttpClient = New HttpClient(ch)
                    client.BaseAddress = New Uri("https://auth.tesla.com")
                    client.DefaultRequestHeaders.Add("Cookie", loginInfo.Cookie)
                    Dim d As Dictionary(Of String, String) = New Dictionary(Of String, String)()
                    d.Add("transaction_id", loginInfo.FormFields("transaction_id"))

                    Using content As FormUrlEncodedContent = New FormUrlEncodedContent(d)
                        Dim b As UriBuilder = New UriBuilder("https://auth.tesla.com/oauth2/v3/authorize")
                        b.Port = -1
                        Dim q = HttpUtility.ParseQueryString(b.Query)
                        q.Add("client_id", "ownerapi")
                        q.Add("code_challenge", loginInfo.CodeChallenge)
                        q.Add("code_challenge_method", "S256")
                        q.Add("redirect_uri", "https://auth.tesla.com/void/callback")
                        q.Add("response_type", "code")
                        q.Add("scope", "openid email offline_access")
                        q.Add("state", loginInfo.State)
                        b.Query = q.ToString()
                        Dim url As String = b.ToString()
                        Dim temp = content.ReadAsStringAsync().Result
                        Dim result As HttpResponseMessage = client.PostAsync(url, content).Result
                        Dim resultContent As String = result.Content.ReadAsStringAsync().Result
                        Dim location As Uri = result.Headers.Location

                        If result.StatusCode = HttpStatusCode.Redirect AndAlso location IsNot Nothing Then
                            Return HttpUtility.ParseQueryString(location.Query).[Get]("code")
                        End If

                        Throw New Exception("Unable to get authorization code")
                    End Using
                End Using
            End Using
        End Function
    End Module
End Namespace
