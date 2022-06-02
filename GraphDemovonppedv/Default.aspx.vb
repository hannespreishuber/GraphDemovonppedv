Imports System.Net.Http
Imports System.Net.Http.Headers
Imports Microsoft.Graph
Imports Microsoft.Graph.Core

Imports Microsoft.Identity.Client

Public Class _Default
    Inherits Page


    Public Shared daToken As String
    Dim username As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    End Sub

    Protected Async Sub Button1_Click(sender As Object, e As EventArgs)
        Dim tenantId = "d044494e-fc77-4ae0-8c6b-8b4520666035"
        Dim ClientID = "ff6fc891-85b5-4998-a35a-6a7f0910744e"
        Dim authorityUri = $"https://login.microsoftonline.com/{tenantId}"
        Dim redirectUri = "http://localhost:44355"
        Dim scopes = {"https://graph.microsoft.com/.default"}
        Dim publicClient = PublicClientApplicationBuilder.Create(ClientID).WithAuthority(New Uri(authorityUri)).WithRedirectUri(redirectUri).Build()
        Dim sString = New System.Security.SecureString()
        For Each ch As Char In "WDTkJp5.cwJP"
            sString.AppendChar(ch)
        Next
        Try
            Dim accounts = Await publicClient.GetAccountsAsync()
            daToken = publicClient.AcquireTokenSilent(scopes, accounts.First) _
                .WithForceRefresh(True) _
                .ExecuteAsync().Result.AccessToken
        Catch ex As Exception
            Dim accessTokenRequest = publicClient.AcquireTokenByUsernamePassword(scopes, "WebChat@ppedv.onmicrosoft.com", sString)
            daToken = accessTokenRequest.ExecuteAsync().Result.AccessToken

        End Try






        Dim _httpClient = New HttpClient()
        _httpClient.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", daToken)
        _httpClient.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))


        Dim graphClient As New GraphServiceClient(_httpClient) With {
    .AuthenticationProvider = New DelegateAuthenticationProvider(Async Function(requestMessage)
                                                                     requestMessage.Headers.Authorization = New AuthenticationHeaderValue("Bearer", daToken)
                                                                 End Function)
}




        Dim ChatMessage = New ChatMessage With
                    {
                        .Body = New ItemBody With
                        {
                            .ContentType = BodyType.Html,
                            .Content = $"[Demo22222]"
                        }
                    }
        'office 365 license
        'answeere chat


        Dim msgresponse As ChatMessage

        msgresponse = Await graphClient.Teams("145db263-90e3-4cae-b401-0381b06ff2b5") _
                .Channels("19:bd7758cea194476fb3627c65ad0bf5bc@thread.skype").Messages _
                .Request() _
                .AddAsync(ChatMessage)
        'msgresponse.Id 1634907923055






    End Sub
End Class