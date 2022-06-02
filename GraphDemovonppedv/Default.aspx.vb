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
        Dim tenantId = "d044ddddddddddddddddd"
        Dim ClientID = "ff6fddddddddddddddddddddddddddd"
        Dim authorityUri = $"https://login.microsoftonline.com/{tenantId}"
        Dim redirectUri = "http://localhost:44355"
        Dim scopes = {"https://graph.microsoft.com/.default"}
        Dim publicClient = PublicClientApplicationBuilder.Create(ClientID).WithAuthority(New Uri(authorityUri)).WithRedirectUri(redirectUri).Build()
        Dim sString = New System.Security.SecureString()
        For Each ch As Char In "dddddddddddddddd"
            sString.AppendChar(ch)
        Next
        Try
            Dim accounts = Await publicClient.GetAccountsAsync()
            daToken = publicClient.AcquireTokenSilent(scopes, accounts.First) _
                .WithForceRefresh(True) _
                .ExecuteAsync().Result.AccessToken
        Catch ex As Exception
            Dim accessTokenRequest = publicClient.AcquireTokenByUsernamePassword(scopes, "xxxxxx@ppedv.onmicrosoft.com", sString)
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

        msgresponse = Await graphClient.Teams("145xxxxxxxxxxxxxxxxx") _
                .Channels("19:bxxxxxxxxxxxxxxxxxxxx@thread.skype").Messages _
                .Request() _
                .AddAsync(ChatMessage)
      





    End Sub
End Class
