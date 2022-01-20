Imports Owin
Imports Microsoft.Owin
Imports Microsoft.Owin.Security
Imports Microsoft.Owin.Security.Cookies
Imports Microsoft.Owin.Security.Notifications
Imports Microsoft.Owin.Security.OpenIdConnect
Imports System.Threading.Tasks
Imports Microsoft.IdentityModel.Protocols.OpenIdConnect
Imports Microsoft.IdentityModel.Tokens
Imports Microsoft.Owin.Extensions
Imports System.Net

Partial Public Class Startup

    Public Sub ConfigureAuth(app As IAppBuilder)

        app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType)

        app.UseCookieAuthentication(New CookieAuthenticationOptions() With {
            .SlidingExpiration = True,
            .ExpireTimeSpan = TimeSpan.FromMinutes(60)
        })

        app.UseOpenIdConnectAuthentication(
            New Microsoft.Owin.Security.OpenIdConnect.OpenIdConnectAuthenticationOptions() With {
                .ClientId = ConfigurationManager.AppSettings.Item("oId_clientId"),
                .Authority = ConfigurationManager.AppSettings.Item("oId_Authority"),
                .RedirectUri = ConfigurationManager.AppSettings.Item("oId_RedirectUri"),
                .PostLogoutRedirectUri = ConfigurationManager.AppSettings.Item("oId_PostLogoutRedirectUri"),
                .Scope = OpenIdConnectScope.OpenIdProfile,
                .ResponseType = OpenIdConnectResponseType.CodeIdToken,
                .ClientSecret = ConfigurationManager.AppSettings.Item("oId_ClientSecret"),
                .Notifications =
                    New Microsoft.Owin.Security.OpenIdConnect.OpenIdConnectAuthenticationNotifications() With {
                        .RedirectToIdentityProvider = AddressOf OnRedirectToIdentityProvider,
                        .MessageReceived = AddressOf OnMessageReceived,
                        .SecurityTokenReceived = AddressOf OnSecurityTokenReceived,
                        .SecurityTokenValidated = AddressOf OnSecurityTokenValidated,
                        .AuthorizationCodeReceived = AddressOf OnAuthorizationCodeReceived,
                        .AuthenticationFailed = AddressOf OnAuthenticationFailed
                    },
                .TokenValidationParameters = New TokenValidationParameters() With {
                    .ValidateIssuer = False,
                    .ValidateAudience = True,
                    .ValidateIssuerSigningKey = True,
                    .ValidateLifetime = True,
                    .ValidateTokenReplay = True
                }
        })
    End Sub

    Private Function OnRedirectToIdentityProvider(arg As RedirectToIdentityProviderNotification(Of OpenIdConnectMessage, OpenIdConnectAuthenticationOptions)) As Task
        Return Task.FromResult(0)
    End Function

    Private Function OnMessageReceived(arg As MessageReceivedNotification(Of OpenIdConnectMessage, OpenIdConnectAuthenticationOptions)) As Task
        Return Task.FromResult(0)
    End Function

    Private Function OnAuthorizationCodeReceived(arg As AuthorizationCodeReceivedNotification) As Task
        Return Task.FromResult(0)
    End Function

    Private Function OnAuthenticationFailed(arg As AuthenticationFailedNotification(Of OpenIdConnectMessage, OpenIdConnectAuthenticationOptions)) As Task

        Dim ml As New Ars.MethodLogger("Startup.Auth")
        ml.write(Ars.LOGLEVEL.INFO,
                 String.Format("Error in OpenId authentication: {0}",
                               arg.Exception.Message))
        ml.done()

        arg.HandleResponse()
        arg.Response.Redirect("/?auth_error=" + arg.Exception.Message)

        Return Task.FromResult(0)
    End Function

    Private Function OnSecurityTokenReceived(arg As SecurityTokenReceivedNotification(Of OpenIdConnectMessage, OpenIdConnectAuthenticationOptions)) As Task
        Return Task.FromResult(0)
    End Function

    Private Function OnSecurityTokenValidated(arg As SecurityTokenValidatedNotification(Of OpenIdConnectMessage, OpenIdConnectAuthenticationOptions)) As Task

        Dim ml As New Ars.MethodLogger("Startup.Auth")
        ml.write(Ars.LOGLEVEL.INFO,
                 String.Format("Successfull OpenId authentication: {0}",
                               arg.AuthenticationTicket.Identity.Name))
        ml.done()

        Return Task.FromResult(0)
    End Function
End Class
