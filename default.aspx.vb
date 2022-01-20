Imports Microsoft.Owin.Security
Imports Microsoft.Owin.Security.Cookies
Imports Microsoft.Owin.Security.OpenIdConnect
Imports Newtonsoft.Json

Public Class _default
    Inherits System.Web.UI.Page

    Private addAdminScrip As Boolean = False

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Request.Form.AllKeys.Contains("ActionId") Then

            Dim lStrCallbackUrl As String = String.Format("{0}://{1}/default.aspx", Request.Url.Scheme, Request.Url.Authority)

            If Request.Form("ActionId") = "342f02ac-8362-4c77-a83c-11ba761fd23b" Then

                'LOGIN
                Context.GetOwinContext().Authentication.Challenge(
                    New AuthenticationProperties With {.RedirectUri = lStrCallbackUrl},
                    OpenIdConnectAuthenticationDefaults.AuthenticationType)

            ElseIf Request.Form("ActionId") = "401e81e9-6792-43f8-b755-8bf0cd40379b" Then

                'LOGOUT
                '   Context.GetOwinContext().Authentication.SignOut(
                '      New AuthenticationProperties With {.RedirectUri = lStrCallbackUrl},
                ' OpenIdConnectAuthenticationDefaults.AuthenticationType, CookieAuthenticationDefaults.AuthenticationType)

                Dim n As Microsoft.Owin.OwinContext = Context.GetOwinContext()

                For Each d In n.Request.Cookies

                    n.Response.Cookies.Delete(d.Key)

                Next

                Session.Clear()
                Session.Abandon()

                Response.Redirect(Request.RawUrl)
            End If
        End If

        If User.arsIsAuthenticated() Then

            If Session("ArsUser") Is Nothing Then

                Session("ArsUser") =
                    New Ars.User(User.azureOnpremSid)
            Else

                'Session("ArsUser").SetHttpContext(Context)
            End If

            'CONFIGURE LOGOUT BUTTON
            ActionId.Value = "401e81e9-6792-43f8-b755-8bf0cd40379b"
            btnAction.Value = "Logout"

            'CHECK USER EXISTS IN DATABASE
            If Session("ArsUser").IsAllowed Then

                arsMain.Visible = True
                arsNav.Visible = True

                If Session("ArsUser").IsAdmin Then

                    addAdminScrip = True
                    arsNavAdminMenu.Visible = True
                End If
            Else

                arsMessage_AccessDenied.Visible = True
            End If
        Else

            Session.Clear()

            'CONFIGURE LOGIN BUTTON
            ActionId.Value = "342f02ac-8362-4c77-a83c-11ba761fd23b"
            btnAction.Value = "Login"

            If Request.QueryString("auth_error") IsNot Nothing Then
                'SHOW LOGIN ERROR
                arsMessage_AuthErrorMessage.InnerText = Request.QueryString("auth_error")
                arsMessage_AuthError.Visible = True
            Else
                'SHOW PLEASE LOGIN
                arsMessage_Login.Visible = True
            End If
        End If

        If User.onpremAuthenticated() = True Then

            arsAdAuth.Visible = True
        End If

        If User.azureAuthenticated() = True Then

            arsAzAuth.Visible = True

            btnAzAuth.Attributes("data-AzExpTs") =
            User.azureAuthenticationExpireTimestamp()
        End If


        'ADD VERSION TO CSS AND JS FILES
        For Each lObjControl In Page.Header.Controls

            If lObjControl.GetType() = GetType(System.Web.UI.HtmlControls.HtmlLink) Then

                lObjControl.href = Replace(lObjControl.href, ".css", ".css?v=" & Application("appVer"))
            End If

            If lObjControl.GetType() = GetType(System.Web.UI.LiteralControl) Or lObjControl.GetType() = GetType(System.Web.UI.WebControls.Literal) Then

                lObjControl.text = Replace(lObjControl.text, ".js", ".js?=" & Application("appVer"))
            End If
        Next

    End Sub

    Private Sub _default_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender

        If addAdminScrip Then

            If Page.Header IsNot Nothing Then

                Dim lObjAdminScript As New HtmlGenericControl("script")
                lObjAdminScript.Attributes.Add("src", "/inc/js/arsAdmin.js")

                Page.Header.Controls.AddAt(6, lObjAdminScript)
            End If
        End If
    End Sub
End Class

