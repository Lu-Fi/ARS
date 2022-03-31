Imports System.Collections.Concurrent
Imports System.Data.SqlClient
Imports System.DirectoryServices.Protocols
Imports System.Net
Imports System.Security.Principal
Imports System.Threading

Public Class ArsGlobal : Inherits System.Web.HttpApplication

    Private _state As Integer = 0
    Private _worker As Thread = Nothing

    Private _users As New List(Of Ars.AssignedUser)
    Private _groups As New List(Of Object)
    Private _assignments As New List(Of Ars.Assignment)

    Private _users_processed As New ConcurrentDictionary(Of String, Ars.AssignedUser)

    Private ApplicationSettings As Ars.ApplicationSettings

    Private SqlConnection As SqlConnection = Nothing
    Private SqlConnectionString As String =
        ConfigurationManager.AppSettings.Item("sqlConnectionString")

    Private ClassContext As HttpContext

    Private ldap As Ars.Ldap

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)

        ClassContext = Context

        ApplicationSettings =
            New Ars.ApplicationSettings

        Application("ApplicationSettings") = ApplicationSettings
        Application("appVer") =
            System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString().Replace(".", "")

        SqlConnection =
            New SqlConnection(ApplicationSettings.SqlConnectionString)

        ldap =
            New Ars.Ldap(Context)

        Application.Lock()

        Ars.log.LogLevel = ApplicationSettings.logLevel

        Dim ml As New Ars.MethodLogger("Global.asax")
        ml.write(Ars.LOGLEVEL.ERR, String.Format("Application_Start, ARS v{0}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()))
        ml.write(Ars.LOGLEVEL.ERR, String.Format("Application_Start, app_logLevel = {0}", ApplicationSettings.logLevel))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, app_BackgroundWorkerSeconds = {0}", ApplicationSettings.BackgroundWorkerSeconds))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, sql_UserRefreshSeconds = {0}", ApplicationSettings.SqlRefreshSeconds))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, ldap_UserCacheSeconds = {0}", ApplicationSettings.adUserCacheSeconds))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, ldap_GroupCacheSeconds = {0}", ApplicationSettings.adGroupCacheSeconds))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, pwd_Length = {0}", ApplicationSettings.pwdLength))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, pwd_SpecialCharacters = {0}", ApplicationSettings.pwdSpecialCharacters))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, ldap_SearchLimit = {0}", ApplicationSettings.SearchResultLimit))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, ldap_Port = {0}", ApplicationSettings.LdapPort))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, ldap_Server = {0}", ApplicationSettings.LdapServer))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, ldap_GroupSearchBase = {0}", ApplicationSettings.GroupSearchBase))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, ldap_UserSearchBase = {0}", ApplicationSettings.UserSearchBase))
        ml.write(Ars.LOGLEVEL.INFO, String.Format("Application_Start, app_accessgroup_regex = {0}", ApplicationSettings.AccessGroupRegex))

        _worker = New Thread(AddressOf BackgroundWorker)
        _worker.Start()
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)

        Dim ml As New Ars.MethodLogger("Global.asax")
        Dim ex As Exception = Server.GetLastError()
        ml.write(Ars.LOGLEVEL.ERR, String.Format("Application_Error: {0}{1}{2}", ex.Message, vbCrLf, ex.StackTrace))
        ml.done()
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)

        _state = 2
    End Sub

    Sub BackgroundWorker()

        Dim ml As New Ars.MethodLogger("Global.asax")

        _state = 1

        Dim _loop As Integer =
            ApplicationSettings.BackgroundWorkerSeconds

        While _state = 1

            Thread.Sleep(1000)

            _loop += 1

            If _loop >= ApplicationSettings.BackgroundWorkerSeconds Then

                _loop = 0
                DisableAndReset()
            End If
        End While

        ml.done()
    End Sub

    Private Sub DisableAndReset()

        Dim ml As New Ars.MethodLogger("Global.asax")

        Try

            _users.Clear()
            _groups.Clear()
            _assignments.Clear()

            If Not SqlConnection.State = ConnectionState.Open Then

                SqlConnection.Open()
            End If

            Dim lObjSqlCommand As New SqlCommand(
            "SELECT * FROM [ars_assignments]",
            SqlConnection)

            Dim lLstSids As New List(Of String)

            Dim lObjDt As New DataTable
            lObjDt.Load(lObjSqlCommand.ExecuteReader())

            For Each lObjRow In lObjDt.Rows

                Dim _Assignment As New Ars.Assignment With {
                .aID = lObjRow(0),
                .uSID = lObjRow(1),
                .aSID = lObjRow(2),
                .flags = lObjRow(3)
            }

                If _Assignment.IsUser Then

                    If Not lLstSids.Contains(_Assignment.aSID) Then

                        _users.Add(
                        New Ars.AssignedUser With {
                            .SID = _Assignment.aSID,
                            .flags = _Assignment.flags
                        })

                        lLstSids.Add(_Assignment.aSID)
                    End If
                ElseIf _Assignment.IsGroup Then

                    _groups.Add(_Assignment.aSID)
                End If

                _assignments.Add(_Assignment)
            Next

            lObjSqlCommand.Dispose()
            SqlConnection.Close()

            ldap.LoadLdapUser(_users)

            _users.ForEach(
            Sub(u)
                u.memberOf.ForEach(
                    Sub(m)
                        If Not _groups.Contains(m) Then _groups.Add(m)
                    End Sub
                )
            End Sub
        )

            ldap.CacheLdapGroupByName(_groups)

            For Each _user In _users.Where(
            Function(a) (a.AutoDisableActive = True Or a.AutoPasswordActive = True))

                _user._Assignments = _assignments.FindAll(
                Function(a) a.uSID = _user.SID And a.IsGroup)

                ldap.AddGroupAssignments(_user, _user.memberOf)

                _user.Status =
                _user.Status And Not 1

                ml.write(Ars.LOGLEVEL.DEBUG,
                     String.Format("Checking user: {0}, {1}", _user.SID, _user.Name))

                If Not _users_processed.ContainsKey(_user.SID) Then

                    _user.Status =
                    _user.Status Or 1

                    ml.write(Ars.LOGLEVEL.DEBUG,
                    String.Format("{0}, not currently processed.", _user.SID))
                Else

                    If _users_processed(_user.SID).UAC <> _user.UAC Or
                   _users_processed(_user.SID).flags <> _user.flags Or
                   _users_processed(_user.SID).pwdLastSet > _user.pwdLastSet Or
                   _users_processed(_user.SID).pwdLastSet = 0 Then

                        _user.Status =
                        _user.Status Or 1

                        ml.write(Ars.LOGLEVEL.DEBUG,
                        String.Format("{0}, was processed, but is changed.", _user.SID))
                    End If
                End If

                If _user.Status.arsIsBitSet(1) Then

                    ldap.ProcessAdvancedUserFeatures(_user)

                    If _users_processed.ContainsKey(_user.SID) Then

                        _users_processed(_user.SID) = _user
                    Else

                        _users_processed.TryAdd(_user.SID, _user)
                    End If
                End If
            Next

        Catch ex As Exception

            ml.write(Ars.LOGLEVEL.ERR,
                 String.Format("ERROR DisableAndReset, {0}", ex.Message))

            If Not SqlConnection.State = ConnectionState.Closed Then

                SqlConnection.Close()
            End If
        End Try

        ml.done()
    End Sub

End Class
