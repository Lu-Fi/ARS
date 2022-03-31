Imports System.Collections.Concurrent
Imports System.Data.SqlClient
Imports System.DirectoryServices.Protocols
Imports System.Net
Imports System.Security.Principal
Imports System.Threading

Public Class ArsGlobal : Inherits System.Web.HttpApplication

    Private _state As Integer = 0
    Private _worker As Thread = Nothing
    Private _OrphanSidWorker As Thread = Nothing

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


        Dim _OrphanSidLastRunDay As Integer = 0

        While _state = 1

            Thread.Sleep(1000)

            _loop += 1

            If _loop >= ApplicationSettings.BackgroundWorkerSeconds Then

                _loop = 0
                DisableAndReset()
            End If

            If _OrphanSidLastRunDay < Now().DayOfYear And ApplicationSettings.OrphanSidRemovalDays > 0 Then

                'Ensure no other thread is running
                If _OrphanSidWorker IsNot Nothing Then

                    If _OrphanSidWorker.IsAlive Then

                        _OrphanSidWorker.Abort()
                    End If
                End If

                _OrphanSidWorker = New Thread(AddressOf CleanupOrphanSidList)
                _OrphanSidWorker.Start()

                _OrphanSidLastRunDay = Now().DayOfYear
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
                .Flags = lObjRow(3)
            }

                If _Assignment.IsUser Then

                    If Not lLstSids.Contains(_Assignment.aSID) Then

                        _users.Add(
                        New Ars.AssignedUser With {
                            .SID = _Assignment.aSID,
                            .Flags = _Assignment.Flags
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
                   _users_processed(_user.SID).Flags <> _user.Flags Or
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

    Public Sub CleanupOrphanSidList()

        Dim ml As New Ars.MethodLogger("Global.asax")

        Try

            If Not SqlConnection.State = ConnectionState.Open Then

                SqlConnection.Open()
            End If

            'First check all known orphan SID's again
            Dim lObjOrphanSidsDt As New DataTable
            With New SqlCommand(
                        "SELECT * FROM [ars_orphansids]",
                        SqlConnection)
                lObjOrphanSidsDt.Load(.ExecuteReader())
                .Dispose()
            End With

            For Each lObjRow In lObjOrphanSidsDt.Rows
                Try

                    With New SecurityIdentifier(lObjRow(0).ToString()).Translate(GetType(NTAccount))

                        ml.write(Ars.LOGLEVEL.DEBUG,
                            String.Format("Remove SID from orphan SID table, {0} ({1})", lObjRow(0).ToString(), .ToString()))

                        With New SqlCommand(
                                    "DELETE FROM [ars_orphansids] WHERE [oSID] = @SID", SqlConnection)
                            .Parameters.AddWithValue("SID", lObjRow(0).ToString())
                            .ExecuteNonQuery()
                            .Dispose()
                        End With
                    End With
                Catch ex As Exception
                End Try
            Next

            lObjOrphanSidsDt.Dispose()

            'Check all SID's and write unresolvable into DB
            Dim lObjAllSidsDt As New DataTable
            With New SqlCommand("
                        WITH T1 AS ( 
                          SELECT [uSID] AS _SID FROM [ars_assignments]
                          UNION ALL
                          SELECT [aSID] AS _SID FROM [ars_assignments]
                          UNION ALL
                          SELECT [uSID] AS _SID FROM [ars_user]
                          UNION ALL
                          SELECT [uSID] AS _SID FROM [ars_users]
                        ) SELECT * FROM T1 GROUP BY _SID",
                        SqlConnection)

                lObjAllSidsDt.Load(.ExecuteReader())
                .Dispose()
            End With

            For Each lObjRow In lObjAllSidsDt.Rows
                Try

                    With New SecurityIdentifier(lObjRow(0).ToString()).Translate(GetType(NTAccount))

                        ml.write(Ars.LOGLEVEL.DEBUG,
                            String.Format("{0} = {1}", lObjRow(0).ToString(), .ToString()))
                    End With
                Catch ex As Exception

                    With New SqlCommand(
                                "IF NOT EXISTS ( SELECT [oSID] FROM [ars_orphansids] WHERE [oSID] = @SID )
                                 BEGIN
                                   INSERT INTO [ars_orphansids] (oSID, ts) VALUES (@SID, GetDate()) 
                                 END",
                                SqlConnection)
                        .Parameters.AddWithValue("SID", lObjRow(0).ToString())
                        .ExecuteNonQuery()
                        .Dispose()
                    End With

                    ml.write(Ars.LOGLEVEL.DEBUG,
                        String.Format("OrphanSid = {0}", lObjRow(0).ToString()))
                End Try
            Next

            lObjAllSidsDt.Dispose()

            'log if something will be deleted
            Dim lObjDataReader As SqlClient.SqlDataReader

            With New SqlCommand(
                                "SELECT * FROM [ars_assignments] 
	                               WHERE 
		                             EXISTS (SELECT * FROM [ars_orphansids] WHERE ( [oSID] = [uSID] OR [oSID] = [aSID] ) AND [ts] < DATEADD(day, @DAYS, GETDATE()));",
                                SqlConnection)
                .Parameters.AddWithValue("DAYS", -ApplicationSettings.OrphanSidRemovalDays)
                lObjDataReader = .ExecuteReader()
                While lObjDataReader.Read()

                    ml.write(Ars.LOGLEVEL.INFO,
                        String.Format("Orpan SID removed from [ars_assignments], aID: {0}, uSID: {1}, aSID: {2}, flags: {3}", lObjDataReader.Item(0), lObjDataReader.Item(1), lObjDataReader.Item(2), lObjDataReader.Item(3)))
                End While
                lObjDataReader.Close()
                .Dispose()
            End With

            With New SqlCommand(
                                "SELECT * FROM [ars_users] 
	                               WHERE 
		                             EXISTS (SELECT * FROM [ars_orphansids] WHERE [oSID] = [uSID] AND [ts] < DATEADD(day, @DAYS, GETDATE()))",
                                SqlConnection)
                .Parameters.AddWithValue("DAYS", -ApplicationSettings.OrphanSidRemovalDays)
                lObjDataReader = .ExecuteReader()
                While lObjDataReader.Read()

                    ml.write(Ars.LOGLEVEL.INFO,
                        String.Format("Orpan SID removed from [ars_users], uSID: {0}, flags: {1}", lObjDataReader.Item(0), lObjDataReader.Item(1)))
                End While
                lObjDataReader.Close()
                .Dispose()
            End With

            'Delete outdated orphan SID's 
            With New SqlCommand(
                                "DELETE FROM [ars_assignments] 
	                               WHERE 
		                             EXISTS (SELECT * FROM [ars_orphansids] WHERE ( [oSID] = [uSID] OR [oSID] = [aSID] ) AND [ts] < DATEADD(day, @DAYS, GETDATE()));
                                 DELETE FROM [ars_user] 
	                               WHERE 
		                             EXISTS (SELECT * FROM [ars_orphansids] WHERE [oSID] = [uSID] AND [ts] < DATEADD(day, @DAYS, GETDATE()));
                                 DELETE FROM [ars_users] 
	                               WHERE 
		                             EXISTS (SELECT * FROM [ars_orphansids] WHERE [oSID] = [uSID] AND [ts] < DATEADD(day, @DAYS, GETDATE()));
                                 
                                 DELETE FROM [ars_orphansids] WHERE [ts] < DATEADD(day, @DAYS, GETDATE());",
                                SqlConnection)
                .Parameters.AddWithValue("DAYS", -ApplicationSettings.OrphanSidRemovalDays)
                .ExecuteNonQuery()
                .Dispose()
            End With

            SqlConnection.Close()
        Catch ex As Exception

            If Not SqlConnection.State = ConnectionState.Closed Then

                SqlConnection.Close()
            End If

            ml.write(Ars.LOGLEVEL.ERR, ex.Message)
            ml.write(Ars.LOGLEVEL.ERR, ex.StackTrace)
        End Try

        ml.done()
    End Sub
End Class
