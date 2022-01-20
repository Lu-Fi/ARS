Imports Newtonsoft.Json
Imports System.Data.SqlClient
Imports System.Security.Principal
Imports System.DirectoryServices.Protocols

Namespace Ars

    Public Class AssignmentFlags

        Public flags As Integer

        Public ReadOnly Property IsUser As Integer
            Get

                Return Flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT)
            End Get
        End Property

        Public ReadOnly Property IsGroup As Integer
            Get

                Return Flags.arsIsBitSet(ASS_FLAG.GROUP_ASSIGNMENT)
            End Get
        End Property

        Public ReadOnly Property IsSetting As Integer
            Get

                Return flags.arsIsBitSet(ASS_FLAG.SETTINGS_ASSIGNMENT)
            End Get
        End Property

        Public ReadOnly Property AutoDisableActive As Integer
            Get

                Return Flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_AUTO_DISABLE_ACTIVE)
            End Get
        End Property

        Public ReadOnly Property AutoPasswordActive As Integer
            Get

                Return Flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_AUTO_PASSWORD_ACTIVE)
            End Get
        End Property

        Public ReadOnly Property IsAdmin As Integer
            Get

                Return Flags.arsIsBitSet(USER_FLAG.ADMIN)
            End Get
        End Property

        Public ReadOnly Property AssignedByGroup As Integer
            Get

                Return flags.arsIsBitSet(USER_FLAG.ASSIGNED_BY_GROUP)
            End Get
        End Property
    End Class

    Public Class Assignment : Inherits AssignmentFlags

        Public aID As Guid
        Public uSID As String
        Public aSID As String
        Public source As ASS_SOURCE
        Public updateId As Integer = -1
    End Class

    Public Class AuditLogEntry

        Public Time As String
        Public Account As String
        Public Source As String
        Public AssignedAccount As String
        Public Action As String
        Public Result As Integer
    End Class

    Public Class User : Inherits AssignedUser

        <JsonIgnore>
        Public SqlConnection As SqlConnection = Nothing
        <JsonIgnore>
        Friend HttpContext As HttpContext = Nothing
        <JsonIgnore>
        Public ApplicationSettings As ApplicationSettings
        <JsonIgnore>
        Friend SqlTimestamp As DateTime =
            DateAdd(DateInterval.Year, -1, Now)
        <JsonIgnore>
        Private AdAssignmentTimestamp As DateTime =
            DateAdd(DateInterval.Year, -1, Now)
        <JsonIgnore>
        Public ldap As New Ldap()
        <JsonIgnore>
        Public Admin As Admin
        <JsonIgnore>
        Public updateId As Integer = 0

        Public Sub New(lStrSid As String, Optional _HttpContext As HttpContext = Nothing, Optional IsAdminObject As Boolean = False)

            SID = lStrSid
            User = Me

            If _HttpContext IsNot Nothing Then

                HttpContext = _HttpContext
            Else

                HttpContext = HttpContext.Current
            End If

            ApplicationSettings = HttpContext.Application("ApplicationSettings")

            SqlConnection =
                New SqlConnection(ApplicationSettings.SqlConnectionString)

            InitializeUser()

            If IsAdmin Or IsAdminObject Then

                Me.Admin =
                    New Admin(Me)
            End If
        End Sub

        <JsonIgnore>
        Public ReadOnly Property IsAllowed() As Boolean
            Get

                InitializeUser()

                Return Flags.arsIsBitSet(USER_FLAG.ACTIVE)
            End Get
        End Property

        Public ReadOnly Property AssignedUsers As List(Of AssignedUser)
            Get

                'InitializeUser()

                Dim _AssignedUsers As New List(Of AssignedUser)

                For Each _AssignedUser In _Assignments.FindAll(Function(u) (u.IsUser And u.uSID = SID))

                    _AssignedUsers.Add(New AssignedUser With {
                            .aID = _AssignedUser.aID,
                            .SID = _AssignedUser.aSID,
                            .flags = _AssignedUser.flags,
                            .source = _AssignedUser.source,
                            .User = Me,
                            ._Assignments =
                                _Assignments.FindAll(
                                    Function(a) a.uSID = _AssignedUser.aSID)
                        })
                Next

                ldap.LoadLdapUser(_AssignedUsers)

                Return _AssignedUsers
            End Get
        End Property

        Public ReadOnly Property Users() As List(Of AssignedUser)
            Get

                Dim lLstUsers As New List(Of AssignedUser)
                Dim lIntId As Integer = DateTime.Now.ToString("mmssfff")

                Try

                    If SID.Length > 0 Then

                        If Not SqlConnection.State = ConnectionState.Open Then

                            SqlConnection.Open()
                        End If

                        Dim lObjSqlCommand As New SqlCommand(
                            "SELECT [aID], [uSID], [aSID], [flag] FROM [ars_assignments] WHERE aSID = @aSID",
                            SqlConnection)

                        lObjSqlCommand.Parameters.AddWithValue("aSID", SID)

                        Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

                        While lObjReader.Read()
                            lLstUsers.Add(New AssignedUser With {
                                .aID = lObjReader.GetGuid(0),
                                .SID = lObjReader.GetString(1),
                                .flags = lObjReader.GetInt32(3)
                            })
                        End While

                        lObjReader.Close()
                        lObjSqlCommand.Dispose()
                        SqlConnection.Close()

                        ldap.LoadLdapUser(lLstUsers)
                    End If

                Catch ex As Exception

                    If Not SqlConnection.State = ConnectionState.Closed Then

                        SqlConnection.Close()
                    End If
                End Try

                Return lLstUsers
            End Get
        End Property

        Public Sub SetHttpContext(lObjHttpContext As HttpContext)

            HttpContext = lObjHttpContext
        End Sub

        Public Function GetAuditLog(lObjFormData As NameValueCollection, Optional lIntIsAdmin As Integer = 0) As Dictionary(Of String, Object)

            Dim lIntRowsTotal As Integer = 0
            Dim lIntRowsFiltered As Integer = 0
            Dim lIntLength As Integer = 0
            Dim lIntStart As Integer = 0
            Dim lStrSearch As String = ""

            Dim ml As New MethodLogger("ArsUser")
            Dim aSIDs As New List(Of String)
            Dim lDictResult As New Dictionary(Of String, Object) From {{"recordsTotal", 0}, {"recordsFiltered", 0}}
            Dim lLstData As New List(Of AuditLogEntry)

            If lObjFormData Is Nothing Then Return Nothing

            If Not IsAdmin And lIntIsAdmin <> 0 Then

                lIntIsAdmin = 0
            End If

            If lObjFormData("length") IsNot Nothing And lObjFormData("start") IsNot Nothing Then

                Int32.TryParse(lObjFormData("length"), lIntLength)
                Int32.TryParse(lObjFormData("start"), lIntStart)
            End If

            If Not String.IsNullOrEmpty(lObjFormData("search[value]")) Then

                lStrSearch = lObjFormData("search[value]")
                If lStrSearch.Length > 64 Then

                    lStrSearch = lStrSearch.Substring(0, 64)
                End If

                lStrSearch = lStrSearch & "%"
            End If

            For Each _Assignment In _Assignments.FindAll(Function(e) e.IsUser And e.uSID = SID)

                aSIDs.Add(String.Format("'{0}'", _Assignment.aSID))
            Next

            If lIntIsAdmin = 1 Or aSIDs.Count > 0 Then

                If Not SqlConnection.State = ConnectionState.Open Then

                    SqlConnection.Open()
                End If

                Dim lObjSqlCommand As New SqlCommand(
                        "SELECT COUNT(*) FROM [ars_audit] WHERE ([aSID] in (@aSIDs) OR @ADMIN = 1);
                     SELECT COUNT(*) FROM [ars_audit] WHERE ([aSID] in (@aSIDs) OR @ADMIN = 1) AND
                        ( 
                               (CONVERT(VARCHAR, [ts], 120) like @search OR @search = '')
                            OR ([uName] like @search OR @search = '')
                            OR ([uSource] like @search OR @search = '')
                            OR ([aName] like @search OR @search = '')
                            OR ([action] like @search OR @search = '')
                            OR ([result] like @search OR @search = '')
                        );
                     SELECT [ts], [uName], [uSource], [aName], [action], [result]
                        FROM [ars_audit] WHERE ([aSID] in (@aSIDs) OR @ADMIN = 1) AND
                        ( 
                               (CONVERT(VARCHAR, [ts], 120) like @search OR @search = '')
                            OR ([uName] like @search OR @search = '')
                            OR ([uSource] like @search OR @search = '')
                            OR ([aName] like @search OR @search = '')
                            OR ([action] like @search OR @search = '')
                            OR ([result] like @search OR @search = '')
                        )
                        ORDER BY [ts] DESC
                        OFFSET @OFFSET ROWS FETCH NEXT @ROWS ROWS ONLY",
                        SqlConnection)

                lObjSqlCommand.Parameters.AddWithValue("ROWS", lIntLength)
                lObjSqlCommand.Parameters.AddWithValue("OFFSET", lIntStart)
                lObjSqlCommand.Parameters.AddWithValue("ADMIN", lIntIsAdmin)
                lObjSqlCommand.Parameters.AddWithValue("search", lStrSearch)
                lObjSqlCommand.CommandText =
                    lObjSqlCommand.CommandText.Replace("@aSIDs", String.Join(",", aSIDs))

                Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

                While lObjReader.Read()

                    lIntRowsTotal = lObjReader.GetInt32(0)
                End While

                lDictResult("recordsTotal") = lIntRowsTotal

                'Return Second Query
                lObjReader.NextResult()

                While lObjReader.Read()

                    lIntRowsTotal = lObjReader.GetInt32(0)
                End While

                lDictResult("recordsFiltered") = lIntRowsTotal

                'Return Third Query
                lObjReader.NextResult()

                While lObjReader.Read()

                    lLstData.Add(
                        New AuditLogEntry With {
                            .Time = lObjReader.GetDateTime(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),
                            .Account = lObjReader.GetString(1),
                            .Source = lObjReader.GetString(2),
                            .AssignedAccount = lObjReader.GetString(3),
                            .Action = lObjReader.GetString(4),
                            .Result = lObjReader.GetInt32(5)
                    })
                End While

                lDictResult.Add("data", lLstData)

                lObjReader.Close()
                lObjSqlCommand.Dispose()
                SqlConnection.Close()
            End If

            Return lDictResult

        End Function

        Private Sub InitializeUser(Optional forceUpdate As Boolean = False, Optional UpdateRoleGroups As Boolean = False)

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("{0}, {1}", forceUpdate, UpdateRoleGroups))

            Dim lIntlastRefresh As Integer =
                (Now() - SqlTimestamp).TotalSeconds

            If lIntlastRefresh < ApplicationSettings.SqlRefreshSeconds And forceUpdate = False Then

                ml.write(LOGLEVEL.DEBUG,
                         String.Format("The last refresh was {0} seconds ago. exit.", lIntlastRefresh))
                ml.done()
                Exit Sub
            End If

            SqlTimestamp = Now()

            Try

                If SID.Length > 0 Then

                    Dim lBolAssignedByGroup As Boolean = False

                    'Get User from LDAP
                    ldap.RefreshUserFromLdap(Me)

                    lIntlastRefresh =
                        (Now() - AdAssignmentTimestamp).TotalSeconds

                    ml.write(LOGLEVEL.DEBUG,
                        String.Format("The last assignment refresh was less than {0} seconds ago.", lIntlastRefresh))

                    'Check Access by Active Directory Group
                    If ApplicationSettings.AccessGroupRegex IsNot Nothing Then

                        With Me.memberOf.FindIndex(
                            Function(g) Regex.IsMatch(g.ToString(), ApplicationSettings.AccessGroupRegex))

                            If Not .Equals(-1) Then

                                lBolAssignedByGroup = True
                            End If
                        End With
                    End If

                    'Set UpdateId for all existing Assignments, this is needed to find 
                    'assignments no longer valid 
                    updateId += 1

                    'Get User from database
                    If Not SqlConnection.State = ConnectionState.Open Then

                        SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "SELECT TOP(1) * FROM [ars_users] WHERE [uSID] = @uSID",
                            SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("uSID", SID)

                    Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

                    If lObjReader.HasRows Then

                        While lObjReader.Read()

                            flags = lObjReader.GetInt32(1)
                        End While
                    Else

                        If lBolAssignedByGroup = True Then

                            flags = USER_FLAG.ACTIVE Or USER_FLAG.ASSIGNED_BY_GROUP
                        End If
                    End If

                    lObjReader.Close()

                    'Load user assignments from database
                    lObjSqlCommand.CommandText =
                    "SELECT [aID], [uSID], [aSID], [flag] FROM [ars_assignments] WHERE
                        uSID = @uSID Or 
	                    uSID in (SELECT aSID FROM [ars_assignments] WHERE uSID = @uSID)"

                    Dim lObjDt As New DataTable
                    lObjDt.Load(lObjSqlCommand.ExecuteReader())

                    _Assignments.RemoveAll(
                        Function(a) a.source = ASS_SOURCE.SQL)

                    For Each lObjRow In lObjDt.Rows

                        If (CInt(lObjRow(3)).arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT) = True Or
                        CInt(lObjRow(3)).arsIsBitSet(ASS_FLAG.GROUP_ASSIGNMENT) = True) And
                        CInt(lObjRow(3)).arsIsBitSet(ASS_FLAG.SETTINGS_ASSIGNMENT) = False Then

                            _Assignments.Add(New Assignment With {
                                .aID = lObjRow(0),
                                .uSID = lObjRow(1),
                                .aSID = lObjRow(2),
                                .source = ASS_SOURCE.SQL,
                                .flags = lObjRow(3),
                                .updateId = updateId
                            })
                        End If
                    Next

                    lObjReader.Close()
                    lObjSqlCommand.Dispose()
                    SqlConnection.Close()

                    'Load user assignments from ActiveDirectory
                    If (lIntlastRefresh > ApplicationSettings.AdAssignmentRefreshSeconds) Or
                        UpdateRoleGroups = True Or forceUpdate = True Then

                        '_Assignments.RemoveAll(
                        'Function(a) a.source = ASS_SOURCE.LDAP)

                        ldap.GetLdapUserAssignments(Me)
                        ldap.GetLdapGroupAssignments(Me)

                        AdAssignmentTimestamp = Now()

                        'Remove all assignments not updated
                        _Assignments.RemoveAll(
                            Function(a) a.updateId > 0 And a.updateId < updateId)
                    End If

                    For Each lObjRow In lObjDt.Rows

                        If CInt(lObjRow(3)).arsIsBitSet(ASS_FLAG.SETTINGS_ASSIGNMENT) = True Then

                            _Assignments.FindAll(
                            Function(a) a.uSID = lObjRow(1) And a.aSID = lObjRow(2) And a.source = ASS_SOURCE.LDAP).ForEach(
                                Sub(a)
                                    a.aID = lObjRow(0)
                                    a.flags = lObjRow(3)
                                End Sub
                        )
                        End If
                    Next
                End If
            Catch ex As Exception

                If Not SqlConnection.State = ConnectionState.Closed Then

                    SqlConnection.Close()
                End If
            End Try

            ml.done()
        End Sub

        Public Function ForceAssignmentRefresh()

            Dim ml As New MethodLogger("ArsUser")

            SqlTimestamp =
                DateAdd(DateInterval.Year, -1, Now)

            AdAssignmentTimestamp =
                DateAdd(DateInterval.Year, -1, Now)

            InitializeUser()

            ml.done()

            Return Me
        End Function

        Public Sub AddMemberToGroup(_AssignmentId As String, ByRef lDicResult As Dictionary(Of String, Object), Optional lIntTime As Integer = 0)

            Dim lDicFncResult As New Dictionary(Of String, Object) From {{"result", 0}}
            lDicResult.Add("AddMemberToGroup", lDicFncResult)

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("assignment = '{0}', ttl = '{1}'", _AssignmentId, lIntTime))

            Try

                'InitializeUser()

                Dim _Assignment As Assignment =
                    _Assignments.FirstOrDefault(Function(a) a.aID.ToString() = _AssignmentId)

                If _Assignment IsNot Nothing Then

                    Dim _AssignedUser As AssignedUser =
                        AssignedUsers.FirstOrDefault(Function(a) a.SID = _Assignment.uSID)

                    If _AssignedUser IsNot Nothing Then

                        Dim _AssignedGroup As AssignedGroup =
                            _AssignedUser.AssignedGroups.FirstOrDefault(Function(a) a.SID = _Assignment.aSID)

                        If _AssignedGroup IsNot Nothing Then

                            lDicFncResult("result") =
                                ldap.AddMemberToGroup(_AssignedUser, _AssignedGroup.DN, lIntTime)

                            ldap.RefreshUserFromLdap(_AssignedUser)

                            If lDicFncResult("result") = 1 Then

                                ldap.Cache.UserExpiry(_AssignedUser)
                                ldap.setPwdLastSet(_AssignedUser, 0)
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception

                lDicFncResult.Add("error", ex.Message)
                ml.write(LOGLEVEL.ERR,
                     String.Format("ERROR addMemberToGroup: {0}", ex.Message))
            End Try

            ml.done()

        End Sub

        Public Sub RemoveMemberFromGroup(_AssignmentId As String, ByRef lDicResult As Dictionary(Of String, Object))

            Dim lDicFncResult As New Dictionary(Of String, Object) From {{"result", 0}}
            lDicResult.Add("RemoveMemberFromGroup", lDicFncResult)

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("assignment = '{0}'", _AssignmentId))

            Try

                'InitializeUser()

                Dim _Assignment As Assignment =
                    _Assignments.FirstOrDefault(Function(a) a.aID.ToString() = _AssignmentId)

                If _Assignment IsNot Nothing Then

                    Dim _AssignedUser As AssignedUser =
                        AssignedUsers.FirstOrDefault(Function(a) a.SID = _Assignment.uSID)

                    If _AssignedUser IsNot Nothing Then

                        Dim _AssignedGroup As AssignedGroup =
                            _AssignedUser.AssignedGroups.FirstOrDefault(Function(a) a.SID = _Assignment.aSID)

                        If _AssignedGroup IsNot Nothing Then

                            lDicFncResult("result") =
                                ldap.RemoveMemberFromGroup(_AssignedUser, _AssignedGroup.DN)

                            ldap.RefreshUserFromLdap(_AssignedUser)

                            If lDicFncResult("result") = 1 Then

                                If HttpContext.Cache.Get(_AssignedUser.SID) IsNot Nothing Then

                                    HttpContext.Cache.Remove(_AssignedUser.SID)
                                End If

                                _AssignedUser.Status = 1
                                ldap.ProcessAdvancedUserFeatures(_AssignedUser)
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception

                lDicFncResult.Add("error", ex.Message)
                ml.write(LOGLEVEL.ERR,
                     String.Format("ERROR removeMemberFromGroup: {0}", ex.Message))
            End Try

            ml.done()
        End Sub

        Public Function ResetPassword(_AssignmentId As String) As String

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("assignment = '{0}'", _AssignmentId))

            Dim _return As String = Nothing

            'InitializeUser()

            Dim _Assignment As Assignment =
                    _Assignments.FirstOrDefault(Function(i) i.aID.ToString() = _AssignmentId)

            If Not _Assignment.Flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_RESET_PASSORD_LOCKED) Then

                If _Assignment IsNot Nothing Then

                    Dim _AssignedUser As AssignedUser =
                        AssignedUsers.FirstOrDefault(Function(a) a.SID = _Assignment.aSID)

                    If _AssignedUser IsNot Nothing Then

                        Dim _se As String = System.Web.Security.Membership.GeneratePassword(
                            ApplicationSettings.pwdLength, ApplicationSettings.pwdSpecialCharacters)

                        If Not String.IsNullOrEmpty(_AssignedUser.DN) And Not String.IsNullOrEmpty(_se) Then

                            ldap.setPwdLastSet(_AssignedUser, 0)

                            _return = ldap.ResetPassword(_AssignedUser)

                            ldap.RefreshUserFromLdap(_AssignedUser)
                        End If
                    End If
                End If
            End If

            ml.done()

            Return _return
        End Function

        Public Function toggleAutoDisable(lStrAssignmentId As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("assignment = '{0}'", lStrAssignmentId))

            Try

                'InitializeUser()

                Dim _Assignment As Assignment =
                    _Assignments.FirstOrDefault(Function(i) i.aID.ToString() = lStrAssignmentId)

                If _Assignment IsNot Nothing Then

                    If Not _Assignment.Flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_AUTO_DISABLE_LOCKED) Then

                        Dim lObjSqlCommand As New SqlCommand(
                                        "If NOT EXISTS (SELECT aID FROM [ars_assignments] WHERE aID = @aID )
                                            INSERT INTO [ars_assignments] ([aID], [uSID], [aSID], [flag]) VALUES (@aID, @uSID, @aSID, ( @OPT | 1 | 4096 ))
                                         ELSE
                                            UPDATE TOP(1) [ars_assignments] SET flag = flag ^ @OPT WHERE [aID] = @aID",
                                        SqlConnection)

                        lObjSqlCommand.Parameters.AddWithValue("aID", _Assignment.aID)
                        lObjSqlCommand.Parameters.AddWithValue("uSID", _Assignment.uSID)
                        lObjSqlCommand.Parameters.AddWithValue("aSID", _Assignment.aSID)
                        lObjSqlCommand.Parameters.AddWithValue("OPT", ASS_FLAG.USER_ASSIGNMENT_AUTO_DISABLE_ACTIVE)

                        Try

                            SqlConnection.Open()
                            lObjSqlCommand.ExecuteNonQuery()
                            SqlConnection.Close()
                            'InitializeUser(True)
                            SqlTimestamp =
                                DateAdd(DateInterval.Year, -1, Now)

                            ml.done(" '1'")

                            Return 1
                        Catch ex As Exception

                            SqlConnection.Close()
                        End Try
                    End If
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR toggleAutoDisable, {0}", ex.Message))
            End Try

            ml.done(" '1'")

            Return 0

        End Function

        Public Function toggleAutoPassword(lStrAssignmentId As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("assignment = '{0}'", lStrAssignmentId))

            Try

                'InitializeUser()

                Dim _Assignment As Assignment =
                    _Assignments.FirstOrDefault(Function(i) i.aID.ToString() = lStrAssignmentId)

                If _Assignment IsNot Nothing Then

                    If Not _Assignment.Flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_AUTO_PASSWORD_LOCKED) Then

                        Dim lObjSqlCommand As New SqlCommand(
                                        "If NOT EXISTS (Select aID FROM [ars_assignments] WHERE aID = @aID )
                                            INSERT INTO [ars_assignments] ([aID], [uSID], [aSID], [flag]) VALUES (@aID, @uSID, @aSID, ( @OPT | 1 | 4096 ))
                                         ELSE
                                            UPDATE TOP(1) [ars_assignments] SET flag = flag ^ @OPT WHERE [aID] = @aID",
                                        SqlConnection)

                        lObjSqlCommand.Parameters.AddWithValue("aID", _Assignment.aID)
                        lObjSqlCommand.Parameters.AddWithValue("uSID", _Assignment.uSID)
                        lObjSqlCommand.Parameters.AddWithValue("aSID", _Assignment.aSID)
                        lObjSqlCommand.Parameters.AddWithValue("OPT", ASS_FLAG.USER_ASSIGNMENT_AUTO_PASSWORD_ACTIVE)

                        Try

                            SqlConnection.Open()
                            lObjSqlCommand.ExecuteNonQuery()
                            SqlConnection.Close()
                            'InitializeUser(True)
                            SqlTimestamp =
                                DateAdd(DateInterval.Year, -1, Now)

                            ml.done(" '1'")

                            Return 1
                        Catch ex As Exception

                            SqlConnection.Close()
                        End Try
                    End If
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR toggleAutoPassword, {0}", ex.Message))
            End Try

            ml.done(" '0'")

            Return 0

        End Function

        Public Function ToggleBaseGroupFlag(lStrAssignmentId As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("assignment = '{0}'", lStrAssignmentId))

            Try

                'InitializeUser()
                Dim _aID As New Guid()

                Dim _Assignment As Assignment =
                    _Assignments.FirstOrDefault(Function(i) i.aID.ToString() = lStrAssignmentId)

                Dim _UserAssignment As Assignment =
                    _Assignments.FirstOrDefault(Function(i) i.uSID = SID And i.aSID = _Assignment.uSID)

                If _UserAssignment IsNot Nothing Then

                    _aID = _UserAssignment.aID
                End If

                If _Assignment IsNot Nothing Then

                    If Not _Assignment.flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_AUTO_PASSWORD_LOCKED) Then

                        Dim lObjSqlCommand As New SqlCommand(
                                        "If NOT EXISTS (Select aID FROM [ars_assignments] WHERE aID = @aID )
                                            BEGIN 
                                                If NOT EXISTS (Select aID FROM [ars_assignments] WHERE [uSID] = @_SID AND [aSID] = @uSID)
                                                    INSERT INTO [ars_assignments] ([aID], [uSID], [aSID], [flag]) VALUES (@_aID, @_SID, @uSID, ( 1 | 4096 ));
                                                INSERT INTO [ars_assignments] ([aID], [uSID], [aSID], [flag]) VALUES (@aID, @uSID, @aSID, ( @OPT | 2 | 4096 ))
                                            END
                                         ELSE
                                            UPDATE TOP(1) [ars_assignments] SET flag = flag ^ @OPT WHERE [aID] = @aID",
                                        SqlConnection)

                        lObjSqlCommand.Parameters.AddWithValue("_SID", SID)
                        lObjSqlCommand.Parameters.AddWithValue("_aID", _aID)

                        lObjSqlCommand.Parameters.AddWithValue("aID", _Assignment.aID)
                        lObjSqlCommand.Parameters.AddWithValue("uSID", _Assignment.uSID)
                        lObjSqlCommand.Parameters.AddWithValue("aSID", _Assignment.aSID)
                        lObjSqlCommand.Parameters.AddWithValue("OPT", ASS_FLAG.GROUP_ASSIGNMENT_BASE)

                        Try

                            SqlConnection.Open()
                            lObjSqlCommand.ExecuteNonQuery()
                            SqlConnection.Close()
                            'InitializeUser(True)
                            SqlTimestamp =
                                DateAdd(DateInterval.Year, -1, Now)

                            ml.done(" '1'")

                            Return 1
                        Catch ex As Exception

                            SqlConnection.Close()
                        End Try
                    End If
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR ToggleBaseGroupFlag, {0}", ex.Message))
            End Try

            ml.done(" '0'")

            Return 0

        End Function

        Private Function NewGuid() As String
            Throw New NotImplementedException()
        End Function

        Public Sub enableAssignedUser(lStrAssignmentId As String, ByRef lDicResult As Dictionary(Of String, Object))

            'InitializeUser()

            Dim lDicFncResult As New Dictionary(Of String, Object) From {{"result", 0}}
            lDicResult.Add("enableAssignedUser", lDicFncResult)

            Dim _AssignedUser As AssignedUser =
                AssignedUsers.FirstOrDefault(Function(u) u.aID.ToString() = lStrAssignmentId)

            If _AssignedUser IsNot Nothing Then

                If Not _AssignedUser.flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_MANUAL_DISABLE_LOCKED) Then

                    'ENABLE USER
                    lDicFncResult("result") =
                        ldap.SetAccountUac(_AssignedUser, 2, False)
                Else

                    lDicResult.Add("error", "option is locked for this account")
                End If

                If lDicFncResult("result") < 2 Then

                    audit.add(_AssignedUser, "Enable user", lDicFncResult("result"))
                End If
            End If
        End Sub

        Public Sub disableAssignedUser(lStrAssignmentId As String, ByRef lDicResult As Dictionary(Of String, Object))

            'InitializeUser()

            Dim lDicFncResult As New Dictionary(Of String, Object) From {{"result", 0}}
            lDicResult.Add("disableAssignedUser", lDicFncResult)

            Dim _AssignedUser As AssignedUser =
                AssignedUsers.FirstOrDefault(Function(u) u.aID.ToString() = lStrAssignmentId)

            If _AssignedUser IsNot Nothing Then

                If Not _AssignedUser.Flags.arsIsBitSet(ASS_FLAG.USER_ASSIGNMENT_MANUAL_DISABLE_LOCKED) Then

                    'DISABLE USER

                    lDicFncResult("result") =
                        ldap.SetAccountUac(_AssignedUser, 2, True)
                Else

                    lDicResult.Add("error", "option is locked for this account")
                End If

                If lDicFncResult("result") < 2 Then

                    audit.add(_AssignedUser, "Disable user", lDicFncResult("result"))
                End If
            End If
        End Sub

        'ADMIN
        'Public Function searchAd(lLstrClasses As List(Of String), lStrSearchString As String) As List(Of Dictionary(Of String, String))

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("classes = '{0}', search = '{1}'", String.Join(",", lLstrClasses), lStrSearchString))

        '    Dim lDicReturn As New List(Of Dictionary(Of String, String))

        '    If IsAdmin Then

        '        Dim lObjLdapConn As LdapConnection =
        '            ldap.getLdapConnection()

        '        If lObjLdapConn IsNot Nothing Then

        '            Try

        '                Dim lStrFilter As String = Nothing

        '                For Each lStrClass In lLstrClasses

        '                    lStrFilter +=
        '                        String.Format("(objectClass={0})", lStrClass)
        '                Next

        '                lStrFilter = String.Format("(&(|{0})(anr={1}*))", lStrFilter, lStrSearchString)

        '                Dim lLstUserAttributes As String() =
        '                    {"name", "description", "displayname", "company", "department", "objectSid", "objectClass", "UserPrincipalName"}

        '                Dim lObjSearchRequest As New SearchRequest(
        '                ApplicationSettings.UserSearchBase,
        '                lStrFilter,
        '                System.DirectoryServices.Protocols.SearchScope.Subtree,
        '                lLstUserAttributes)

        '                lObjSearchRequest.SizeLimit =
        '                    ApplicationSettings.SearchResultLimit

        '                Dim lObjResponse As SearchResponse =
        '                lObjLdapConn.SendRequest(lObjSearchRequest)

        '                If lObjResponse.Entries.Count > 0 Then

        '                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

        '                        Dim lObjResultEntry As New Dictionary(Of String, String) From {
        '                            {"sid", New SecurityIdentifier(lObjEntry.Attributes("objectSid")(0), 0).ToString()},
        '                            {"description", ""},
        '                            {"type", "2"},
        '                            {"name", lObjEntry.Attributes("name")(0)}
        '                        }

        '                        If lObjEntry.Attributes.Contains("description") Then

        '                            lObjResultEntry("description") =
        '                                lObjEntry.Attributes("description")(0)
        '                        End If

        '                        If Not ldap.Cache.getListFromAttribute(lObjEntry.Attributes, "objectclass").Contains("group") Then

        '                            lObjResultEntry("type") = "1"

        '                            If lObjEntry.Attributes.Contains("displayname") Then

        '                                lObjResultEntry.Add("displayname", lObjEntry.Attributes("displayname")(0))
        '                            Else

        '                                lObjResultEntry.Add("displayname", "")
        '                            End If

        '                            If lObjEntry.Attributes.Contains("company") Then

        '                                lObjResultEntry.Add("company", lObjEntry.Attributes("company")(0))
        '                            Else

        '                                lObjResultEntry.Add("company", "")
        '                            End If

        '                            If lObjEntry.Attributes.Contains("department") Then

        '                                lObjResultEntry.Add("department", lObjEntry.Attributes("department")(0))
        '                            Else

        '                                lObjResultEntry.Add("department", "")
        '                            End If

        '                            If lObjEntry.Attributes.Contains("UserPrincipalName") Then

        '                                lObjResultEntry.Add("UserPrincipalName", lObjEntry.Attributes("UserPrincipalName")(0))
        '                            Else

        '                                lObjResultEntry.Add("UserPrincipalName", "")
        '                            End If
        '                        End If

        '                        lDicReturn.Add(lObjResultEntry)
        '                    Next
        '                End If

        '                lObjLdapConn.Dispose()

        '            Catch ex As Exception

        '            End Try
        '        End If
        '    End If

        '    ml.done()

        '    Return lDicReturn

        'End Function

        'Public Function GetAccountStatus(_SID As String) As Integer

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("SID = '{0}'", _SID))

        '    If Not IsAdmin Then

        '        ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
        '        Return -99
        '    End If

        '    Dim lIntReturn As Integer = -1

        '    Try

        '        If SID.Length > 0 Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                    "SELECT TOP(1) * FROM [ars_users] WHERE [uSID] = @uSID",
        '                    SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("uSID", _SID)

        '            Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

        '            While lObjReader.Read()

        '                lIntReturn = lObjReader.GetInt32(1)
        '            End While

        '            lObjReader.Close()
        '            SqlConnection.Close()
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR GetAccountStatus, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return lIntReturn
        'End Function

        'Public Function AddUserAssignment(_uSID As String, _aSID As String) As Integer

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("uSID = '{0}', aSID = '{1}'", _uSID, _aSID))

        '    If Not IsAdmin Then

        '        ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
        '        Return -99
        '    End If

        '    Try

        '        If _uSID.Length > 0 And _aSID.Length > 0 Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                    "IF NOT EXISTS (SELECT uSID FROM [ars_users] WHERE uSID = @uSID)
        '                     INSERT INTO [ars_users] ([uSID], [flag]) VALUES (@uSID, '1')
        '                     ELSE
        '                     UPDATE [ars_users] SET [flag] = [flag] | 1 WHERE uSID = @uSID;

        '                     IF NOT EXISTS (SELECT uSID FROM [ars_assignments] WHERE uSID = @uSID AND aSID = @aSID)
        '                     INSERT INTO [ars_assignments] ([uSID], [aSID], [flag]) VALUES (@uSID, @aSID, '1')
        '                     ELSE
        '                     UPDATE [ars_assignments] SET [flag] = [flag] | 1 WHERE uSID = @uSID AND aSID = @aSID;
        '                    ",
        '                    SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)
        '            lObjSqlCommand.Parameters.AddWithValue("aSID", _aSID)

        '            lObjSqlCommand.ExecuteNonQuery()

        '            SqlConnection.Close()

        '            SqlTimestamp =
        '                DateAdd(DateInterval.Day, -1, Now)

        '            Return 1
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR AddUserAssignment, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return 0
        'End Function

        'Public Function AddGroupAssignment(_uSID As String, _aSID As String) As Integer

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("uSID = '{0}', aSID = '{1}'", _uSID, _aSID))

        '    If Not IsAdmin Then

        '        ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
        '        Return -99
        '    End If

        '    Try

        '        If _uSID.Length > 0 And _aSID.Length > 0 Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                    "IF NOT EXISTS (SELECT uSID FROM [ars_assignments] WHERE uSID = @uSID AND aSID = @aSID)
        '                     INSERT INTO [ars_assignments] ([uSID], [aSID], [flag]) VALUES (@uSID, @aSID, '2')
        '                     ELSE
        '                     UPDATE [ars_assignments] SET [flag] = [flag] | 2 WHERE uSID = @uSID AND aSID = @aSID;
        '                    ",
        '                    SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)
        '            lObjSqlCommand.Parameters.AddWithValue("aSID", _aSID)

        '            lObjSqlCommand.ExecuteNonQuery()

        '            SqlConnection.Close()

        '            SqlTimestamp =
        '                DateAdd(DateInterval.Day, -1, Now)

        '            Return 1
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR AddGroupAssignment, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return 0
        'End Function

        'Public Function RemoveAssignment(_aID As String) As Integer

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("Assignment = '{0}'", _aID))

        '    If Not IsAdmin Then

        '        ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
        '        Return -99
        '    End If

        '    Try

        '        If Not String.IsNullOrEmpty(_aID) Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                    "DELETE FROM [ars_assignments] WHERE aID = @aID;",
        '                    SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("aID", _aID)

        '            lObjSqlCommand.ExecuteNonQuery()

        '            SqlConnection.Close()

        '            SqlTimestamp =
        '                DateAdd(DateInterval.Day, -1, Now)

        '            Return 1
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR RemoveAssignment, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return 0
        'End Function

        'Public Function GetAccountList() As List(Of AssignedUser)

        '    If Not IsAdmin Then

        '        Return Nothing
        '    End If

        '    Dim _Accounts As New List(Of AssignedUser)

        '    Try

        '        If SID.Length > 0 Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                    "SELECT * FROM [ars_users]",
        '                    SqlConnection)

        '            Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

        '            While lObjReader.Read()

        '                _Accounts.Add(New AssignedUser With {
        '                    .SID = lObjReader.GetString(0),
        '                    .flags = lObjReader.GetInt32(1)
        '                })
        '            End While

        '            lObjReader.Close()
        '            SqlConnection.Close()
        '        End If
        '    Catch ex As Exception

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ldap.LoadLdapUser(_Accounts)

        '    Return _Accounts
        'End Function

        'Public Function RemoveUser(_uSID As String) As Integer

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("uSID = '{0}'", _uSID))

        '    If Not IsAdmin Then

        '        ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
        '        Return -99
        '    End If

        '    Try

        '        If Not String.IsNullOrEmpty(_uSID) Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                    "DELETE FROM [ars_users] WHERE uSID = @uSID;
        '                     DELETE FROM [ars_assignments] WHERE (SELECT TOP(1) [uSID] FROM [ars_users] WHERE uSID = [ars_assignments].[uSID]) IS NULL AND (flag & 1 = 1);
        '                     DELETE M FROM [ars_assignments] M WHERE (SELECT TOP(1) [uSID] FROM [ars_assignments] WHERE [aSID] = M.[uSID] AND ([flag] & 1 = 1)) IS NULL AND (M.[flag] & 2 = 2);",
        '                    SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)

        '            lObjSqlCommand.ExecuteNonQuery()

        '            SqlConnection.Close()

        '            Return 1
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR RemoveUser, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return 0
        'End Function

        'Public Function AddUser(_uSID As String) As Integer

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("uSID = '{0}'", _uSID))

        '    If Not IsAdmin Then

        '        ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
        '        Return -99
        '    End If

        '    Try

        '        If Not String.IsNullOrEmpty(_uSID) Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                    "IF NOT EXISTS (SELECT uSID FROM [ars_users] WHERE uSID = @uSID)
        '                     INSERT INTO [ars_users] ([uSID], [flag]) VALUES (@uSID, '1')",
        '                    SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)

        '            lObjSqlCommand.ExecuteNonQuery()

        '            SqlConnection.Close()

        '            Return 1
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR AddUser, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return 0
        'End Function

        'Public Function ToggleUserFlag(_uSID As String, _flag As Integer) As Integer

        '    '------------------------------------------
        '    '
        '    '  Toggle User Admin / Active Flag
        '    '
        '    '------------------------------------------

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("_uSID = '{0}', flag = '{1}'", _uSID, _flag))

        '    Try

        '        If Not String.IsNullOrEmpty(_uSID) Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                                "UPDATE TOP(1) [ars_users] SET flag = flag ^ @flag WHERE [uSID] = @uSID",
        '                                SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)
        '            lObjSqlCommand.Parameters.AddWithValue("flag", _flag)

        '            lObjSqlCommand.ExecuteNonQuery()

        '            SqlConnection.Close()

        '            SqlTimestamp =
        '                DateAdd(DateInterval.Day, -1, Now)

        '            Return 1
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR ToggleFlag, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return 0
        'End Function

        'Public Function ToggleAssignmentFlag(_aID As String, _flag As Integer) As Integer

        '    '------------------------------------------
        '    '
        '    '  Toggle Assignment Flags
        '    '
        '    '------------------------------------------

        '    Dim ml As New MethodLogger("ArsUser",
        '                               String.Format("assignment = '{0}', flag = '{1}'", _aID, _flag))

        '    Try

        '        If Not String.IsNullOrEmpty(_aID) Then

        '            If Not SqlConnection.State = ConnectionState.Open Then

        '                SqlConnection.Open()
        '            End If

        '            Dim _Assignment As Assignment =
        '                _Assignments.FirstOrDefault(Function(e) e.aID.ToString() = _aID)

        '            Dim lIntType As Integer = ASS_FLAG.USER_ASSIGNMENT

        '            If _flag = ASS_FLAG.GROUP_ASSIGNMENT_BASE Then

        '                lIntType = ASS_FLAG.GROUP_ASSIGNMENT
        '            End If

        '            Dim lObjSqlCommand As New SqlCommand(
        '                                "If NOT EXISTS (Select aID FROM [ars_assignments] WHERE aID = @aID )
        '                                    INSERT INTO [ars_assignments] ([aID], [uSID], [aSID], [flag]) VALUES (@aID, @uSID, @aSID, (@flag | 4096 | @type))
        '                                 ELSE
        '                                    UPDATE TOP(1) [ars_assignments] SET flag = flag ^ @flag WHERE [aID] = @aID",
        '                                SqlConnection)

        '            lObjSqlCommand.Parameters.AddWithValue("aID", _Assignment.aID)
        '            lObjSqlCommand.Parameters.AddWithValue("uSID", SID)
        '            lObjSqlCommand.Parameters.AddWithValue("aSID", _Assignment.aSID)
        '            lObjSqlCommand.Parameters.AddWithValue("type", lIntType)
        '            lObjSqlCommand.Parameters.AddWithValue("flag", _flag)

        '            lObjSqlCommand.ExecuteNonQuery()

        '            SqlConnection.Close()

        '            SqlTimestamp =
        '                DateAdd(DateInterval.Day, -1, Now)
        '            Return 1
        '        End If
        '    Catch ex As Exception

        '        ml.write(LOGLEVEL.ERR, String.Format("ERROR ToggleAssignmentFlag, {0}", ex.Message))

        '        If Not SqlConnection.State = ConnectionState.Closed Then

        '            SqlConnection.Close()
        '        End If
        '    End Try

        '    ml.done()

        '    Return 0
        'End Function

    End Class

    Public Class AssignedUser : Inherits AssignmentFlags

        Public aID As Guid = Nothing
        Public SID As String = Nothing
        Public Name As String = Nothing
        Public DisplayName As String = Nothing
        Public Department As String = Nothing
        Public Company As String = Nothing
        Public UPN As String = Nothing
        Public UAC As Integer = Nothing
        Public DN As String = Nothing
        Public Info As String = Nothing
        Public Description As String = Nothing
        Public TimeStamp As DateTime = Nothing
        Public pwdLastSet As Long = Nothing
        Public Status As Integer = 0
        Public source As ASS_SOURCE
        Public lockoutTime As Long = 0

        <JsonIgnore>
        Public User As User = Nothing
        <JsonIgnore>
        Friend memberOf As New List(Of Object)
        <JsonIgnore>
        Public _Assignments As New List(Of Assignment)

        Public ReadOnly Property MinGroupTTL As Integer
            Get

                Dim _return As Integer = -1

                If memberOf IsNot Nothing Then

                    For Each _member In memberOf

                        If _member.ToString.ToUpper.IndexOf("TTL") >= 0 Then

                            Dim lObjMatch As Match =
                                Regex.Match(_member, "^<TTL=([0-9]+)>,.*$")

                            If lObjMatch.Success Then

                                If ((lObjMatch.Groups(1).Value > 0 And _return = -1) Or (lObjMatch.Groups(1).Value < _return)) Then

                                    _return = lObjMatch.Groups(1).Value
                                End If
                            End If
                        End If
                    Next
                End If

                Return _return
            End Get
        End Property

        Public ReadOnly Property MaxGroupTTL As Integer
            Get

                Dim _return As Integer = -1

                If memberOf IsNot Nothing Then

                    For Each _member In memberOf

                        If _member.ToString.ToUpper.IndexOf("TTL") >= 0 Then

                            Dim lObjMatch As Match =
                                Regex.Match(_member, "^<TTL=([0-9]+)>,.*$")

                            If lObjMatch.Success Then

                                If ((lObjMatch.Groups(1).Value > 0 And _return = -1) Or (lObjMatch.Groups(1).Value > _return)) Then

                                    _return = lObjMatch.Groups(1).Value
                                End If
                            End If
                        End If
                    Next
                End If

                Return _return
            End Get
        End Property

        Public ReadOnly Property AssignedGroups As List(Of AssignedGroup)
            Get

                Return AssignedGroups(Nothing)
            End Get
        End Property

        Public ReadOnly Property AssignedGroups(Optional _ldap As Ldap = Nothing) As List(Of AssignedGroup)
            Get

                Dim _AssignedGroups As New List(Of AssignedGroup)

                For Each _Assignment In _Assignments.FindAll(Function(e) e.IsGroup And e.uSID = SID)

                    _AssignedGroups.Add(New AssignedGroup With {
                                .aID = _Assignment.aID,
                                .SID = _Assignment.aSID,
                                .flags = _Assignment.flags,
                                .source = _Assignment.source,
                                .AssignedUser = Me
                            })
                Next

                If _ldap IsNot Nothing Then

                    _ldap.LoadLdapGroup(_AssignedGroups)

                ElseIf User IsNot Nothing Then

                    User.ldap.LoadLdapGroup(_AssignedGroups)
                End If

                Return _AssignedGroups
            End Get
        End Property

        Public ReadOnly Property AccountStatus As Long
            Get

                Dim lLngResult As Long = 0

                'ACCOUNTDISABLE	0x0002	2
                If UAC.arsIsBitSet(2) Then

                    lLngResult = lLngResult Or 2
                End If

                'LOCKOUT	lockoutTime > 0
                If lockoutTime > 0 Then

                    lLngResult = lLngResult Or 16
                End If

                'NOT_DELEGATED	0x100000	1048576
                If UAC.arsIsBitSet(1048576) Then

                    lLngResult = lLngResult Or 1048576
                End If

                'PASSWORD_EXPIRED	0x800000	8388608
                If UAC.arsIsBitSet(8388608) Then

                    lLngResult = lLngResult Or 8388608
                End If

                Return lLngResult
            End Get
        End Property
    End Class

    Public Class AssignedGroup

        Public aID As Guid = Nothing
        Public SID As String = Nothing
        Public rSID As String = Nothing
        Public Name As String = Nothing
        Public Info As String = Nothing
        Public Description As String = Nothing
        Public DN As String = Nothing
        Public flags As Integer = 0
        Public managedBy As String
        Public managedObjects As List(Of Object)
        Public source As ASS_SOURCE

        <JsonIgnore>
        Public User As User
        <JsonIgnore>
        Public AssignedUser As AssignedUser
        <JsonIgnore>
        Public TimeStamp As DateTime = Nothing
        <JsonIgnore>
        Private _flag As Integer = 0
        <JsonIgnore>
        Private _ttl As Integer = 0
        <JsonIgnore>
        Private _exp As Integer = 0
        <JsonIgnore>
        Private _init As Boolean = False

        Private Sub getTtlAndExp()

            If AssignedUser.memberOf IsNot Nothing Then

                Dim lStrMembership As String =
                    AssignedUser.memberOf.FirstOrDefault(
                        Function(n) n.IndexOf(DN) >= 0)

                If lStrMembership IsNot Nothing Then

                    If lStrMembership.StartsWith("<TTL=") Then

                        Dim lObjMatch As Match =
                            Regex.Match(lStrMembership, "^<TTL=([0-9]+)>,.*$")

                        If lObjMatch.Success Then

                            _ttl = lObjMatch.Groups(1).Value
                            _exp = (DateAdd(DateInterval.Second, _ttl, AssignedUser.TimeStamp) - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
                        End If
                    Else

                        _ttl = -1
                        _exp = 0
                    End If
                End If
            End If

            _init = True
        End Sub

        Public ReadOnly Property IsRoleAssignment As Integer
            Get

                If String.IsNullOrEmpty(rSID) Then

                    Return False
                End If

                Return True
            End Get
        End Property

        Public ReadOnly Property TTL As Integer
            Get

                If Not _init Then

                    getTtlAndExp()
                End If

                Return _ttl
            End Get
        End Property

        Public ReadOnly Property Expire As Long
            Get

                If Not _init Then

                    getTtlAndExp()
                End If

                Return _exp
            End Get
        End Property

        Public ReadOnly Property Type As ArsGroupType
            Get

                If flags.arsIsBitSet(ASS_FLAG.GROUP_ASSIGNMENT_BASE) Then

                    Return ArsGroupType.BASE
                End If

                Return ArsGroupType.OTIONAL
            End Get
        End Property

        Public ReadOnly Property TypeString As String
            Get

                If flags.arsIsBitSet(ASS_FLAG.GROUP_ASSIGNMENT_BASE) Then

                    Return "base"
                End If

                Return "optional"
            End Get
        End Property

        Public ReadOnly Property Users() As List(Of AssignedUser)
            Get

                Dim lLstUsers As New List(Of AssignedUser)

                Try

                    If SID.Length > 0 Then

                        If Not User.SqlConnection.State = ConnectionState.Open Then

                            User.SqlConnection.Open()
                        End If

                        Dim lObjSqlCommand As New SqlCommand(
                            "SELECT [aID], [uSID], [aSID], [flag] FROM [ars_assignments] WHERE aSID = @aSID",
                            User.SqlConnection)

                        lObjSqlCommand.Parameters.AddWithValue("aSID", SID)

                        Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

                        While lObjReader.Read()
                            lLstUsers.Add(New AssignedUser With {
                                .aID = lObjReader.GetGuid(0),
                                .SID = lObjReader.GetString(1),
                                .flags = lObjReader.GetInt32(3)
                            })
                        End While

                        lObjReader.Close()
                        lObjSqlCommand.Dispose()
                        User.SqlConnection.Close()

                        User.ldap.LoadLdapUser(lLstUsers)
                    End If

                Catch ex As Exception

                    If Not User.SqlConnection.State = ConnectionState.Closed Then

                        User.SqlConnection.Close()
                    End If
                End Try

                Return lLstUsers
            End Get
        End Property

    End Class

    Class AssignmentRequest

        Public SID As String = Nothing
        Public Type As String = Nothing
        Public nType As Integer = 0
        Public State As Integer = 0

        Public Sub New(_SID As String, _Type As Integer)

            SID = _SID
            Type = _Type
            Int32.TryParse(Type, nType)
        End Sub

        Public ReadOnly Property IsValid() As Boolean
            Get

                If Not (New SecurityIdentifier(SID)).IsAccountSid Then

                    Return False
                End If

                If Not Int32.TryParse(Type, nType) And Not (nType > 0 And nType < 4) Then

                    Return False
                End If

                Return True
            End Get
        End Property
    End Class
End Namespace

