Imports System.Data.SqlClient
Imports System.Security.Claims
Imports System.Runtime.CompilerServices
Imports System.DirectoryServices.Protocols
Imports System.Net
Imports System.Security.Principal
Imports Newtonsoft.Json
Imports System.Reflection
Imports System.Web.Security
Imports Ars

Public Class ArsUser

    Public AssignmentTimes() As Integer = {1, 2, 4, 8}

    Public Enum ArsGroupType
        BASE = 1
        OTIONAL = 2
    End Enum

    Public Class cArsAssignedGroup
        Public aID As Guid = Nothing
        <JsonIgnore>
        Public SID As String = Nothing
        Public Type As ArsGroupType = Nothing
        <JsonIgnore>
        Public IsGroupAdmin = 0
        Public TTL As Integer = Nothing
        Public Name As String = Nothing
        <JsonIgnore>
        Public Info As String = Nothing
        Public Description As String = Nothing
        <JsonIgnore>
        Public distinguishedName As String = Nothing
        '<JsonIgnore>
        'Public TimeStamp As DateTime
        Public TimeStamp As Long
    End Class

    Public Class cArsAssignedUser
        Public aID As Guid = Nothing
        <JsonIgnore>
        Public SID As String = Nothing
        Public Name As String = Nothing
        Public DisplayName As String = Nothing
        Public UserPrincipalName As String = Nothing
        Public Info As String = Nothing
        Public Description As String = Nothing
        Public AccountDisabled As Integer = Nothing
        Public AccountLocked As Integer = Nothing
        Public Flags As Integer = Nothing

        <JsonIgnore>
        Public memberOf As New List(Of Object)
        <JsonIgnore>
        Public userAccountControl As Integer = Nothing
        <JsonIgnore>
        Public distinguishedName As String = Nothing
        Public Groups As New List(Of cArsAssignedGroup)
        <JsonIgnore>
        Public TimeStamp As DateTime
    End Class

    Public Class cArsUser

        <JsonIgnore>
        Public SID As String = Nothing
        Public ars As String = Nothing
        Public Name As String = Nothing
        Public DisplayName As String = Nothing
        Public Department As String = Nothing
        Public Company As String = Nothing
        Public UserPrincipalName As String = Nothing
        Public Options As Integer = 0
        <JsonIgnore>
        Public userAccountControl As Integer = Nothing
        <JsonIgnore>
        Public distinguishedName As String = Nothing
        <JsonIgnore>
        Public IsAllowed As Boolean = False
        <JsonIgnore>
        Public IsAdmin As Boolean = False
        Public Users As New List(Of cArsAssignedUser)
        <JsonIgnore>
        Public UsrTimeStamp As DateTime = DateAdd(DateInterval.Day, -1, Now())                'Logged on User Details 
        <JsonIgnore>
        Public SqlTimeStamp As DateTime = DateAdd(DateInterval.Day, -1, Now())                'SQL Base Infos, Permissions and Assignments
        <JsonIgnore>
        Public AssTimeStamp As DateTime = DateAdd(DateInterval.Day, -1, Now())                'AD Assignment Details, User and Group Details
        <JsonIgnore>
        Public InfoLoaded As Integer = Nothing
    End Class

    Public Class vArsUser0

        Public ars As String = Nothing
        Public Name As String = Nothing
        Public DisplayName As String = Nothing
        Public Department As String = Nothing
        Public Company As String = Nothing
        Public UserPrincipalName As String = Nothing

        Public Sub New(_ArsUser As cArsUser)

            ars = _ArsUser.ars
            Name = _ArsUser.Name
            DisplayName = _ArsUser.DisplayName
            Department = _ArsUser.Department
            Company = _ArsUser.Company
            UserPrincipalName = _ArsUser.UserPrincipalName
        End Sub
    End Class

    Public Class vArsUser1
        'USER INFO WITH ASSIGNED USERS BUT WITHOUT GROUPS
        Public ars As String = Nothing
        Public Name As String = Nothing
        Public DisplayName As String = Nothing
        Public Department As String = Nothing
        Public Company As String = Nothing
        Public UserPrincipalName As String = Nothing
        Public Users As New List(Of cArsAssignedUser0)

        Public Sub New(_ArsUser As cArsUser)

            ars = _ArsUser.ars
            Name = _ArsUser.Name
            DisplayName = _ArsUser.DisplayName
            Department = _ArsUser.Department
            Company = _ArsUser.Company
            UserPrincipalName = _ArsUser.UserPrincipalName

            For Each lObjUser In _ArsUser.Users

                Users.Add(New cArsAssignedUser0(lObjUser))
            Next
        End Sub
    End Class

    Public Class cArsAssignedUser0

        Public Name As String = Nothing
        Public aID As Guid = Nothing
        Public DisplayName As String = Nothing
        Public Description As String = Nothing
        Public Info As String = Nothing
        Public AccountDisabled As Integer = Nothing
        Public AccountLocked As Integer = Nothing
        Public UserPrincipalName As String = Nothing
        Public Flags As Integer = 0

        Public Sub New(_ArsAssignedUser As cArsAssignedUser)

            aID = _ArsAssignedUser.aID
            Name = _ArsAssignedUser.Name
            DisplayName = _ArsAssignedUser.DisplayName
            Description = _ArsAssignedUser.Description
            Info = _ArsAssignedUser.Info
            AccountDisabled = _ArsAssignedUser.AccountDisabled
            AccountLocked = _ArsAssignedUser.AccountLocked
            UserPrincipalName = _ArsAssignedUser.UserPrincipalName
            Flags = _ArsAssignedUser.Flags
        End Sub
    End Class

    Public Class cArsAssignmentInfo

        Public aID As String = Nothing
        Public uSID As String = Nothing
        Public gSID As String = Nothing
        Public uDn As String = Nothing
        Public gDn As String = Nothing
        Public gType As Integer = -1
    End Class

    Public SID As String = Nothing

    Public arsDebug As String = Nothing
    Private _arsDebugLevel As Integer = 0

    Public LastError As String = Nothing
    Public CurrentPage As String = ""

    Private _SearchResultLimit As Integer = 8
    Private _DbgInitTime As DateTime
    Private _ArsUser As New cArsUser
    Private gObjPrincipal As System.Security.Principal.IPrincipal

    Private gObjSqlConnection As New SqlConnection(
                ConfigurationManager.AppSettings.Item("sqlConnectionString"))

    Public LdapServer As String = Nothing
    Public LdapPort As Integer = 389
    Private IsPersistentLdapServer As Boolean = False

    Public SqlCacheSeconds As Integer = 30
    Public UsrCacheSeconds As Integer = 30
    Public AssCacheSeconds As Integer = 60
    Public LastNoCacheRequest As DateTime = Now()

    Public GroupSearchBase As String = Nothing
    Public UserSearchBase As String = Nothing

    Private pwdLength As Integer = 8
    Private pwdSpecialCharacters As Integer = 2
    Private adObjectCacheSeconds As Integer = 1440

    Public HttpContext As HttpContext = Nothing

    Public Sub New(lObjPrincipal As System.Security.Principal.IPrincipal, lStrPage As String)

        If ConfigurationManager.AppSettings.AllKeys.Contains("ldap_ObjectCacheSeconds") Then

            If IsNumeric(ConfigurationManager.AppSettings.Item("ldap_ObjectCacheSeconds")) Then

                If ConfigurationManager.AppSettings.Item("ldap_ObjectCacheSeconds") > adObjectCacheSeconds Then

                    adObjectCacheSeconds = ConfigurationManager.AppSettings.Item("ldap_ObjectCacheSeconds")
                End If
            End If
        End If

        If ConfigurationManager.AppSettings.AllKeys.Contains("pwd_Length") Then

            If IsNumeric(ConfigurationManager.AppSettings.Item("pwd_Length")) Then

                If ConfigurationManager.AppSettings.Item("pwd_Length") > pwdLength Then

                    pwdLength = ConfigurationManager.AppSettings.Item("pwd_Length")
                End If
            End If
        End If

        If ConfigurationManager.AppSettings.AllKeys.Contains("pwd_SpecialCharacters") Then

            If IsNumeric(ConfigurationManager.AppSettings.Item("pwd_SpecialCharacters")) Then

                If ConfigurationManager.AppSettings.Item("pwd_SpecialCharacters") > pwdSpecialCharacters Then

                    pwdSpecialCharacters = ConfigurationManager.AppSettings.Item("pwd_SpecialCharacters")
                End If
            End If
        End If

        If ConfigurationManager.AppSettings.AllKeys.Contains("ldap_SearchLimit") Then

            If IsNumeric(ConfigurationManager.AppSettings.Item("ldap_SearchLimit")) Then

                If ConfigurationManager.AppSettings.Item("ldap_SearchLimit") > 0 And ConfigurationManager.AppSettings.Item("ldap_SearchLimit") < 100 Then

                    _SearchResultLimit = ConfigurationManager.AppSettings.Item("ldap_SearchLimit")
                End If
            End If
        End If

        If ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Then

            If IsNumeric(ConfigurationManager.AppSettings.Item("ldap_Port")) Then

                LdapPort = ConfigurationManager.AppSettings.Item("ldap_Port")
            End If
        End If

        If Not ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Or Not ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Or Not ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Then

            LastError = "ARS: Incomplete LDAP Configuration."
            Exit Sub
        Else

            LdapServer = ConfigurationManager.AppSettings.Item("ldap_Server")
            GroupSearchBase = ConfigurationManager.AppSettings.Item("ldap_GroupSearchBase")
            UserSearchBase = ConfigurationManager.AppSettings.Item("ldap_UserSearchBase")

            CurrentPage = lStrPage

            gObjPrincipal = lObjPrincipal
            LoadArsUserInfoFromSql()
        End If
    End Sub

    Public ReadOnly Property ArsUser(Optional lIntLevel As Integer = 1) As cArsUser

        Get
            DebugOut(String.Format("InfoLoaded = {0}, ArsUser({1})", _ArsUser.InfoLoaded, lIntLevel), 1)

            Dim _DbgInitTime As DateTime = Now()

            If (lIntLevel And 128) = 128 Then

                If DateDiff(DateInterval.Second, LastNoCacheRequest, _DbgInitTime) < 15 Then

                    lIntLevel = lIntLevel Xor 128
                Else

                    LastNoCacheRequest = _DbgInitTime
                End If
            End If

            _ArsUser.ars = arsDebug

            Dim lObjTimeSpan As TimeSpan = Nothing

            If (lIntLevel And 1) = 1 Then

                lObjTimeSpan = Now() - _ArsUser.SqlTimeStamp
                DebugOut(String.Format("SqlAge = {0}/{1}", lObjTimeSpan.TotalSeconds, SqlCacheSeconds), 1)

                If lObjTimeSpan.TotalSeconds >= SqlCacheSeconds Or (_ArsUser.InfoLoaded And 1) = 0 Or (lIntLevel And 128) = 128 Then

                    LoadArsUserInfoFromSql()
                End If
            End If

            If (lIntLevel And 2) = 2 Then

                lObjTimeSpan = Now() - _ArsUser.UsrTimeStamp
                DebugOut(String.Format("UsrAge = {0}/{1}", lObjTimeSpan.TotalSeconds, UsrCacheSeconds), 1)

                If lObjTimeSpan.TotalSeconds >= UsrCacheSeconds Or (_ArsUser.InfoLoaded And 2) = 0 Or (lIntLevel And 128) = 128 Then

                    LoadArsUserBaseInfoFromAd()
                End If
            End If

            If (lIntLevel And 4) = 4 Then

                lObjTimeSpan = Now() - _ArsUser.AssTimeStamp
                DebugOut(String.Format("AssAge = {0}/{1}", lObjTimeSpan.TotalSeconds, AssCacheSeconds), 1)

                If lObjTimeSpan.TotalSeconds >= AssCacheSeconds Or (_ArsUser.InfoLoaded And 4) = 0 Or (lIntLevel And 128) = 128 Then

                    LoadArsUserAssignmentInfoFromAd()
                End If
            End If

            DebugOut(String.Format("ArsUser({0}) done.", lIntLevel), 8, lObjDateTime:=_DbgInitTime)

            Return _ArsUser
        End Get
    End Property

    Public Function getAssignment(aID As String) As cArsAssignmentInfo

        Dim _DbgInitTime As DateTime = Now()

        Dim lStrUser As String = Nothing
        Dim lStrGroup As String = Nothing
        Dim lObjReturnInfo As New cArsAssignmentInfo

        Dim lObjSqlCommand As New SqlCommand(
            "SELECT TOP(1) * FROM [ars_assignments] WHERE [aID] = @aID AND ( flag & 2 ) = 2",
            gObjSqlConnection)

        lObjSqlCommand.Parameters.AddWithValue("aID", aID)

        gObjSqlConnection.Open()

        Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

        While lObjReader.Read()

            lStrUser = lObjReader.GetString(1)
            lStrGroup = lObjReader.GetString(2)
        End While

        If lStrUser IsNot Nothing And lStrGroup IsNot Nothing Then

            Dim _tmpArsUser As cArsAssignedUser =
                _ArsUser.Users.Find(Function(p) p.SID = lStrUser)

            If _tmpArsUser IsNot Nothing Then

                Dim _tmpArsGroup As cArsAssignedGroup =
                _tmpArsUser.Groups.Find(Function(p) p.SID = lStrGroup)

                If _tmpArsGroup IsNot Nothing Then

                    lObjReturnInfo.aID = aID
                    lObjReturnInfo.uSID = _tmpArsUser.SID
                    lObjReturnInfo.gSID = _tmpArsGroup.SID
                    lObjReturnInfo.uDn = _tmpArsUser.distinguishedName
                    lObjReturnInfo.gDn = _tmpArsGroup.distinguishedName
                    lObjReturnInfo.gType = _tmpArsGroup.Type
                End If
            End If
        End If

        gObjSqlConnection.Close()

        DebugOut("getAssignment", 8, lObjDateTime:=_DbgInitTime)

        Return lObjReturnInfo

    End Function

    Private Sub LoadArsUserInfoFromSql()

        Dim _RefreshRequired As Boolean = False
        Dim _DbgInitTime As DateTime = Now()

        Try
            Dim lObjClaimsPrincipal As ClaimsPrincipal = gObjPrincipal

            Dim lObjClaimsIdentity =
                lObjClaimsPrincipal.Identities.Where(
                    Function(c) c.GetType() = GetType(System.Security.Claims.ClaimsIdentity))(0)

            Dim lObjSidClaim As Claim =
                lObjClaimsIdentity.Claims.Where(Function(n) n.Type = "onprem_sid")(0)

            If lObjSidClaim.Value.Length > 0 Then

                Dim lObjSqlCommand As New SqlCommand(
                    "SELECT TOP(1) * FROM [ars_users] WHERE [uSID] = @uSID",
                    gObjSqlConnection)

                SID = lObjSidClaim.Value
                _ArsUser.SID = lObjSidClaim.Value

                lObjSqlCommand.Parameters.AddWithValue("uSID", _ArsUser.SID)

                gObjSqlConnection.Open()

                Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

                While lObjReader.Read()

                    Dim lIntFlags As Integer = lObjReader.GetInt32(1)

                    If lIntFlags >= 0 Then

                        _ArsUser.SID = lObjReader.GetString(0)

                        'USER IS ALLOWED
                        If lIntFlags.arsIsBitSet(1) Then

                            _ArsUser.IsAllowed = True
                        End If

                        'USER IS ADMIN
                        If lIntFlags.arsIsBitSet(256) Then

                            _ArsUser.IsAdmin = True
                        End If
                    End If
                End While

                lObjReader.Close()
                lObjSqlCommand.Dispose()

                'GET USER GROUP ASSIGNMENTS
                If _ArsUser.IsAllowed Then

                    Dim lLstUsers As New List(Of cArsAssignedUser)

                    'USER ASSIGNMENT
                    Dim lObjSqlUserCommand As New SqlCommand(
                            "SELECT [aID], [aSID], [flag] FROM [ars_assignments] WHERE [uSID] = @uSID AND ( [flag] & 1 ) = 1",
                            gObjSqlConnection)

                    lObjSqlUserCommand.Parameters.AddWithValue("uSID", _ArsUser.SID)

                    Dim lObjSqlUserReader As SqlDataReader = lObjSqlUserCommand.ExecuteReader()

                    While lObjSqlUserReader.Read()

                        If lObjSqlUserReader.GetInt32(2).arsIsBitSet(1) Then

                            lLstUsers.Add(New cArsAssignedUser With {
                                    .aID = lObjSqlUserReader.GetGuid(0),
                                    .SID = lObjSqlUserReader.GetString(1),
                                    .Flags = lObjSqlUserReader.GetInt32(2)
                                })
                        End If
                    End While

                    lObjSqlUserReader.Close()
                    lObjSqlCommand.Dispose()

                    If lLstUsers.Count = _ArsUser.Users.Count Then

                        For Each lObjUser In lLstUsers

                            Dim _lObjArsAssignedUser As cArsAssignedUser =
                                _ArsUser.Users.Find(Function(i) i.SID = lObjUser.SID)

                            If _lObjArsAssignedUser IsNot Nothing Then

                                _lObjArsAssignedUser.Flags = lObjUser.Flags
                            Else

                                _ArsUser.Users = lLstUsers
                                _RefreshRequired = True
                                Exit For
                            End If
                        Next
                    Else

                        _ArsUser.Users = lLstUsers
                        _RefreshRequired = True
                    End If

                    'GROUP ASSIGNMENTS
                    Dim lObjSqlGroupCommand As New SqlCommand(
                                                "SELECT [aID], [aSID], [flag] FROM [ars_assignments] WHERE [uSID] = @uSID AND ( flag & 2 ) = 2",
                                                gObjSqlConnection)

                    For Each lObjUser In _ArsUser.Users

                        Dim lLstGroups As New List(Of cArsAssignedGroup)

                        lObjSqlGroupCommand.Parameters.Clear()
                        lObjSqlGroupCommand.Parameters.AddWithValue("uSID", lObjUser.SID)

                        Dim lObjSqlGroupReader As SqlDataReader = lObjSqlGroupCommand.ExecuteReader()

                        While lObjSqlGroupReader.Read()

                            Dim lIntType As ArsGroupType = ArsGroupType.OTIONAL

                            'Optional Group Assignment Type
                            If lObjSqlGroupReader.GetInt32(2).arsIsBitSet(4) Then

                                lIntType = ArsGroupType.BASE
                            End If

                            lLstGroups.Add(New cArsAssignedGroup With {
                                    .aID = lObjSqlGroupReader.GetGuid(0),
                                    .SID = lObjSqlGroupReader.GetString(1),
                                    .IsGroupAdmin = lObjSqlGroupReader.GetInt32(2).arsIsBitSet(256),
                                    .Type = lIntType
                                })

                        End While

                        If lLstGroups.Count <> lObjUser.Groups.Count Or _RefreshRequired Then

                            lObjUser.Groups = lLstGroups
                        Else

                            For Each lObjGroup In lLstGroups

                                Dim _lObjArsAssignedGroup As cArsAssignedGroup =
                                    lObjUser.Groups.Find(Function(i) i.SID = lObjGroup.SID)

                                If lObjUser.Groups.Find(Function(i) i.SID = lObjGroup.SID) Is Nothing Then

                                    lObjUser.Groups = lLstGroups
                                    _RefreshRequired = True
                                    Exit For
                                End If
                            Next
                        End If

                        lObjSqlGroupReader.Close()
                    Next

                    lObjSqlGroupCommand.Dispose()

                End If

                gObjSqlConnection.Close()

                'If ArsUserChanged(lObjNewArsUser) Then

                '_ArsUser = lObjNewArsUser
                '_ArsUser.InfoLoaded = 0
                'End If

                If _RefreshRequired = True Then

                    _ArsUser.InfoLoaded = 0
                End If

                _ArsUser.SqlTimeStamp = Now()
                _ArsUser.InfoLoaded = _ArsUser.InfoLoaded Or 1

            End If
        Catch ex As Exception

            DebugOut("LoadArsUserInfoFromSql", 16, lStrError:=ex.StackTrace)
        End Try

        DebugOut("LoadArsUserInfoFromSql", 8, lObjDateTime:=_DbgInitTime)

    End Sub

    Private Sub LoadArsUserBaseInfoFromAd()

        Dim _DbgInitTime As DateTime = Now()

        Dim lObjLdapConn As LdapConnection = getLdapConnection()

        If lObjLdapConn IsNot Nothing Then

            Try

                Dim lLstAttributes As String() =
                    {"name", "description", "info", "distinguishedname", "company", "department", "displayname", "UserPrincipalName"}

                Dim lStrFilter As String =
                    String.Format("(objectSid={0})", _ArsUser.SID)

                Dim lObjSearchRequest As New SearchRequest(
                    UserSearchBase,
                    lStrFilter,
                    System.DirectoryServices.Protocols.SearchScope.Subtree,
                    lLstAttributes)

                lObjSearchRequest.SizeLimit = 1

                Dim lObjResponse As SearchResponse =
                    lObjLdapConn.SendRequest(lObjSearchRequest)

                If lObjResponse.Entries.Count > 0 Then

                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                        If lObjEntry.Attributes.Contains("name") Then

                            _ArsUser.Name = lObjEntry.Attributes("name")(0)
                        Else

                            _ArsUser.Name = vbNullString
                        End If

                        If lObjEntry.Attributes.Contains("distinguishedname") Then

                            _ArsUser.distinguishedName = lObjEntry.Attributes("distinguishedname")(0)
                        Else

                            _ArsUser.distinguishedName = vbNullString
                        End If

                        If lObjEntry.Attributes.Contains("department") Then

                            _ArsUser.Department = lObjEntry.Attributes("department")(0)
                        Else

                            _ArsUser.Department = vbNullString
                        End If

                        If lObjEntry.Attributes.Contains("displayname") Then

                            _ArsUser.DisplayName = lObjEntry.Attributes("displayname")(0)
                        Else

                            _ArsUser.DisplayName = vbNullString
                        End If

                        If lObjEntry.Attributes.Contains("company") Then

                            _ArsUser.Company = lObjEntry.Attributes("company")(0)
                        Else

                            _ArsUser.Company = vbNullString
                        End If

                        If lObjEntry.Attributes.Contains("UserPrincipalName") Then

                            _ArsUser.UserPrincipalName = lObjEntry.Attributes("UserPrincipalName")(0)
                        Else

                            _ArsUser.UserPrincipalName = vbNullString
                        End If
                    Next
                End If

                lObjLdapConn.Dispose()

                _ArsUser.UsrTimeStamp = Now()
                _ArsUser.InfoLoaded = _ArsUser.InfoLoaded Or 2

            Catch ex As Exception

                DebugOut("LoadArsUserBaseInfoFromAd", 16, lStrError:=ex.StackTrace)
            End Try
        End If

        DebugOut("LoadArsUserBaseInfoFromAd", 8, lObjDateTime:=_DbgInitTime)

    End Sub

    Private Sub LoadArsUserAssignmentInfoFromAd()

        Dim _DbgInitTime As DateTime = Now()

        Dim lObjLdapConn As LdapConnection = getLdapConnection()

        If lObjLdapConn IsNot Nothing Then

            Try

                Dim lStrFilter As String = Nothing
                Dim lLstUserAttributes As String() =
                    {"memberOf", "name", "description", "distinguishedname", "displayname", "objectSid", "userAccountControl", "UserPrincipalName"}
                Dim lObjTtlControl As New DirectoryControl(
                "1.2.840.113556.1.4.2309", Nothing, True, True)

                Dim lObjSearchRequest As New SearchRequest(
                    UserSearchBase,
                    lStrFilter,
                    System.DirectoryServices.Protocols.SearchScope.Subtree,
                    lLstUserAttributes)

                lObjSearchRequest.Controls.Add(lObjTtlControl)

                Dim lObjResponse As SearchResponse = Nothing

                If _ArsUser.Users.Count > 0 Then

                    lStrFilter = "(&(objectClass=user)(|"

                    For u = 0 To _ArsUser.Users.Count - 1

                        lStrFilter += String.Format("(objectSid={0})", _ArsUser.Users(u).SID)
                    Next

                    lStrFilter += "))"

                    lObjSearchRequest.Filter = lStrFilter
                    lObjResponse = lObjLdapConn.SendRequest(lObjSearchRequest)

                    If lObjResponse.Entries.Count > 0 Then

                        For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                            For u = 0 To _ArsUser.Users.Count - 1

                                Dim lObjTmpSid As New SecurityIdentifier(lObjEntry.Attributes("objectSid")(0), 0)

                                If lObjTmpSid.ToString() = _ArsUser.Users(u).SID Then

                                    'empty attributes ar not exist in attribute collection
                                    'require to check existence
                                    If lObjEntry.Attributes.Contains("name") Then

                                        _ArsUser.Users(u).Name = lObjEntry.Attributes("name")(0)
                                    Else

                                        _ArsUser.Users(u).Name = vbNullString
                                    End If

                                    If lObjEntry.Attributes.Contains("displayname") Then

                                        _ArsUser.Users(u).DisplayName = lObjEntry.Attributes("displayname")(0)
                                    Else

                                        _ArsUser.Users(u).DisplayName = vbNullString
                                    End If

                                    If lObjEntry.Attributes.Contains("description") Then

                                        _ArsUser.Users(u).Description = lObjEntry.Attributes("description")(0)
                                    Else

                                        _ArsUser.Users(u).Description = vbNullString
                                    End If

                                    If lObjEntry.Attributes.Contains("info") Then

                                        _ArsUser.Users(u).Info = lObjEntry.Attributes("info")(0)
                                    Else

                                        _ArsUser.Users(u).Info = vbNullString
                                    End If

                                    If lObjEntry.Attributes.Contains("distinguishedname") Then

                                        _ArsUser.Users(u).distinguishedName = lObjEntry.Attributes("distinguishedname")(0)
                                    Else

                                        _ArsUser.Users(u).distinguishedName = vbNullString
                                    End If

                                    If lObjEntry.Attributes.Contains("UserPrincipalName") Then

                                        _ArsUser.Users(u).UserPrincipalName = lObjEntry.Attributes("UserPrincipalName")(0)
                                    Else

                                        _ArsUser.Users(u).UserPrincipalName = vbNullString
                                    End If

                                    If lObjEntry.Attributes.Contains("userAccountControl") Then

                                        Int32.TryParse(lObjEntry.Attributes("userAccountControl")(0), _ArsUser.Users(u).userAccountControl)
                                    Else

                                        _ArsUser.Users(u).userAccountControl = 0
                                    End If

                                    If (_ArsUser.Users(u).userAccountControl And 2) = 2 Then

                                        _ArsUser.Users(u).AccountDisabled = 1
                                    Else

                                        _ArsUser.Users(u).AccountDisabled = 0
                                    End If

                                    If (_ArsUser.Users(u).userAccountControl And 16) = 16 Then

                                        _ArsUser.Users(u).AccountLocked = 1
                                    Else

                                        _ArsUser.Users(u).AccountLocked = 0
                                    End If

                                    If lObjEntry.Attributes.Contains("memberof") Then

                                        _ArsUser.Users(u).memberOf =
                                            lObjEntry.Attributes("memberof").GetValues(GetType(String)).ToList()
                                    End If

                                End If
                            Next
                        Next
                    End If

                    'GET GROUP INFO
                    For u = 0 To _ArsUser.Users.Count - 1

                        If _ArsUser.Users(u).Groups.Count > 0 Then



                            lObjSearchRequest.Attributes.Clear()
                            lObjSearchRequest.Attributes.AddRange(
                                {"name", "description", "info", "distinguishedname", "objectSid"})

                            For g = 0 To _ArsUser.Users(u).Groups.Count - 1

                                Dim lobjCachedGroupInfo As Ars.CachedGroup =
                                    HttpContext.Cache.Get(_ArsUser.Users(u).Groups(g).SID)

                                If lobjCachedGroupInfo IsNot Nothing Then

                                    _ArsUser.Users(u).Groups(g).Name = lobjCachedGroupInfo.Name & "_"
                                    _ArsUser.Users(u).Groups(g).Description = lobjCachedGroupInfo.Description
                                    _ArsUser.Users(u).Groups(g).Info = lobjCachedGroupInfo.info
                                    _ArsUser.Users(u).Groups(g).distinguishedName = lobjCachedGroupInfo.dn
                                    ArsUser.Users(u).Groups(g).TTL = getMembershipTTL(_ArsUser.Users(u), _ArsUser.Users(u).Groups(g).distinguishedName)
                                Else

                                    lStrFilter += String.Format("(objectSid={0})", _ArsUser.Users(u).Groups(g).SID)
                                End If
                            Next

                            If Not String.IsNullOrEmpty(lStrFilter) Then

                                lObjSearchRequest.Filter = String.Format("(&(objectClass=group)(|{0}))", lStrFilter)
                                lObjSearchRequest.DistinguishedName = UserSearchBase
                                lObjResponse = lObjLdapConn.SendRequest(lObjSearchRequest)

                                If lObjResponse.Entries.Count > 0 Then

                                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                                        For g = 0 To _ArsUser.Users(u).Groups.Count - 1

                                            Dim lObjTmpSid As String =
                                                New SecurityIdentifier(lObjEntry.Attributes("objectSid")(0), 0).ToString()

                                            If lObjTmpSid = _ArsUser.Users(u).Groups(g).SID Then

                                                Dim tt As New Dictionary(Of String, String)
                                                'empty attributes ar not exist in attribute collection
                                                'require to check existence
                                                If lObjEntry.Attributes.Contains("name") Then

                                                    _ArsUser.Users(u).Groups(g).Name = lObjEntry.Attributes("name")(0)
                                                Else

                                                    _ArsUser.Users(u).Groups(g).Name = vbNullString
                                                End If

                                                If lObjEntry.Attributes.Contains("description") Then

                                                    _ArsUser.Users(u).Groups(g).Description = lObjEntry.Attributes("description")(0)
                                                Else

                                                    _ArsUser.Users(u).Groups(g).Description = vbNullString
                                                End If

                                                If lObjEntry.Attributes.Contains("info") Then

                                                    _ArsUser.Users(u).Groups(g).Info = lObjEntry.Attributes("info")(0)
                                                Else

                                                    _ArsUser.Users(u).Groups(g).Info = vbNullString
                                                End If

                                                If lObjEntry.Attributes.Contains("distinguishedname") Then

                                                    _ArsUser.Users(u).Groups(g).distinguishedName = lObjEntry.Attributes("distinguishedname")(0)
                                                Else

                                                    _ArsUser.Users(u).Groups(g).distinguishedName = vbNullString
                                                End If

                                                ArsUser.Users(u).Groups(g).TTL =
                                                    getMembershipTTL(_ArsUser.Users(u), _ArsUser.Users(u).Groups(g).distinguishedName)

                                                'Dim lStrMembership As String =
                                                '    _ArsUser.Users(u).memberOf.Find(
                                                '        Function(n) n.IndexOf(_ArsUser.Users(u).Groups(g).distinguishedName) >= 0)

                                                'If lStrMembership IsNot Nothing Then

                                                '    DebugOut(lStrMembership, 8, lObjDateTime:=_DbgInitTime)

                                                '    If lStrMembership.StartsWith("<TTL=") Then

                                                '        Dim lObjMatch As Match = Regex.Match(lStrMembership, "^<TTL=([0-9]+)>,.*$")

                                                '        If lObjMatch.Success Then

                                                '            _ArsUser.Users(u).Groups(g).TTL = CLng(lObjMatch.Groups(1).Value)
                                                '        End If
                                                '    Else

                                                '        _ArsUser.Users(u).Groups(g).TTL = -1
                                                '    End If
                                                'End If

                                                '_ArsUser.Users(u).Groups(g).TTL = 0

                                                'If lObjEntry.Attributes.Contains("member") Then

                                                '    Dim lStrGroupMemberAttribute As String = "member"

                                                '    For Each lStrlStrAttributeName As String In lObjEntry.Attributes.AttributeNames

                                                '        If lStrlStrAttributeName.ToLower().IndexOf("member;range") >= 0 Then

                                                '            lStrGroupMemberAttribute = lStrlStrAttributeName
                                                '            Exit For
                                                '        End If
                                                '    Next

                                                '    For i = 0 To lObjEntry.Attributes(lStrGroupMemberAttribute).Count - 1

                                                '        If lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString().IndexOf(_ArsUser.Users(u).distinguishedName) >= 0 Then

                                                '            If lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString().StartsWith("<TTL=") Then

                                                '                Dim lObjMatch As Match = Regex.Match(lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString(), "^<TTL=([0-9]+)>,.*$")

                                                '                If lObjMatch.Success Then

                                                '                    _ArsUser.Users(u).Groups(g).TTL = CLng(lObjMatch.Groups(1).Value)
                                                '                End If
                                                '            Else

                                                '                _ArsUser.Users(u).Groups(g).TTL = -1
                                                '            End If

                                                '            Exit For
                                                '        End If
                                                '    Next

                                                '    If lStrGroupMemberAttribute.Contains("range") Then

                                                '        LoadLargeGroupMembership(_ArsUser.Users(u).Groups(g), _ArsUser.Users(u).distinguishedName, lStrGroupMemberAttribute)
                                                '    End If
                                                'End If
                                            End If

                                            _ArsUser.Users(u).Groups(g).TimeStamp = Math.Round(DateTime.UtcNow.Subtract(
                                                New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, 0)
                                        Next
                                    Next
                                End If
                            End If
                        End If
                    Next
                End If

                lObjLdapConn.Dispose()

                _ArsUser.AssTimeStamp = Now()
                _ArsUser.InfoLoaded = _ArsUser.InfoLoaded Or 4

            Catch ex As Exception

                DebugOut("LoadArsUserAssignmentInfoFromAd", 16, lStrError:=ex.StackTrace)
            End Try
        End If

        DebugOut("LoadArsUserAssignmentInfoFromAd", 8, lObjDateTime:=_DbgInitTime)

    End Sub

    Public Function addMemberToGroup(lObjAssignmentInfo As cArsAssignmentInfo, Optional lIntTime As Integer = 0) As Integer

        Try

            If lIntTime > 0 Then

                lObjAssignmentInfo.uDn =
                    String.Format("<TTL={0},{1}>", lIntTime, lObjAssignmentInfo.uDn)
            End If

            Dim lObjLdapConnection As LdapConnection = getLdapConnection()

            Dim lObjDirectoryAttributeModification As New DirectoryAttributeModification()
            lObjDirectoryAttributeModification.Name = "member"
            lObjDirectoryAttributeModification.Add(lObjAssignmentInfo.uDn)
            lObjDirectoryAttributeModification.Operation = DirectoryAttributeOperation.Add

            Dim lObjModifyRequest As New ModifyRequest(
                lObjAssignmentInfo.gDn, lObjDirectoryAttributeModification)

            Dim lObjModifyResponse As ModifyResponse =
                lObjLdapConnection.SendRequest(lObjModifyRequest)

            lObjLdapConnection.Dispose()

            Return lObjModifyResponse.ResultCode

        Catch ex As Exception

            DebugOut("addMemberToGroup", 16, lStrError:=ex.StackTrace)
            Return -1
        End Try

    End Function

    Public Function removeMemberFromGroup(lObjAssignmentInfo As cArsAssignmentInfo) As Integer

        Try

            Dim lObjLdapConnection As LdapConnection = getLdapConnection()

            Dim lObjDirectoryAttributeModification As New DirectoryAttributeModification()
            lObjDirectoryAttributeModification.Name = "member"
            lObjDirectoryAttributeModification.Add(lObjAssignmentInfo.uDn)
            lObjDirectoryAttributeModification.Operation = DirectoryAttributeOperation.Delete

            Dim lObjModifyRequest As New ModifyRequest(
                lObjAssignmentInfo.gDn, lObjDirectoryAttributeModification)

            Dim lObjModifyResponse As ModifyResponse =
                lObjLdapConnection.SendRequest(lObjModifyRequest)

            lObjLdapConnection.Dispose()

            Return lObjModifyResponse.ResultCode

        Catch ex As Exception

            DebugOut("removeMemberFromGroup", 16, lStrError:=ex.StackTrace)
            Return -1
        End Try

    End Function

    Public Function searchAd(lLstrClasses As List(Of String), lStrSearchString As String) As List(Of Dictionary(Of String, String))

        Dim _DbgInitTime As DateTime = Now()
        Dim lDicReturn As New List(Of Dictionary(Of String, String))

        Dim lObjLdapConn As LdapConnection = getLdapConnection()

        If lObjLdapConn IsNot Nothing Then

            Try

                Dim lStrFilter As String = Nothing

                For Each lStrClass In lLstrClasses

                    lStrFilter += String.Format("(objectClass={0})", lStrClass)
                Next

                lStrFilter = String.Format("(&(|{0})(anr={1}*))", lStrFilter, lStrSearchString)

                Dim lLstUserAttributes As String() =
                    {"name", "description", "displayname", "company", "department", "objectSid", "objectClass", "UserPrincipalName"}

                Dim lObjSearchRequest As New SearchRequest(
                    UserSearchBase,
                    lStrFilter,
                    System.DirectoryServices.Protocols.SearchScope.Subtree,
                    lLstUserAttributes)

                lObjSearchRequest.SizeLimit = _SearchResultLimit

                Dim lObjResponse As SearchResponse =
                    lObjLdapConn.SendRequest(lObjSearchRequest)

                If lObjResponse.Entries.Count > 0 Then

                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                        Dim lObjResultEntry As New Dictionary(Of String, String) From {
                                {"sid", New SecurityIdentifier(lObjEntry.Attributes("objectSid")(0), 0).ToString()},
                                {"description", ""},
                                {"type", "group"},
                                {"name", lObjEntry.Attributes("name")(0)}
                            }

                        If lObjEntry.Attributes.Contains("description") Then

                            lObjResultEntry("description") = lObjEntry.Attributes("description")(0)
                        End If

                        If Not daToString(lObjEntry.Attributes("objectclass")).IndexOf("group") >= 0 Then

                            lObjResultEntry("type") = "user"

                            If lObjEntry.Attributes.Contains("displayname") Then

                                lObjResultEntry.Add("displayname", lObjEntry.Attributes("displayname")(0))
                            Else

                                lObjResultEntry.Add("displayname", "")
                            End If

                            If lObjEntry.Attributes.Contains("company") Then

                                lObjResultEntry.Add("company", lObjEntry.Attributes("company")(0))
                            Else

                                lObjResultEntry.Add("company", "")
                            End If

                            If lObjEntry.Attributes.Contains("department") Then

                                lObjResultEntry.Add("department", lObjEntry.Attributes("department")(0))
                            Else

                                lObjResultEntry.Add("department", "")
                            End If

                            If lObjEntry.Attributes.Contains("UserPrincipalName") Then

                                lObjResultEntry.Add("UserPrincipalName", lObjEntry.Attributes("UserPrincipalName")(0))
                            Else

                                lObjResultEntry.Add("UserPrincipalName", "")
                            End If
                        End If

                        lDicReturn.Add(lObjResultEntry)
                    Next
                End If

                lObjLdapConn.Dispose()

            Catch ex As Exception

                DebugOut("searchAd", 16, lStrError:=ex.StackTrace)
            End Try
        End If

        DebugOut("searchAd", 8, lObjDateTime:=_DbgInitTime)

        Return lDicReturn

    End Function

    Private Function daToString(da As DirectoryAttribute) As String

        Dim lStrReturn As New List(Of String)

        For Each de In da

            lStrReturn.Add(System.Text.Encoding.UTF8.GetString(de))
        Next

        Return String.Join(",", lStrReturn)

    End Function

    Private Function getLdapConnection() As LdapConnection

        Try

            Dim lObjCreds As New NetworkCredential()

            Dim lObjLdapIdent As New LdapDirectoryIdentifier(
                LdapServer, LdapPort, True, False)

            Dim lObjLdapConn As New LdapConnection(
                lObjLdapIdent,
                lObjCreds,
                DirectoryServices.Protocols.AuthType.Negotiate)

            Try

                lObjLdapConn.Bind()
            Catch ex As Exception

                If IsPersistentLdapServer Then

                    IsPersistentLdapServer = False
                    LdapServer = ConfigurationManager.AppSettings.Item("ldap_Server")
                    lObjLdapConn = getLdapConnection()
                End If
            End Try

            If Not IsPersistentLdapServer Then

                getPersistentLdapServer(lObjLdapConn)
            End If

            Return lObjLdapConn
        Catch ex As Exception
            '
            DebugOut("getLdapConnection", 16, lStrError:=ex.StackTrace)
            Return Nothing
        End Try
    End Function

    Private Sub getPersistentLdapServer(ByRef lObjLdapConn As LdapConnection)

        Dim lObjSearchRequest As New SearchRequest(
            "",
            "(&(objectClass=*))",
            System.DirectoryServices.Protocols.SearchScope.Base,
            "dnsHostName")

        lObjSearchRequest.SizeLimit = 1

        Dim lObjResponse As SearchResponse =
            lObjLdapConn.SendRequest(lObjSearchRequest)

        If lObjResponse.Entries.Count > 0 Then

            For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                If lObjEntry.Attributes.Contains("dnshostname") Then

                    LdapServer = lObjEntry.Attributes("dnshostname")(0)
                    IsPersistentLdapServer = True
                End If
            Next
        End If
    End Sub

    Private Function getMembershipTTL(lObjAssignedUser As cArsAssignedUser, lStrDn As String) As Integer

        Dim lStrMembership As String =
            lObjAssignedUser.memberOf.Find(
                Function(n) n.IndexOf(lStrDn) >= 0)

        If lStrMembership IsNot Nothing Then

            If lStrMembership.StartsWith("<TTL=") Then

                Dim lObjMatch As Match = Regex.Match(lStrMembership, "^<TTL=([0-9]+)>,.*$")

                If lObjMatch.Success Then

                    Return CLng(lObjMatch.Groups(1).Value)
                End If
            Else

                Return -1
            End If
        End If

        Return 0

    End Function

    Private Function dnGetParent(lStrDn As String) As String

        Try

            Dim lStrExpression As String = "^[^,]*,(.*)"

            If lStrDn.IndexOf("TTL") >= 0 Then

                lStrExpression = "^[^,]*,[^,]*,(.*)>"
            End If

            Dim lObjMatch As Match =
                Regex.Match(lStrDn, lStrExpression)

            If lObjMatch.Success Then

                Return lObjMatch.Groups(1).Value
            End If
        Catch ex As Exception

            DebugOut("dnGetParent", 16, lStrError:=ex.StackTrace)
        End Try

        Return Nothing
    End Function

    Public Function toggleAutoDisable(lStrAssignmentId As String) As Integer

        Dim lObjAssignedUser As cArsAssignedUser =
            _ArsUser.Users.Find(Function(i) i.aID.ToString() = lStrAssignmentId)

        If lObjAssignedUser IsNot Nothing Then

            If (lObjAssignedUser.Flags And Ars.ASSIGNMENT_FLAGS.ALLOW_AUTO_DISABLE) = 0 Then

                Dim lObjSqlCommand As New SqlCommand(
                                    "UPDATE TOP(1) [ars_assignments] SET flag = flag ^ @OPT WHERE [aID] = @aID",
                                    gObjSqlConnection)

                lObjSqlCommand.Parameters.AddWithValue("aID", lStrAssignmentId)
                lObjSqlCommand.Parameters.AddWithValue("OPT", Ars.ASSIGNMENT_FLAGS.ALLOW_AUTO_DISABLE_ACTIVE)

                Try

                    gObjSqlConnection.Open()
                    lObjSqlCommand.ExecuteNonQuery()
                    gObjSqlConnection.Close()
                    LoadArsUserInfoFromSql()

                    Return 1
                Catch ex As Exception

                    gObjSqlConnection.Close()
                End Try
            End If
        End If

        Return 0

    End Function

    Public Function toggleAutoPassword(lStrAssignmentId As String) As Integer

        Dim lObjAssignedUser As cArsAssignedUser =
            _ArsUser.Users.Find(Function(i) i.aID.ToString() = lStrAssignmentId)

        If lObjAssignedUser IsNot Nothing Then

            If (lObjAssignedUser.Flags And Ars.ASSIGNMENT_FLAGS.ALLOW_AUTO_PASSWORD) = 0 Then

                Dim lObjSqlCommand As New SqlCommand(
                                    "UPDATE TOP(1) [ars_assignments] SET flag = flag ^ @OPT WHERE [aID] = @aID",
                                    gObjSqlConnection)

                lObjSqlCommand.Parameters.AddWithValue("aID", lStrAssignmentId)
                lObjSqlCommand.Parameters.AddWithValue("OPT", Ars.ASSIGNMENT_FLAGS.ALLOW_AUTO_PASSWORD_ACTIVE)

                Try

                    gObjSqlConnection.Open()
                    lObjSqlCommand.ExecuteNonQuery()
                    gObjSqlConnection.Close()
                    LoadArsUserInfoFromSql()

                    Return 1
                Catch ex As Exception

                    gObjSqlConnection.Close()
                End Try
            End If
        End If

        Return 0

    End Function

    Public Function resetPassword(lStrAssignmentId As String) As String

        Dim lObjAssignedUser As cArsAssignedUser =
            _ArsUser.Users.Find(Function(i) i.aID.ToString() = lStrAssignmentId)

        If lObjAssignedUser IsNot Nothing Then

            If (lObjAssignedUser.Flags And Ars.ASSIGNMENT_FLAGS.ALLOW_RESET_PASSWORD) = 0 Then

                'TODO !! RESET PASSWORD FUNCTION
                Return System.Web.Security.Membership.GeneratePassword(pwdLength, pwdSpecialCharacters)
            End If
        End If

        Return Nothing

    End Function

    Private Sub DebugOut(lStrText As String, lIntLevel As Integer, Optional lStrError As String = vbNullString, Optional lObjDateTime As DateTime = Nothing)

        '0x08 = Times
        '0x16 = Error

        If _arsDebugLevel > 0 Then

            If lIntLevel = 1 Then

                Debug.Print(String.Format("{0,-20} {1,-50}", CurrentPage, lStrText))
            ElseIf lIntLevel = 8 Then

                Debug.Print(String.Format("{0,-20} {1,-50}: {2,10}", CurrentPage, lStrText, (Now() - lObjDateTime).TotalMilliseconds))
            ElseIf lIntLevel = 16 Then

                Debug.Print(String.Format("{0,-20} {1,-50}: {2}", CurrentPage, lStrText, lStrError))
            End If
        End If
    End Sub


    Public Sub updateGroupInfo(lObjAssignmentInfo As cArsAssignmentInfo)

        Dim _DbgInitTime As DateTime = Now()

        Dim lObjLdapConn As LdapConnection = getLdapConnection()

        If lObjLdapConn IsNot Nothing Then

            Try

                Dim lIntLoadRange As Integer = 1500
                Dim lBolEndIsNear As Boolean = False

                Dim lLstAttributes As String() = {"memberof"}

                Dim lObjTtlControl As New DirectoryControl(
                "1.2.840.113556.1.4.2309", Nothing, True, True)

                Dim lObjSearchRequest As New SearchRequest(
                    dnGetParent(lObjAssignmentInfo.uDn),
                    String.Format("(&(objectClass=user)(objectSid={0}))", lObjAssignmentInfo.uSID),
                    System.DirectoryServices.Protocols.SearchScope.OneLevel,
                    lLstAttributes)

                Dim lObjResponse As SearchResponse = Nothing

                Dim u As Integer =
                    _ArsUser.Users.FindIndex(Function(p) p.SID = lObjAssignmentInfo.uSID)

                If u >= 0 Then

                    Dim g As Integer =
                        _ArsUser.Users(u).Groups.FindIndex(Function(p) p.SID = lObjAssignmentInfo.gSID)

                    If g >= 0 Then

                        lObjSearchRequest.SizeLimit = 1
                        lObjSearchRequest.Controls.Add(lObjTtlControl)
                        lObjResponse = lObjLdapConn.SendRequest(lObjSearchRequest)

                        For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                            If lObjEntry.Attributes.Contains("memberof") Then

                                _ArsUser.Users(u).memberOf =
                                    lObjEntry.Attributes("memberof").GetValues(GetType(String)).ToList()
                            End If

                            _ArsUser.Users(u).Groups(g).TTL =
                                getMembershipTTL(_ArsUser.Users(u), _ArsUser.Users(u).Groups(g).distinguishedName)
                        Next
                    End If
                End If
            Catch ex As Exception

                DebugOut("updateGroupInfo", 16, lStrError:=ex.StackTrace)
            End Try
        End If

        DebugOut("updateGroupInfo", 8, lObjDateTime:=_DbgInitTime)

    End Sub







    Public Sub updateGroupInfo_(lObjAssignmentInfo As cArsAssignmentInfo)

        Dim _DbgInitTime As DateTime = Now()

        Dim lObjLdapConn As LdapConnection = getLdapConnection()

        If lObjLdapConn IsNot Nothing Then

            Try

                Dim lIntLoadRange As Integer = 1500
                Dim lBolEndIsNear As Boolean = False

                Dim lLstAttributes As String() = {"member"}

                Dim lObjTtlControl As New DirectoryControl(
                "1.2.840.113556.1.4.2309", Nothing, True, True)

                Dim lObjSearchRequest As New SearchRequest(
                    dnGetParent(lObjAssignmentInfo.gDn),
                    String.Format("(&(objectClass=group)(objectSid={0}))", lObjAssignmentInfo.gSID),
                    System.DirectoryServices.Protocols.SearchScope.OneLevel,
                    lLstAttributes)

                Dim lObjResponse As SearchResponse = Nothing

                lObjSearchRequest.SizeLimit = 1
                lObjSearchRequest.Controls.Add(lObjTtlControl)
                lObjResponse = lObjLdapConn.SendRequest(lObjSearchRequest)

                If lObjResponse.Entries.Count > 0 Then

                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                        Dim u As Integer =
                            _ArsUser.Users.FindIndex(Function(p) p.SID = lObjAssignmentInfo.uSID)

                        If u >= 0 Then

                            Dim g As Integer =
                                _ArsUser.Users(u).Groups.FindIndex(Function(p) p.SID = lObjAssignmentInfo.gSID)

                            If g >= 0 Then

                                _ArsUser.Users(u).Groups(g).TTL = 0

                                If lObjEntry.Attributes.Contains("member") Then

                                    Dim lStrGroupMemberAttribute As String = "member"

                                    For Each lStrlStrAttributeName As String In lObjEntry.Attributes.AttributeNames

                                        If lStrlStrAttributeName.ToLower().IndexOf("member;range") >= 0 Then

                                            lStrGroupMemberAttribute = lStrlStrAttributeName
                                            Exit For
                                        End If
                                    Next

                                    For i = 0 To lObjEntry.Attributes(lStrGroupMemberAttribute).Count - 1

                                        If lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString().IndexOf(_ArsUser.Users(u).distinguishedName) >= 0 Then

                                            If lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString().StartsWith("<TTL=") Then

                                                Dim lObjMatch As Match = Regex.Match(lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString(), "^<TTL=([0-9]+)>,.*$")

                                                If lObjMatch.Success Then

                                                    _ArsUser.Users(u).Groups(g).TTL = CLng(lObjMatch.Groups(1).Value)
                                                End If
                                            Else

                                                _ArsUser.Users(u).Groups(g).TTL = -1
                                            End If

                                            Exit For
                                        End If
                                    Next

                                    If lStrGroupMemberAttribute.Contains("range") Then

                                        LoadLargeGroupMembership(_ArsUser.Users(u).Groups(g), _ArsUser.Users(u).distinguishedName, lStrGroupMemberAttribute)
                                    End If
                                End If

                                _ArsUser.Users(u).Groups(g).TimeStamp = Math.Round(DateTime.UtcNow.Subtract(
                                    New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, 0)
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception

                DebugOut("updateGroupInfo", 16, lStrError:=ex.StackTrace)
            End Try
        End If

        DebugOut("updateGroupInfo", 8, lObjDateTime:=_DbgInitTime)

    End Sub

    Private Sub LoadLargeGroupMembership(ByRef _ArsAssignedGroup As cArsAssignedGroup, ByVal lStrMemberDn As String, ByVal lStrGroupMemberAttribute As String)

        Dim _DbgInitTime As DateTime = Now()

        Dim lObjLdapConn As LdapConnection = getLdapConnection()

        If lObjLdapConn IsNot Nothing Then

            Try

                Dim lIntLoadRange As Integer = 1500
                Dim lBolEndIsNear As Boolean = False

                Dim lLstAttributes As String() =
                    {lStrGroupMemberAttribute}

                Dim lObjTtlControl As New DirectoryControl(
                "1.2.840.113556.1.4.2309", Nothing, True, True)

                Dim lObjSearchRequest As New SearchRequest(
                    dnGetParent(_ArsAssignedGroup.distinguishedName),
                    String.Format("(&(objectClass=group)(objectSid={0}))", _ArsAssignedGroup.SID),
                    System.DirectoryServices.Protocols.SearchScope.OneLevel,
                    lLstAttributes)

                lObjSearchRequest.Controls.Add(lObjTtlControl)
                lObjSearchRequest.SizeLimit = 1

                Dim lObjResponse As SearchResponse = Nothing

                While Not lBolEndIsNear

                    If lStrGroupMemberAttribute.IndexOf("=") > 0 And lStrGroupMemberAttribute.IndexOf("-") > 0 Then

                        Dim lStrRange() As String = lStrGroupMemberAttribute.Split("=")(1).Split("-")

                        lObjSearchRequest.Attributes(0) = String.Format("member;range={0}-{1}", CInt(lStrRange(1)) + 1, CInt(lStrRange(1)) + lIntLoadRange)

                        lObjResponse = lObjLdapConn.SendRequest(lObjSearchRequest)

                        If lObjResponse.Entries.Count > 0 Then

                            For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                                For Each lStrAttributeName As String In lObjEntry.Attributes.AttributeNames

                                    If lStrAttributeName.ToLower().IndexOf("member;range") >= 0 Then

                                        lStrGroupMemberAttribute = lStrAttributeName

                                        If lStrAttributeName.IndexOf("*") > 0 Then

                                            lBolEndIsNear = True
                                        End If
                                        Exit For
                                    End If
                                Next

                                For i = 0 To lObjEntry.Attributes(lStrGroupMemberAttribute).Count - 1

                                    If lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString().IndexOf(lStrMemberDn) >= 0 Then

                                        If lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString().StartsWith("<TTL=") Then

                                            Dim lObjMatch As Match = Regex.Match(lObjEntry.Attributes(lStrGroupMemberAttribute)(i).ToString(), "^<TTL=([0-9]+)>,.*$")

                                            If lObjMatch.Success Then

                                                _ArsAssignedGroup.TTL = CLng(lObjMatch.Groups(1).Value)
                                            End If
                                        Else

                                            _ArsAssignedGroup.TTL = -1
                                        End If

                                        lBolEndIsNear = True
                                        Exit For
                                    End If

                                    If lBolEndIsNear Then Exit For
                                Next
                            Next
                        End If
                    Else

                        lBolEndIsNear = True
                    End If
                End While

                lObjLdapConn.Dispose()

            Catch ex As Exception

                DebugOut("LoadLargeGroupMembership", 16, lStrError:=ex.StackTrace)
            End Try
        End If

        DebugOut("LoadLargeGroupMembership", 8, lObjDateTime:=_DbgInitTime)

    End Sub

End Class
