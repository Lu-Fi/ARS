Imports System.Data.SqlClient
Imports System.DirectoryServices.Protocols
Imports System.Net
Imports System.Security.Principal

Namespace Ars

    Public Class Ldap

        Private Class SearchEntry

            Public Value As String = Nothing
            Public State As Integer = 0
            Public ShortName As String = Nothing
            Public IsDistinguishedName As Boolean = False

            Public Sub New(_value As Object)

                Value = _value.ToString()
                IsDistinguishedName = (_value.IndexOf("=") > -1)

                If IsDistinguishedName Then

                    Value = Regex.Match(Value, "cn=.*$", RegexOptions.IgnoreCase).Value

                    Dim _match As Match =
                        Regex.Match(_value, "cn=([^,]+)", RegexOptions.IgnoreCase)

                    ShortName = _match.Groups(1).ToString()
                End If
            End Sub
        End Class

        Public ReadOnly Cache As Cache

        Private ReadOnly HttpContext As HttpContext
        Private ReadOnly ArsApplicationSettings As ApplicationSettings

        Private SqlConnection As SqlConnection = Nothing
        Private ReadOnly SqlConnectionString As String = Nothing

        Public Sub New(Optional _HttpContext As HttpContext = Nothing, Optional _cache As Cache = Nothing)

            If _HttpContext IsNot Nothing Then

                HttpContext = _HttpContext
            Else

                HttpContext = HttpContext.Current
            End If

            If HttpContext IsNot Nothing Then

                ArsApplicationSettings = HttpContext.Application("ApplicationSettings")
            Else

                ArsApplicationSettings = New ApplicationSettings
            End If

            If _cache IsNot Nothing Then

                Cache = _cache
            Else

                Cache = New Cache
            End If

            SqlConnectionString =
                ConfigurationManager.AppSettings.Item("sqlConnectionString")
        End Sub

        Public Function GetLdapConnection() As LdapConnection

            Dim ml As New MethodLogger("Ars.Ldap")

            Try

                Dim lObjCreds As New NetworkCredential()

                Dim lObjLdapIdent As New LdapDirectoryIdentifier(
                    ArsApplicationSettings.LdapServer, ArsApplicationSettings.LdapPort, True, False)

                Dim lObjLdapConn As New LdapConnection(
                    lObjLdapIdent,
                    lObjCreds,
                    AuthType.Negotiate)

                Try

                    lObjLdapConn.Bind()
                Catch ex As Exception

                    If ArsApplicationSettings.IsPersistentLdapServer Then

                        ml.write(LOGLEVEL.ERR,
                        String.Format("Unable to bind to LDAP Server: {0}", ArsApplicationSettings.LdapServer))

                        ArsApplicationSettings.IsPersistentLdapServer = False
                        ArsApplicationSettings.LdapServer = ConfigurationManager.AppSettings.Item("ldap_Server")
                        lObjLdapConn = GetLdapConnection()
                    End If
                End Try

                If Not ArsApplicationSettings.IsPersistentLdapServer Then

                    getPersistentLdapServer(lObjLdapConn)
                End If

                ml.done(ArsApplicationSettings.LdapServer)

                Return lObjLdapConn
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR,
             String.Format("Unable to get LDAP connection on server: {0}, {1}", ArsApplicationSettings.LdapServer, ex.Message))

                ml.done(ArsApplicationSettings.LdapServer)

                Return Nothing
            End Try
        End Function

        Private Sub getPersistentLdapServer(ByRef lObjLdapConn As LdapConnection)

            Dim ml As New MethodLogger("Ars.Ldap")

            Dim lObjSearchRequest As New SearchRequest(
                "",
                "(&(objectClass=*))",
                SearchScope.Base,
                "dnsHostName")

            lObjSearchRequest.SizeLimit = 1

            Dim lObjResponse As SearchResponse =
                lObjLdapConn.SendRequest(lObjSearchRequest)

            If lObjResponse.Entries.Count > 0 Then

                For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                    If lObjEntry.Attributes.Contains("dnshostname") Then

                        ArsApplicationSettings.LdapServer = lObjEntry.Attributes("dnshostname")(0)
                        ArsApplicationSettings.IsPersistentLdapServer = True

                        ml.write(LOGLEVEL.INFO,
                         String.Format("Set persistent LDAP server to: {0}", ArsApplicationSettings.LdapServer))
                    End If
                Next
            End If

            ml.done()

        End Sub

        Sub RefreshUserFromLdap(_Assignment As Assignment, Optional doSearch As Boolean = False)

            RefreshUserFromLdap(New AssignedUser With {
                .SID = _Assignment.aSID
            }, True)
        End Sub

        Sub RefreshUserFromLdap(_AssignedUser As AssignedUser, Optional doSearch As Boolean = False)

            Dim ml As New MethodLogger("Ars.Ldap",
                                       String.Format("SID = {0}, doSearch = {0}", doSearch))

            Dim lStrDn As String

            Dim lObjScope As SearchScope =
                SearchScope.Base

            If doSearch = True Then

                lStrDn = ArsApplicationSettings.UserSearchBase
                lObjScope = SearchScope.Subtree
            Else

                lStrDn = _AssignedUser.DN
            End If

            Dim lObjLdapConn As LdapConnection =
                    getLdapConnection()

            Dim lObjTtlControl As New DirectoryControl(
                "1.2.840.113556.1.4.2309", Nothing, True, True)

            Dim lObjSearchRequest As New SearchRequest(
                lStrDn,
                String.Format("(&(objectClass=user)(objectSid={0}))", _AssignedUser.SID),
                lObjScope,
                ArsApplicationSettings.LdapUserAttributes)

            lObjSearchRequest.Controls.Add(lObjTtlControl)

            If doSearch = True Or String.IsNullOrEmpty(lStrDn) Then

                lObjSearchRequest.DistinguishedName = ArsApplicationSettings.UserSearchBase
                lObjSearchRequest.Scope = SearchScope.Subtree
            End If

            Dim lObjResponse As SearchResponse =
                lObjLdapConn.SendRequest(lObjSearchRequest)

            If lObjResponse.Entries.Count > 0 Then

                For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                    Dim lObjEntrySid As String = New SecurityIdentifier(
                        lObjEntry.Attributes("objectSid")(0), 0).ToString()

                    ml.write(LOGLEVEL.DEBUG,
                            String.Format("{0}, update user data from ldap.", lObjEntrySid))

                    Cache.UpdateUserFromLdap(lObjEntry.Attributes, _AssignedUser)

                    Cache.CacheUser(lObjEntrySid, _AssignedUser, DateAdd(DateInterval.Second, _AssignedUser.MinGroupTTL, Now()))
                Next
            Else

                RefreshUserFromLdap(_AssignedUser, True)
            End If

            ml.done()
        End Sub

        Sub LoadLdapUser(ByRef _AssignedUsers As List(Of AssignedUser), Optional _SearchRoot As String = Nothing, Optional _SearchScope As SearchScope = SearchScope.Subtree, Optional _DisableCache As Boolean = False)

            Dim ml As New MethodLogger("Ars.Ldap",
                                        String.Format("[], Root = '{0}', Scope = '{1}', DisableCache = '{2}'",
                                             _SearchRoot, _SearchScope, _DisableCache))

            Dim lStrLdapFilter As String = vbNullString

            Dim lStrCacheMessage As String = "not found in cache"
            If _DisableCache Then

                lStrCacheMessage = "will be updated from ldap, cache is disabled"
            End If

            If _SearchRoot Is Nothing Then

                _SearchRoot = ArsApplicationSettings.UserSearchBase
            End If

            Dim _CachedUser As CachedUser = Nothing

            For Each _AssignedUser In _AssignedUsers

                _CachedUser = HttpContext.Cache.Get(_AssignedUser.SID)

                If _CachedUser Is Nothing Or _DisableCache Then

                    lStrLdapFilter += String.Format("(objectSid={0})", _AssignedUser.SID)

                    ml.write(LOGLEVEL.DEBUG,
                         String.Format("{0}, account {1}", _AssignedUser.SID, lStrCacheMessage))
                Else

                    Cache.UpdateUserFromCache(_CachedUser, _AssignedUser)

                    ml.write(LOGLEVEL.DEBUG,
                         String.Format("{0}, read account data from cache.", _AssignedUser.SID))
                End If
            Next

            If Not String.IsNullOrEmpty(lStrLdapFilter) Then

                Dim lObjLdapConn As LdapConnection =
                    getLdapConnection()

                Dim lObjTtlControl As New DirectoryControl(
                "1.2.840.113556.1.4.2309", Nothing, True, True)

                Dim lObjSearchRequest As New SearchRequest(
                    _SearchRoot,
                    lStrLdapFilter,
                    _SearchScope,
                    ArsApplicationSettings.LdapUserAttributes)

                Dim lObjResponse As SearchResponse = Nothing

                lObjSearchRequest.Controls.Add(lObjTtlControl)
                lObjSearchRequest.Filter = String.Format("(&(objectClass=user)(|{0}))", lStrLdapFilter)
                lObjResponse = lObjLdapConn.SendRequest(lObjSearchRequest)

                If lObjResponse.Entries.Count > 0 Then

                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                        Dim lObjEntrySid As String = New SecurityIdentifier(
                            lObjEntry.Attributes("objectSid")(0), 0).ToString()

                        ml.write(LOGLEVEL.DEBUG,
                             String.Format("{0}, update account data from ldap.", lObjEntrySid))

                        Dim _AssignedUser As AssignedUser =
                            _AssignedUsers.FirstOrDefault(Function(n) n.SID = lObjEntrySid)

                        If _AssignedUser IsNot Nothing Then

                            Cache.UpdateUserFromLdap(lObjEntry.Attributes, _AssignedUser)

                            Cache.CacheUser(lObjEntrySid, _AssignedUser, DateAdd(DateInterval.Second, _AssignedUser.MinGroupTTL, Now()))
                        End If
                    Next
                End If
            End If

            ml.done()
        End Sub

        Sub LoadLdapGroup(ByRef _AssignedGroups As List(Of AssignedGroup), Optional _SearchRoot As String = Nothing, Optional _SearchScope As SearchScope = SearchScope.Subtree, Optional _DisableCache As Boolean = False)

            Dim ml As New MethodLogger("Ars.Ldap",
                                        String.Format("[], Root = '{0}', Scope = '{1}', DisableCache = '{2}'",
                                             _SearchRoot, _SearchScope, _DisableCache))

            Dim lStrCacheMessage As String = "not found in cache"
            If _DisableCache Then

                lStrCacheMessage = "will be updated from ldap, cache is disabled"
            End If

            Dim lStrLdapFilter As String = vbNullString

            If _SearchRoot Is Nothing Then

                _SearchRoot = ArsApplicationSettings.GroupSearchBase
            End If

            Dim _CachedGroup As CachedGroup = Nothing

            For Each _AssignedGroup In _AssignedGroups

                _CachedGroup = HttpContext.Cache.Get(_AssignedGroup.SID)

                If _CachedGroup Is Nothing Or _DisableCache Then

                    lStrLdapFilter += String.Format("(objectSid={0})", _AssignedGroup.SID)

                    ml.write(LOGLEVEL.DEBUG,
                         String.Format("{0}, group {1}.", _AssignedGroup.SID, lStrCacheMessage))
                Else

                    Cache.UpdateGroupFromCache(_CachedGroup, _AssignedGroup)

                    ml.write(LOGLEVEL.DEBUG,
                         String.Format("{0}, read group data from cache.", _AssignedGroup.SID))

                End If
            Next

            If Not String.IsNullOrEmpty(lStrLdapFilter) Then

                Dim lObjLdapConn As LdapConnection =
                    getLdapConnection()

                Dim lObjTtlControl As New DirectoryControl(
                "1.2.840.113556.1.4.2309", Nothing, True, True)

                Dim lObjSearchRequest As New SearchRequest(
                    _SearchRoot,
                    lStrLdapFilter,
                    _SearchScope,
                    ArsApplicationSettings.LdapGroupAttributes)

                Dim lObjResponse As SearchResponse = Nothing

                lObjSearchRequest.Controls.Add(lObjTtlControl)
                lObjSearchRequest.Filter = String.Format("(&(objectClass=group)(|{0}))", lStrLdapFilter)
                lObjResponse = lObjLdapConn.SendRequest(lObjSearchRequest)

                If lObjResponse.Entries.Count > 0 Then

                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                        Dim lObjEntrySid As String = New SecurityIdentifier(
                            lObjEntry.Attributes("objectSid")(0), 0).ToString()

                        ml.write(LOGLEVEL.DEBUG,
                            String.Format("{0}, update group data from ldap.", lObjEntrySid))

                        Dim _AssignedGroup As AssignedGroup =
                            _AssignedGroups.FirstOrDefault(Function(n) n.SID = lObjEntrySid)

                        If _AssignedGroup IsNot Nothing Then

                            Cache.UpdateGroupFromLdap(lObjEntry.Attributes, _AssignedGroup)

                            Cache.CacheGroup(lObjEntrySid, _AssignedGroup)
                        End If
                    Next
                End If
            End If

            ml.done()

        End Sub

        Sub GetLdapUserAssignments(_User As User)

            Dim ml As New MethodLogger("Ars.Ldap")

            If Not String.IsNullOrEmpty(_User.UPN) Then

                Dim lObjLdapConn As LdapConnection =
                    getLdapConnection()

                Dim lObjTtlControl As New DirectoryControl(
                    "1.2.840.113556.1.4.2309", Nothing, True, True)

                Dim lObjSearchRequest As New SearchRequest(
                    ArsApplicationSettings.UserSearchBase,
                    String.Format("(&(objectClass=user)(extensionAttribute15={0}))", _User.UPN),
                    SearchScope.Subtree,
                    ArsApplicationSettings.LdapUserAttributes)

                lObjSearchRequest.Controls.Add(lObjTtlControl)

                Dim lObjResponse As SearchResponse =
                    lObjLdapConn.SendRequest(lObjSearchRequest)

                If lObjResponse.Entries.Count > 0 Then

                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                        Dim _FoundAssignedUser As New AssignedUser()

                        Dim lObjEntrySid As String = New SecurityIdentifier(
                            lObjEntry.Attributes("objectSid")(0), 0).ToString()

                        ml.write(LOGLEVEL.DEBUG,
                            String.Format("{0}, update user data from ldap.", lObjEntrySid))

                        Cache.UpdateUserFromLdap(lObjEntry.Attributes, _FoundAssignedUser)

                        Cache.CacheUser(lObjEntrySid, _FoundAssignedUser, DateAdd(DateInterval.Second, _FoundAssignedUser.MinGroupTTL, Now()))

                        Dim lLstPerms As List(Of Object) =
                                Cache.getListFromAttribute(lObjEntry.Attributes, "allowedAttributesEffective")

                        If Not _User._Assignments.Exists(Function(a) a.aSID = lObjEntrySid And a.uSID = _User.SID) Then

                            If lLstPerms IsNot Nothing Then

                                If lLstPerms.Contains("userAccountControl") Then

                                    _User._Assignments.Add(New Assignment With {
                                        .aID = Guid.NewGuid(),
                                        .uSID = _User.SID,
                                        .aSID = lObjEntrySid,
                                        .source = ASS_SOURCE.LDAP,
                                        .flags = 1,
                                        .updateId = _User.updateId
                                    })
                                End If
                            End If
                        Else

                            If lLstPerms IsNot Nothing Then

                                If lLstPerms.Contains("userAccountControl") Then

                                    _User._Assignments.FindAll(
                                        Function(a) a.aSID = lObjEntrySid And a.uSID = _User.SID).ForEach(
                                            Sub(a) a.updateId = _User.updateId)
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            ml.done()

        End Sub

        Sub GetLdapGroupAssignments(_User As User)

            Dim ml As New MethodLogger("Ars.Ldap",
                                       String.Format("_User: {0} ({1})", _User.SID, _User.Name))

            Dim lLstGroups As New List(Of Object)

            lLstGroups.AddRange(_User.memberOf)

            _User.AssignedUsers.ForEach(
                Sub(u) lLstGroups.AddRange(u.memberOf))

            CacheLdapGroupByName(lLstGroups.Distinct().ToList())

            AddGroupAssignments(_User, _User.memberOf)

            _User.AssignedUsers.ForEach(
                Sub(u) AddGroupAssignments(u, u.memberOf))

            ml.done()
        End Sub

        Public Sub AddGroupAssignments(_AssignedUser As AssignedUser, _Groups As List(Of Object), Optional _processedGroups As List(Of Object) = Nothing)

            Dim lBolAddAssignments As Boolean =
                False

            'Normally _processedGroups is passed if this function is called recursive
            If _processedGroups Is Nothing Then

                _processedGroups = New List(Of Object)
            Else

                lBolAddAssignments = True
            End If

            For Each _group In _Groups

                If Not _processedGroups.Contains(_group) Then

                    _processedGroups.Add(_group)

                    Dim lStrSid As String =
                        HttpContext.Cache.Get(_group)

                    If lStrSid IsNot Nothing Then

                        If lStrSid <> "-" Then

                            Dim lObjGroup As CachedGroup =
                                HttpContext.Cache.Get(lStrSid)

                            If lObjGroup.managedObjects.Count > 0 Then

                                AddGroupAssignments(_AssignedUser, lObjGroup.managedObjects, _processedGroups)
                            Else

                                If lBolAddAssignments Then

                                    If _AssignedUser.User IsNot Nothing Then

                                        AddGroupAssignmentIfNotExists(_AssignedUser.User._Assignments, lStrSid, _AssignedUser.SID, _AssignedUser.User.updateId)
                                        '_AssignedUser.User._Assignments.Add(New Assignment With {
                                        '    .aID = Guid.NewGuid,
                                        '    .uSID = _AssignedUser.SID,
                                        '    .aSID = lStrSid,
                                        '    .source = ASS_SOURCE.LDAP,
                                        '    .flags = 2
                                        '})
                                    Else

                                        AddGroupAssignmentIfNotExists(_AssignedUser._Assignments, lStrSid, _AssignedUser.SID)
                                        '_AssignedUser._Assignments.Add(New Assignment With {
                                        '    .aID = Guid.NewGuid,
                                        '    .uSID = _AssignedUser.SID,
                                        '    .aSID = lStrSid,
                                        '    .source = ASS_SOURCE.LDAP,
                                        '    .flags = 2
                                        '})
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End Sub

        Public Sub AddGroupAssignmentIfNotExists(_Assignments As List(Of Assignment), _aSid As String, _uSid As String, Optional _updateId As Integer = 0)

            If Not _Assignments.Exists(Function(a) a.aSID = _aSid And a.uSID = _uSid) Then

                _Assignments.Add(New Assignment With {
                                            .aID = Guid.NewGuid,
                                            .uSID = _uSid,
                                            .aSID = _aSid,
                                            .source = ASS_SOURCE.LDAP,
                                            .flags = 2
                                        })

            Else

                _Assignments.FindAll(Function(a) a.aSID = _aSid And a.uSID = _uSid).ForEach(Sub(a) a.updateId = _updateId)
            End If
        End Sub
        Public Sub CacheLdapGroupByName(_searchList As List(Of Object), Optional forceSearch As Boolean = False)

            Dim ml As New MethodLogger("Ars.Ldap", "[]")

            Dim _managedByFilter As String = "(|(managedObjects=*)(managedBy=*))"
            Dim _EntryCount As Integer = 0
            Dim _searchEntryList As New List(Of SearchEntry)

            Dim _ProcessedList As New List(Of String)

            Dim lObjLdapConn As LdapConnection = Nothing
            Dim lObjSearchRequest As New SearchRequest(
                "",
                "",
                SearchScope.Subtree,
                ArsApplicationSettings.LdapGroupAttributes)

            _searchList.RemoveAll(
                    Function(e) HttpContext.Cache.Get(e.ToString()) IsNot Nothing)

            _searchList.ForEach(
                Sub(e) _searchEntryList.Add(New SearchEntry(e)))

            While _searchEntryList.FindAll(Function(e) e.State = 0).Count > 0

                _EntryCount = 0
                Dim lStrDnFilter As String = ""
                Dim lStrSidFilter As String = ""

                If lObjLdapConn Is Nothing Then

                    lObjLdapConn = getLdapConnection()
                End If

                If _searchEntryList.FindAll(Function(e) e.State = 0).Count > 10 Or forceSearch Then
                    'USE SEARCH

                    lObjSearchRequest.DistinguishedName = ArsApplicationSettings.GroupSearchBase
                    lObjSearchRequest.Scope = SearchScope.Subtree

                    'BUILD LDAP QUERY
                    For Each _entry In _searchEntryList.FindAll(Function(e) e.State = 0)

                        _EntryCount += 1
                        _entry.State = 1

                        If _entry.IsDistinguishedName Then

                            lStrDnFilter = String.Format("{0}(name={1})", lStrDnFilter, _entry.ShortName)

                            ml.write(LOGLEVEL.DEBUG,
                                String.Format("{0}, add entry to search.", _entry.ShortName))
                        Else

                            lStrSidFilter = String.Format("{0}(objectSid={1})", lStrSidFilter, _entry.Value)

                            ml.write(LOGLEVEL.DEBUG,
                                String.Format("{0}, add entry to search.", _entry.Value))
                        End If

                        If _EntryCount >= ArsApplicationSettings.LdapMaxSearchEntries Then

                            Exit For
                        End If
                    Next

                    If lStrDnFilter.Length > 0 Or lStrSidFilter.Length > 0 Then

                        If lStrDnFilter.Length > 0 Then

                            lStrDnFilter = String.Format("(&{0}(|{1}))", _managedByFilter, lStrDnFilter)
                        End If

                        lObjSearchRequest.Filter =
                            String.Format("(&(objectClass=group)(|{0}{1}))", lStrDnFilter, lStrSidFilter)

                        ml.write(LOGLEVEL.VERBOSE,
                                    String.Format("ldap filter length: {0} bytes for {1} Groups",
                                                  lObjSearchRequest.Filter.Length, _searchEntryList.Count))
                    End If

                Else
                    'ITEM BY ITEM

                    With _searchEntryList.First(Function(e) e.State = 0)
                        .State = 1
                        lObjSearchRequest.DistinguishedName = .Value
                    End With

                    ml.write(LOGLEVEL.DEBUG,
                                String.Format("{0}, do base search.", lObjSearchRequest.DistinguishedName))

                    lObjSearchRequest.Scope = SearchScope.Base
                    lObjSearchRequest.Filter = "(&(objectClass=group))"
                End If

                Dim lObjResponse As SearchResponse =
                    lObjLdapConn.SendRequest(lObjSearchRequest)

                ml.write(LOGLEVEL.VERBOSE,
                         String.Format("ldap search found: {0} objects",
                                       lObjResponse.Entries.Count))

                ml.write(LOGLEVEL.DEBUG, lObjSearchRequest.Filter)

                If lObjResponse.Entries.Count > 0 Then

                    For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                        Dim lObjEntrySid As String = New SecurityIdentifier(
                                lObjEntry.Attributes("objectSid")(0), 0).ToString()

                        ml.write(LOGLEVEL.DEBUG,
                                String.Format("{0}, update group data from ldap.", lObjEntry.DistinguishedName))

                        Dim _AssignedGroup As AssignedGroup =
                                New AssignedGroup

                        Cache.UpdateGroupFromLdap(lObjEntry.Attributes, _AssignedGroup)
                        Cache.CacheGroup(lObjEntrySid, _AssignedGroup)

                        For Each _managedObject In Cache.getListFromAttribute(lObjEntry.Attributes, "managedObjects")

                            If HttpContext.Cache.Get(_managedObject) Is Nothing And Not _searchEntryList.Exists(Function(e) e.Value = _managedObject) Then

                                _searchEntryList.Add(New SearchEntry(_managedObject))

                                ml.write(LOGLEVEL.DEBUG,
                                    String.Format("{0}, add managedObject to search list.", _managedObject))
                            End If
                        Next

                        _searchEntryList.FindAll(
                            Function(e) e.State = 1 And (e.Value = lObjEntrySid Or e.Value = lObjEntry.DistinguishedName)).ForEach(Sub(e) e.State = 2)
                    Next
                End If

                _searchEntryList.FindAll(
                    Function(e) e.State = 1).ForEach(
                    Sub(e)
                        e.State = 2
                        If e.IsDistinguishedName Then

                            Cache.CacheDummyGroup(e.Value)
                            ml.write(LOGLEVEL.DEBUG,
                                String.Format("{0}, write dummy cache entry.", e.Value))
                        End If
                    End Sub)

                _managedByFilter = ""
            End While

            ml.done()

        End Sub

        Public Function SetAccountUac(_AssignedUser As AssignedUser, _flag As Integer, _set As Boolean, Optional _refresh As Boolean = True) As Integer

            Dim ml As New MethodLogger("Ars.Ldap",
                                       String.Format("sid = '{0}', flag = '{1}', set = '{2}'", _AssignedUser.SID, _flag, _set))

            Dim _uac As Integer
            Dim _return As Integer = -1

            If _refresh Then

                RefreshUserFromLdap(_AssignedUser)
            End If

            If Not _AssignedUser.UAC.arsIsBitSet(_flag) And _set = True Then

                _uac = _AssignedUser.UAC Or _flag
            ElseIf _AssignedUser.UAC.arsIsBitSet(_flag) And _set = False Then

                _uac = _AssignedUser.UAC And Not _flag
            Else

                _uac = _AssignedUser.UAC
            End If

            If _uac <> _AssignedUser.UAC Then

                ml.write(LOGLEVEL.INFO,
                    String.Format("Set userAccountControl Flag '{1}' for '{0}' = {2}.", _AssignedUser.SID, _flag, _set))

                Dim lObjLdapConnection As LdapConnection = getLdapConnection()
                Dim lObjDirectoryAttributeModification As New DirectoryAttributeModification With {
                    .Name = "userAccountControl"
                }
                lObjDirectoryAttributeModification.Add(_uac)
                lObjDirectoryAttributeModification.Operation = DirectoryAttributeOperation.Replace

                Dim lObjModifyRequest As New ModifyRequest(
                _AssignedUser.DN, lObjDirectoryAttributeModification)

                Try

                    Dim lObjModifyResponse As ModifyResponse =
                        lObjLdapConnection.SendRequest(lObjModifyRequest)

                    If lObjModifyResponse.ResultCode = 0 Then

                        _return = 1
                    End If

                Catch ex As Exception

                    ml.write(LOGLEVEL.ERR,
                        String.Format("ERROR: {0}", ex.GetBaseException().Message))
                End Try

                lObjLdapConnection.Dispose()

                If _refresh Then

                    RefreshUserFromLdap(_AssignedUser)
                End If
            Else

                _return = 2

                ml.write(LOGLEVEL.INFO,
                    String.Format("Set userAccountControl Flag '{1}' for '{0}' = {2}. No work to do. Already set.", _AssignedUser.SID, _flag, _set))
            End If

            ml.done()

            Return _return

        End Function

        Public Function AddMemberToGroup(_AssignedUser As AssignedUser, _group As String, Optional _time As Integer = 0) As Integer

            Dim ml As New MethodLogger("Ars.Ldap",
                                       String.Format("user = '{0}', group = '{1}'", _AssignedUser.DN, _group))

            Dim _return As Integer

            Try

                Dim lStrDn As String = _AssignedUser.DN
                Dim lIntResult As Integer = 0

                If _time > 0 Then

                    lStrDn =
                        String.Format("<TTL={0},{1}>", _time, _AssignedUser.DN)
                End If

                Dim lObjLdapConnection As LdapConnection = getLdapConnection()

                Dim lObjDirectoryAttributeModification As New DirectoryAttributeModification() With {.Name = "member"}
                lObjDirectoryAttributeModification.Add(lStrDn)
                lObjDirectoryAttributeModification.Operation = DirectoryAttributeOperation.Add

                Dim lObjModifyRequest As New ModifyRequest(
                                _group, lObjDirectoryAttributeModification)

                Dim lObjModifyResponse As ModifyResponse =
                                lObjLdapConnection.SendRequest(lObjModifyRequest)

                lObjLdapConnection.Dispose()

                lIntResult =
                    SetAccountUac(_AssignedUser, 2, False, False)

                If lIntResult < 2 Then

                    audit.add(_AssignedUser, "Enable user", lIntResult)
                End If

                If lObjModifyResponse.ResultCode = 0 Then

                    _return = 1
                Else

                    _return = lObjModifyResponse.ResultCode
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR,
                     String.Format("ERROR addMemberToGroup: {0}", ex.Message))
            End Try

            audit.add(_AssignedUser, String.Format("Add member to group: {0} for {1} seconds.", _group, _time), _return)

            ml.done()

            Return _return

        End Function

        Public Function RemoveMemberFromGroup(_AssignedUser As AssignedUser, _group As String) As Integer

            Dim _return As Integer

            Dim ml As New MethodLogger("Ars.Ldap",
                                       String.Format("user = '{0}', group = '{1}'", _AssignedUser.DN, _group))

            Try

                Dim lObjLdapConnection As LdapConnection = getLdapConnection()

                Dim lObjDirectoryAttributeModification As New DirectoryAttributeModification()
                lObjDirectoryAttributeModification.Name = "member"
                lObjDirectoryAttributeModification.Add(_AssignedUser.DN)
                lObjDirectoryAttributeModification.Operation = DirectoryAttributeOperation.Delete

                Dim lObjModifyRequest As New ModifyRequest(
                                _group, lObjDirectoryAttributeModification)

                Dim lObjModifyResponse As ModifyResponse =
                                lObjLdapConnection.SendRequest(lObjModifyRequest)

                lObjLdapConnection.Dispose()

                If lObjModifyResponse.ResultCode = 0 Then

                    _return = 1
                Else

                    _return = lObjModifyResponse.ResultCode
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR,
                     String.Format("ERROR removeMemberFromGroup: {0}", ex.Message))
            End Try

            audit.add(_AssignedUser, String.Format("Remove member from group: {0}", _group), _return)

            ml.done()

            Return _return

        End Function

        Public Function ResetPassword(_AssignedUser As AssignedUser) As String

            Dim ml As New MethodLogger("Ars.Ldap",
                                       String.Format("_AssignedUser = '{0}'", _AssignedUser.SID))

            Dim _return As String = Nothing


            Dim _se As String = System.Web.Security.Membership.GeneratePassword(
                            ArsApplicationSettings.pwdLength, ArsApplicationSettings.pwdSpecialCharacters)

            If Not String.IsNullOrEmpty(_AssignedUser.DN) And Not String.IsNullOrEmpty(_se) Then

                Dim lObjLdapConnection As LdapConnection = getLdapConnection()
                Dim lObjDirectoryAttributeModification As New DirectoryAttributeModification With {
                    .Name = "unicodePwd"
                }
                lObjDirectoryAttributeModification.Add(Encoding.Unicode.GetBytes(Chr(34) & _se & Chr(34)))
                lObjDirectoryAttributeModification.Operation = DirectoryAttributeOperation.Replace

                Dim lObjModifyRequest As New ModifyRequest(
                    _AssignedUser.DN, lObjDirectoryAttributeModification)

                lObjModifyRequest.Controls.Add(New DirectoryControl("1.2.840.113556.1.4.2066", Nothing, False, True))
                lObjModifyRequest.Controls.Add(New DirectoryControl("1.2.840.113556.1.4.2239", Nothing, False, True))

                Try

                    Dim lObjModifyResponse As ModifyResponse =
                        lObjLdapConnection.SendRequest(lObjModifyRequest)

                    _return = _se
                Catch ex As Exception

                    ml.write(LOGLEVEL.ERR,
                         String.Format("ERROR removeMemberFromGroup: {0}", ex.GetBaseException().Message))
                    _return = ex.GetBaseException().Message
                End Try

                RefreshUserFromLdap(_AssignedUser)

                lObjLdapConnection.Dispose()
            End If

            Dim lIntResult As Integer = 0

            If _return.Length > 0 Then

                lIntResult = 1
            End If

            audit.add(_AssignedUser, "Reset password", lIntResult)

            ml.done()

            Return _return
        End Function

        Public Sub ProcessAdvancedUserFeatures(_user As AssignedUser)

            Dim ml As New MethodLogger("Ars.Ldap", String.Format("{0} ({1})", _user.SID, _user.Name))

            'CHECK FOR ACTIVE ASSIGNED GROUP MEMBERSHIPS
            If _user.memberOf IsNot Nothing Then

                For Each _Group In _user.AssignedGroups(Me)

                    For Each _membership As String In _user.memberOf

                        If _membership.IndexOf(_Group.DN) >= 0 Then

                            _user.Status = 0

                            ml.write(LOGLEVEL.VERBOSE,
                                String.Format("{0} ({1}), is member of {2}. Nothing to do.", _user.SID, _user.Name, _membership))
                            Exit For
                        End If
                    Next
                Next
            End If

            If _user.Status.arsIsBitSet(1) Then

                If _user.AutoDisableActive Then
                    'AUTO DISBALE

                    Dim lIntResult As Integer = 0

                    ml.write(LOGLEVEL.INFO,
                         String.Format("AutoDisable {0} ({1})", _user.Name, _user.SID))

                    lIntResult =
                        SetAccountUac(_user, 2, True)

                    If lIntResult < 2 Then

                        audit.add(_user, "Disable user", lIntResult)
                    End If
                End If

                If _user.AutoPasswordActive Then

                    'AUTO PASSWORD
                    Dim pwdLastSet As Int64 =
                        getPwdLastSet(_user)

                    If _user.pwdLastSet < pwdLastSet Or pwdLastSet = 0 Then

                        ml.write(LOGLEVEL.INFO,
                             String.Format("AutoPasswordReset {0} ({1})", _user.Name, _user.SID))

                        ResetPassword(_user)

                        setPwdLastSet(_user)
                    Else

                        ml.write(LOGLEVEL.INFO,
                            String.Format("AutoPasswordReset not required, pwdLastSet '{2}' done by ARS. {0} ({1})", _user.Name, _user.SID, _user.pwdLastSet))
                    End If

                    _user.Status =
                        _user.Status Or 4
                End If
            End If

            ml.done()

        End Sub

        Private Function getPwdLastSet(_AssignedUser As AssignedUser) As Int64

            SqlConnection =
                New SqlConnection(ArsApplicationSettings.SqlConnectionString)

            SqlConnection.Open()

            Dim lObjSqlCommand As New SqlCommand(
                "SELECT [pwdLastSet] FROM [ars_user] WHERE [uSID] = @uSID",
                    SqlConnection)

            lObjSqlCommand.Parameters.AddWithValue("uSID", _AssignedUser.SID)

            Dim lLngResult As Long = 0
            Dim lObjReader As SqlDataReader =
                lObjSqlCommand.ExecuteReader()

            While lObjReader.Read()

                lLngResult = lObjReader.GetInt64(0)
            End While

            lObjReader.Close()
            lObjSqlCommand.Dispose()

            SqlConnection.Close()
            SqlConnection.Dispose()

            Return lLngResult

        End Function

        Public Sub setPwdLastSet(_AssignedUser As AssignedUser, Optional Value As Int64 = -1)

            SqlConnection =
                New SqlConnection(ArsApplicationSettings.SqlConnectionString)

            SqlConnection.Open()

            Dim lObjSqlCommand As New SqlCommand(
            "IF NOT EXISTS (SELECT uSID FROM [ars_user] WHERE uSID = @uSID)
	            INSERT INTO [ars_user] ([uSID], [pwdLastSet]) VALUES (@uSID,  @pwdLastSet)
             ELSE
	            UPDATE [ars_user] SET [pwdLastSet] = @pwdLastSet WHERE [uSID] = @uSID",
            SqlConnection)

            lObjSqlCommand.Parameters.AddWithValue("uSID", _AssignedUser.SID)

            If Value >= 0 Then

                lObjSqlCommand.Parameters.AddWithValue("pwdLastSet", Value)
            Else

                lObjSqlCommand.Parameters.AddWithValue("pwdLastSet", _AssignedUser.pwdLastSet)
            End If

            lObjSqlCommand.ExecuteNonQuery()
            lObjSqlCommand.Dispose()

            SqlConnection.Close()
            SqlConnection.Dispose()

        End Sub

        Private Function getParentOu(lStrDn As String) As String

            Try

                Dim lIntPos As Integer =
                    lStrDn.IndexOf(",")

                If lIntPos >= 0 Then

                    lIntPos += 1
                    Return lStrDn.Substring(lIntPos, lStrDn.Length - lIntPos)
                End If
            Catch ex As Exception

            End Try

            Return Nothing
        End Function

    End Class
End Namespace