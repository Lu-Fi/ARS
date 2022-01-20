Imports System.DirectoryServices.Protocols

Namespace Ars

    Public Class CachedGroup

        Public name As String = Nothing
        Public info As String = Nothing
        Public description As String = Nothing
        Public dn As String = Nothing
        Public TimeStamp As DateTime = Nothing
        Public managedBy As String
        Public managedObjects As New List(Of Object)

        Public Sub New(_ArsGroup As AssignedGroup)

            With _ArsGroup
                name = .Name
                info = .Info
                description = .Description
                dn = .DN
                managedBy = .managedBy
                managedObjects = .managedObjects
                TimeStamp = .TimeStamp
            End With
        End Sub
    End Class

    Public Class CachedUser

        Public name As String = Nothing
        Public displayname As String = Nothing
        Public department As String = Nothing
        Public company As String = Nothing
        Public upn As String = Nothing
        Public uac As Integer = Nothing
        Public dn As String = Nothing
        Public info As String = Nothing
        Public description As String = Nothing
        Public pwdLastSet As Long = Nothing
        Public TimeStamp As DateTime = Nothing
        Public memberOf As New List(Of Object)
        Public lockoutTime As Long = 0

        Public Sub New(_ArsUser As User)

            With _ArsUser
                name = .Name
                displayname = .DisplayName
                department = .Department
                company = .Company
                upn = .UPN
                uac = .UAC
                dn = .DN
                info = .Info
                description = .Description
                memberOf = .memberOf
                TimeStamp = .TimeStamp
                pwdLastSet = .pwdLastSet
                lockoutTime = .lockoutTime
            End With
        End Sub

        Public Sub New(_ArsUser As AssignedUser)

            With _ArsUser
                name = .Name
                displayname = .DisplayName
                department = .Department
                company = .Company
                upn = .UPN
                uac = .UAC
                dn = .DN
                info = .Info
                description = .Description
                memberOf = .memberOf
                TimeStamp = .TimeStamp
                pwdLastSet = .pwdLastSet
                lockoutTime = .lockoutTime
            End With
        End Sub
    End Class

    Public Class Cache

        Private ReadOnly HttpContext As HttpContext
        Private ReadOnly ArsApplicationSettings As ApplicationSettings

        Sub New()

            If HttpContext.Current IsNot Nothing Then

                HttpContext = HttpContext.Current
            End If

            ArsApplicationSettings = HttpContext.Application("ApplicationSettings")
        End Sub

        Public Sub CacheDummyGroup(_Key As String)

            HttpContext.Cache.Insert(_Key, "-", Nothing,
                                     DateAdd(DateInterval.Second, ArsApplicationSettings.adGroupCacheSeconds, Now()), Nothing, CacheItemPriority.Default, Nothing)
        End Sub

        Public Sub UserExpiry(_AssignedUser As AssignedUser)

            HttpContext.Cache.Insert(_AssignedUser.aID.ToString(), _AssignedUser, Nothing,
                                     DateAdd(DateInterval.Second, _AssignedUser.MaxGroupTTL, Now()), Nothing, CacheItemPriority.NotRemovable, AddressOf CheckCallback)
        End Sub

        Public Sub RemoveUserExpiry(_AssignedUser As AssignedUser)

            HttpContext.Cache.Remove(_AssignedUser.aID.ToString())
        End Sub

        Sub CheckCallback(_key As String, _AssignedUser As AssignedUser, _reason As CacheItemRemovedReason)

            If _reason = CacheItemRemovedReason.Expired Then

                Dim _ldap As New Ldap(HttpContext, Me)

                _ldap.RefreshUserFromLdap(_AssignedUser)

                _AssignedUser.Status = 1

                _ldap.ProcessAdvancedUserFeatures(_AssignedUser)
            End If
        End Sub

        Public Sub CacheGroup(_Key As String, _AssignedGroup As AssignedGroup, Optional _Expire As Date? = Nothing)

            If Not _Expire.HasValue Then

                _Expire = DateAdd(DateInterval.Second, ArsApplicationSettings.adGroupCacheSeconds, Now())
            End If

            HttpContext.Cache.Insert(_AssignedGroup.DN, _Key, Nothing,
                                     _Expire, Nothing, CacheItemPriority.Default, Nothing)

            HttpContext.Cache.Insert(_Key, New CachedGroup(_AssignedGroup), Nothing,
                                     _Expire, Nothing, CacheItemPriority.Default, Nothing)
        End Sub

        Public Sub CacheUser(_Key As String, _AssignedUser As AssignedUser, Optional _Expire As Date? = Nothing)

            If Not _Expire.HasValue Then

                _Expire = DateAdd(DateInterval.Second, ArsApplicationSettings.adUserCacheSeconds, Now())

            Else

                If _Expire <= Now() Then

                    _Expire = DateAdd(DateInterval.Second, ArsApplicationSettings.adUserCacheSeconds, Now())
                End If
            End If

            HttpContext.Cache.Insert(_AssignedUser.DN, _Key, Nothing,
                                     _Expire, Nothing, CacheItemPriority.Default, Nothing)

            HttpContext.Cache.Insert(_Key, New CachedUser(_AssignedUser), Nothing,
                                     _Expire, Nothing, CacheItemPriority.Default, Nothing)
        End Sub

        'Public Function getCacheLifetime(_memberOf As List(Of Object)) As DateTime

        '    Dim ml As New MethodLogger("Ars.Cache")

        '    Dim lIntTtl As Integer = 0
        '    Dim lDatReturn As DateTime =
        '        DateAdd(DateInterval.Second, ArsApplicationSettings.adUserCacheSeconds, Now())

        '    If _memberOf IsNot Nothing Then

        '        For Each _member In _memberOf

        '            If _member.IndexOf("TTL") >= 0 Then

        '                Dim lObjMatch As Match =
        '                    Regex.Match(_member, "^<TTL=([0-9]+)>,.*$")

        '                If lObjMatch.Success Then

        '                    If ((lObjMatch.Groups(1).Value > 0 And lIntTtl = 0) Or (lObjMatch.Groups(1).Value < lIntTtl)) Then

        '                        lIntTtl = lObjMatch.Groups(1).Value
        '                    End If
        '                End If
        '            End If
        '        Next

        '        If lIntTtl > 0 Then

        '            lDatReturn = DateAdd(DateInterval.Second, lIntTtl, Now())
        '        End If
        '    End If

        '    ml.done(String.Format("expire = '{0}'", lDatReturn))

        '    Return lDatReturn
        'End Function

        Public Function getStringFromAttribute(lObjAttributeCollection As SearchResultAttributeCollection, lStrAttributeName As String) As String

            If lObjAttributeCollection.Contains(lStrAttributeName) Then

                Return lObjAttributeCollection(lStrAttributeName)(0)
            End If

            Return Nothing
        End Function

        Public Function getListFromAttribute(lObjAttributeCollection As SearchResultAttributeCollection, lStrAttributeName As String) As List(Of Object)

            If lObjAttributeCollection.Contains(lStrAttributeName) Then

                Return lObjAttributeCollection(lStrAttributeName).GetValues(GetType(String)).ToList()
            End If

            Return (New List(Of Object))
        End Function

        Public Function getLongFromAttribute(lObjAttributeCollection As SearchResultAttributeCollection, lStrAttributeName As String) As Long

            If lObjAttributeCollection.Contains(lStrAttributeName) Then

                Dim lLngReturn As Long = Nothing

                If Int64.TryParse(lObjAttributeCollection(lStrAttributeName)(0), lLngReturn) Then

                    Return lLngReturn
                End If
            End If

            Return Nothing
        End Function

        Public Sub UpdateUserFromCache(ByRef _CachedUser As CachedUser, ByRef _AssignedUser As AssignedUser)

            With _AssignedUser
                .Name = _CachedUser.name
                .DisplayName = _CachedUser.displayname
                .Department = _CachedUser.department
                .Company = _CachedUser.company
                .UPN = _CachedUser.upn
                .UAC = _CachedUser.uac
                .DN = _CachedUser.dn
                .Info = _CachedUser.info
                .Description = _CachedUser.description
                .memberOf = _CachedUser.memberOf
                .TimeStamp = _CachedUser.TimeStamp
                .pwdLastSet = _CachedUser.pwdLastSet
                .lockoutTime = _CachedUser.lockoutTime
            End With
        End Sub

        Public Sub UpdateUserFromLdap(ByRef _Attributes As SearchResultAttributeCollection, ByRef _AssignedUser As AssignedUser)

            With _AssignedUser
                .Name = getStringFromAttribute(_Attributes, "name")
                .DisplayName = getStringFromAttribute(_Attributes, "displayname")
                .Department = getStringFromAttribute(_Attributes, "department")
                .Company = getStringFromAttribute(_Attributes, "company")
                .UAC = getLongFromAttribute(_Attributes, "userAccountControl")
                .UPN = getStringFromAttribute(_Attributes, "UserPrincipalName")
                .DN = getStringFromAttribute(_Attributes, "distinguishedname")
                .Info = getStringFromAttribute(_Attributes, "info")
                .Description = getStringFromAttribute(_Attributes, "description")
                .memberOf = getListFromAttribute(_Attributes, "memberof")
                .pwdLastSet = getLongFromAttribute(_Attributes, "pwdLastSet")
                .TimeStamp = DateTime.UtcNow
                .lockoutTime = getLongFromAttribute(_Attributes, "lockoutTime")
            End With
        End Sub

        Public Sub UpdateGroupFromCache(ByRef _CachedGroup As CachedGroup, ByRef _AssignedGroup As AssignedGroup)

            With _AssignedGroup
                .Name = _CachedGroup.name
                .DN = _CachedGroup.dn
                .Info = _CachedGroup.info
                .Description = _CachedGroup.description
                .managedBy = _CachedGroup.managedBy
                .managedObjects = _CachedGroup.managedObjects
                .TimeStamp = _CachedGroup.TimeStamp
            End With
        End Sub

        Public Sub UpdateGroupFromLdap(ByRef _Attributes As SearchResultAttributeCollection, ByRef _AssignedGroup As AssignedGroup)

            With _AssignedGroup
                .Name = getStringFromAttribute(_Attributes, "name")
                .DN = getStringFromAttribute(_Attributes, "distinguishedname")
                .Info = getStringFromAttribute(_Attributes, "info")
                .Description = getStringFromAttribute(_Attributes, "description")
                .managedBy = getStringFromAttribute(_Attributes, "managedBy")
                .managedObjects = getListFromAttribute(_Attributes, "managedObjects")
                .TimeStamp = DateTime.UtcNow
            End With
        End Sub

    End Class
End Namespace