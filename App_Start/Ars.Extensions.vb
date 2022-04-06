Imports System.Security.Claims
Imports System.Security.Principal
Imports System.Runtime.CompilerServices
Imports Newtonsoft.Json.Serialization
Imports Newtonsoft.Json

Module IntegerExtensions

    <Extension()>
    Public Function arsIsBitSet(ByVal lInt As Integer, ByVal lBit As Integer) As Boolean

        If (lInt And lBit) = lBit Then

            Return True
        Else

            Return False
        End If
    End Function
End Module

Module PrincipalExtension

    <Extension()>
    Function arsIsAuthenticated(lObjPrincipal As IPrincipal) As Boolean

        Dim lIntRequire As Integer = 2
        Dim lIntValue As Integer = 0

        If ConfigurationManager.AppSettings.AllKeys.Contains("auth_requireNegotiate") Then

            'Currently disabled, sometimes AD authentication is lost
            'this happens, because of azure Auth is set as Primary

            'lIntRequire = lIntRequire Or 1

            If lObjPrincipal.onpremAuthenticated() Then

                lIntValue = lIntValue Or 1
            Else

                lIntValue = lIntValue And Not 1
            End If
        End If

        If lObjPrincipal.azureAuthenticated() Then

            lIntValue = lIntValue Or 2
        Else

            lIntValue = lIntValue And Not 2
        End If

        Return CBool(lIntValue >= lIntRequire)
    End Function

    <Extension()>
    Function onpremAuthenticated(lObjPrincipal As IPrincipal) As Boolean

        If lObjPrincipal IsNot Nothing Then

            Dim lObjClaimsPrincipal As ClaimsPrincipal = lObjPrincipal

            Dim lObjNegotiateIdentities =
                    lObjClaimsPrincipal.Identities.Where(
                        Function(c) c.GetType() = GetType(System.Security.Principal.WindowsIdentity))

            For Each lObjNegotiateIdentity In lObjNegotiateIdentities

                If lObjNegotiateIdentity.IsAuthenticated Then

                    Return True
                Else

                    Return False
                    Exit For
                End If
            Next
        End If

        Return False
    End Function

    <Extension()>
    Function azureAuthenticated(lObjPrincipal As IPrincipal) As Boolean

        If lObjPrincipal IsNot Nothing Then

            Dim lObjClaimsPrincipal As ClaimsPrincipal = lObjPrincipal

            Dim lObjClaimsIdentities =
                lObjClaimsPrincipal.Identities.Where(
                    Function(c) c.GetType() = GetType(System.Security.Claims.ClaimsIdentity))

            For Each lObjClaimsIdentity In lObjClaimsIdentities

                If lObjClaimsIdentity.IsAuthenticated Then

                    Return True
                Else

                    Return False
                    Exit For
                End If
            Next
        End If

        Return False
    End Function

    <Extension()>
    Function azureAuthenticationExpireTimestamp(lObjPrincipal As IPrincipal) As Integer

        If lObjPrincipal IsNot Nothing Then

            Dim lObjClaimsPrincipal As ClaimsPrincipal = lObjPrincipal

            Dim lObjClaimsIdentities =
                    lObjClaimsPrincipal.Identities.Where(
                        Function(c) c.GetType() = GetType(System.Security.Claims.ClaimsIdentity))

            For Each lObjClaimsIdentity In lObjClaimsIdentities

                Return lObjClaimsIdentity.FindFirst("exp").Value
            Next
        End If

        Return 0
    End Function

    <Extension()>
    Function azureOnpremSid(lObjPrincipal As IPrincipal) As String

        If lObjPrincipal IsNot Nothing Then

            Dim lObjClaimsPrincipal As ClaimsPrincipal = lObjPrincipal

            Dim lObjClaimsIdentities =
                    lObjClaimsPrincipal.Identities.Where(
                        Function(c) c.GetType() = GetType(System.Security.Claims.ClaimsIdentity))

            For Each lObjClaimsIdentity In lObjClaimsIdentities

                Return lObjClaimsIdentity.FindFirst("onprem_sid").Value
            Next
        End If

        Return Nothing
    End Function

    <Extension()>
    Function userName(lObjPrincipal As IPrincipal) As String

        If lObjPrincipal IsNot Nothing Then

            Dim lObjClaimsPrincipal As ClaimsPrincipal = lObjPrincipal

            Dim lObjClaimsIdentities =
                    lObjClaimsPrincipal.Identities.Where(
                        Function(c) c.GetType() = GetType(System.Security.Claims.ClaimsIdentity))

            If lObjClaimsIdentities.Count > 0 Then

                If Not String.IsNullOrEmpty(lObjClaimsIdentities.First().Name) Then

                    Return lObjClaimsIdentities.First().Name
                End If
            End If
        End If
        Return "Anonymous" 'WindowsIdentity.GetCurrent.Name
    End Function
End Module

Namespace Ars

    Public Enum ASS_SOURCE
        SQL = 1
        LDAP = 2
    End Enum

    Public Enum ArsGroupType
        BASE = 1
        OTIONAL = 2
        ROLE = 3
    End Enum

    Public Enum ASS_FLAG
        'SETTINGS_ASSIGNMENT = NOT USER AND NOT GROUP
        USER_ASSIGNMENT = 1
        GROUP_ASSIGNMENT = 2

        USER_ASSIGNMENT_AUTO_DISABLE_ACTIVE = 4
        USER_ASSIGNMENT_AUTO_DISABLE_LOCKED = 8
        USER_ASSIGNMENT_AUTO_PASSWORD_ACTIVE = 16
        USER_ASSIGNMENT_AUTO_PASSWORD_LOCKED = 32
        USER_ASSIGNMENT_RESET_PASSORD_LOCKED = 64
        USER_ASSIGNMENT_MANUAL_DISABLE_LOCKED = 128

        GROUP_ASSIGNMENT_BASE = 1024

        SETTINGS_ASSIGNMENT = 4096
    End Enum

    Public Enum USER_FLAG
        ACTIVE = 1
        ADMIN = 256
        ASSIGNED_BY_GROUP = 512
    End Enum

    Public Enum LOGLEVEL
        ERR = 0
        INFO = 1
        VERBOSE = 2
        DEBUG = 3
    End Enum

    Module Ars

        Public log As New Logger(LOGLEVEL.INFO, "ars-")
        Public audit As New Auditor()

        Function arsValidateGuid(lStrGuid As String) As String

            Try

                If lStrGuid IsNot Nothing And lStrGuid.Length = 36 Then

                    Dim lObjRegex As New Regex("^({){0,1}[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}(}){0,1}$")

                    If lObjRegex.Match(lStrGuid).Success Then

                        Return lStrGuid
                    End If
                End If
            Catch ex As Exception

            End Try

            Return Nothing

        End Function

    End Module

    Public Class DynamicContractResolver : Inherits DefaultContractResolver

        Private _propertySets As New Dictionary(Of String, List(Of String)) From {
            {"user_0", New List(Of String) From {"User.Name", "User.DisplayName", "User.Company", "User.Department"}},
            {"user_1", New List(Of String) From {"User.Name", "User.DisplayName", "User.Company", "User.AssignedUsers", "User.aID", "User.Department",
                                                 "AssignedUser.Name", "AssignedUser.DisplayName", "AssignedUser.Description", "AssignedUser.flags", "AssignedUser.aID", "AssignedUser.AccountStatus", "AssignedUser.source"}},
            {"user_2", New List(Of String) From {"User.Name", "User.DisplayName", "User.Company", "User.AssignedUsers", "User.aID", "User.Department",
                                                 "AssignedUser.Name", "AssignedUser.DisplayName", "AssignedUser.Description", "AssignedUser.AssignedGroups", "AssignedUser.flags", "AssignedUser.aID", "AssignedUser.source",
                                                 "AssignedGroup.Name", "AssignedGroup.Description", "AssignedGroup.Type", "AssignedGroup.TypeString", "AssignedGroup.TTL", "AssignedGroup.Expire", "AssignedGroup.aID", "AssignedGroup.source"}},
            {"user_3", New List(Of String) From {"User.Name", "User.DisplayName", "User.Company", "User.AssignedUsers", "User.aID", "User.Department", "User.Users", "User.AssignedGroups",
                                                 "AssignedUser.Name", "AssignedUser.DisplayName", "AssignedUser.Description", "AssignedUser.aID", "AssignedUser.Company", "AssignedUser.Department", "AssignedUser.SID", "AssignedUser.flags", "AssignedUser.source",
                                                 "AssignedGroup.Name", "AssignedGroup.Description", "AssignedGroup.Type", "AssignedGroup.aID", "AssignedGroup.SID", "AssignedGroup.source"}},
            {"user_4", New List(Of String) From {"AssignedUser.Name", "AssignedUser.DisplayName", "AssignedUser.Company", "AssignedUser.Department", "AssignedUser.Description", "AssignedUser.flags", "AssignedUser.SID", "AssignedUser.source"}},
            {"group_0", New List(Of String) From {"AssignedGroup.Name", "AssignedGroup.Description", "AssignedGroup.Type", "AssignedGroup.Users", "AssignedGroup.SID", "AssignedGroup.source",
                                                  "AssignedUser.Name", "AssignedUser.DisplayName", "AssignedUser.Description", "AssignedUser.Company", "AssignedUser.Department", "AssignedUser.SID", "AssignedUser.source"}}
        }

        Private _propertySet As String

        Public Sub New(propertySet As String)

            _propertySet = propertySet
        End Sub

        Protected Overrides Function CreateProperties(type As Type, memberSerialization As MemberSerialization) As IList(Of JsonProperty)

            Dim properties As IList(Of JsonProperty) =
                MyBase.CreateProperties(type, memberSerialization)

            properties =
                properties.Where(Function(p) _propertySets(_propertySet).Contains(
                    String.Format("{0}.{1}", type.Name, p.PropertyName))).ToList()

            Return properties
        End Function
    End Class
End Namespace