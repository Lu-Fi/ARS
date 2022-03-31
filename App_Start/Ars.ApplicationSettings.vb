Imports System.Linq

Namespace Ars
    Public Class ApplicationSettings

        Public logLevel As Integer = 0

        Public SearchResultLimit As Integer = 8

        Public LdapServer As String = Nothing
        Public LdapPort As Integer = 389
        Public IsPersistentLdapServer As Boolean = False

        Public SearchBase As String = Nothing
        Public GroupSearchBase As String = Nothing
        Public UserSearchBase As String = Nothing

        Public pwdLength As Integer = 8
        Public pwdSpecialCharacters As Integer = 2
        Public adUserCacheSeconds As Integer = 1
        Public adGroupCacheSeconds As Integer = 1

        Public SqlConnectionString As String =
            ConfigurationManager.AppSettings.Item("sqlConnectionString")
        Public SqlRefreshSeconds As Integer = 15
        Public AdAssignmentRefreshSeconds As Integer = 60

        Public BackgroundWorkerSeconds As Integer = 60
        Public OrphanSidRemovalDays As Integer = 1

        Public LdapGroupAttributes As String() =
            {"name", "description", "info", "distinguishedname", "objectSid", "managedBy", "managedObjects"}

        Public LdapUserAttributes As String() =
            {"memberOf", "name", "info", "description", "department", "company", "distinguishedname", "displayname", "objectSid", "userAccountControl", "UserPrincipalName", "pwdLastSet", "allowedAttributesEffective", "lockoutTime"}

        Public LdapMaxSearchEntries As Integer = 1000

        Public AssignmentHours() As Integer = {900, 3600, 7200, 14400, 28800}

        Public AccessGroupRegex As String = Nothing

        Sub New()

            If ConfigurationManager.AppSettings.AllKeys.Contains("app_accessgroup_regex") Then

                If ConfigurationManager.AppSettings.Item("app_accessgroup_regex").Length > 5 Then

                    AccessGroupRegex = ConfigurationManager.AppSettings.Item("app_accessgroup_regex")
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("app_logLevel") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("app_logLevel")) Then

                    If ConfigurationManager.AppSettings.Item("app_logLevel") > logLevel Then

                        logLevel = ConfigurationManager.AppSettings.Item("app_logLevel")
                    End If
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("app_BackgroundWorkerSeconds") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("app_BackgroundWorkerSeconds")) Then

                    If ConfigurationManager.AppSettings.Item("app_BackgroundWorkerSeconds") > BackgroundWorkerSeconds Then

                        BackgroundWorkerSeconds = ConfigurationManager.AppSettings.Item("app_BackgroundWorkerSeconds")
                    End If
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("app_OrphanSidRemovalDays") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("app_OrphanSidRemovalDays")) Then

                    If ConfigurationManager.AppSettings.Item("app_OrphanSidRemovalDays") > OrphanSidRemovalDays Then

                        OrphanSidRemovalDays = ConfigurationManager.AppSettings.Item("app_OrphanSidRemovalDays")
                    End If
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("sql_UserRefreshSeconds") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("sql_UserRefreshSeconds")) Then

                    If ConfigurationManager.AppSettings.Item("sql_UserRefreshSeconds") > SqlRefreshSeconds Then

                        SqlRefreshSeconds = ConfigurationManager.AppSettings.Item("sql_UserRefreshSeconds")
                    End If
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("Ad_AssignmentRefreshSeconds") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("Ad_AssignmentRefreshSeconds")) Then

                    If ConfigurationManager.AppSettings.Item("Ad_AssignmentRefreshSeconds") > AdAssignmentRefreshSeconds Then

                        AdAssignmentRefreshSeconds = ConfigurationManager.AppSettings.Item("Ad_AssignmentRefreshSeconds")
                    End If
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("ldap_UserCacheSeconds") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("ldap_UserCacheSeconds")) Then

                    If ConfigurationManager.AppSettings.Item("ldap_UserCacheSeconds") > adUserCacheSeconds Then

                        adUserCacheSeconds = ConfigurationManager.AppSettings.Item("ldap_UserCacheSeconds")
                    End If
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("ldap_GroupCacheSeconds") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("ldap_GroupCacheSeconds")) Then

                    If ConfigurationManager.AppSettings.Item("ldap_GroupCacheSeconds") > adGroupCacheSeconds Then

                        adGroupCacheSeconds = ConfigurationManager.AppSettings.Item("ldap_GroupCacheSeconds")
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

                        SearchResultLimit = ConfigurationManager.AppSettings.Item("ldap_SearchLimit")
                    End If
                End If
            End If

            If ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Then

                If IsNumeric(ConfigurationManager.AppSettings.Item("ldap_Port")) Then

                    LdapPort = ConfigurationManager.AppSettings.Item("ldap_Port")
                End If
            End If

            If Not ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Or Not ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Or Not ConfigurationManager.AppSettings.AllKeys.Contains("ldap_Port") Then

                Exit Sub
            Else

                LdapServer = ConfigurationManager.AppSettings.Item("ldap_Server")
                SearchBase = ConfigurationManager.AppSettings.Item("ldap_SearchBase")
                GroupSearchBase = ConfigurationManager.AppSettings.Item("ldap_GroupSearchBase")
                UserSearchBase = ConfigurationManager.AppSettings.Item("ldap_UserSearchBase")
            End If
        End Sub
    End Class
End Namespace

