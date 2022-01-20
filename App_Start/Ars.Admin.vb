Imports System.Data.SqlClient
Imports System.Security.Principal
Imports System.DirectoryServices.Protocols

Namespace Ars

    Public Class Admin

        Private ReadOnly User As User

        Public Sub New(_User As User)

            User = _User
        End Sub

        Public Function searchAd(lLstrClasses As List(Of String), lStrSearchString As String) As List(Of Dictionary(Of String, String))

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("classes = '{0}', search = '{1}'", String.Join(",", lLstrClasses), lStrSearchString))

            Dim lDicReturn As New List(Of Dictionary(Of String, String))

            If User.IsAdmin Then

                Dim lObjLdapConn As LdapConnection =
                    User.ldap.GetLdapConnection()

                If lObjLdapConn IsNot Nothing Then

                    Try

                        Dim lStrFilter As String = Nothing

                        For Each lStrClass In lLstrClasses

                            lStrFilter +=
                                String.Format("(objectClass={0})", lStrClass)
                        Next

                        lStrFilter = String.Format("(&(|{0})(anr={1}*))", lStrFilter, lStrSearchString)

                        Dim lLstUserAttributes As String() =
                            {"name", "description", "displayname", "company", "department", "objectSid", "objectClass", "UserPrincipalName"}

                        Dim lObjSearchRequest As New SearchRequest(
                            User.ApplicationSettings.UserSearchBase,
                        lStrFilter,
                        System.DirectoryServices.Protocols.SearchScope.Subtree,
                        lLstUserAttributes)

                        lObjSearchRequest.SizeLimit =
                            User.ApplicationSettings.SearchResultLimit

                        Dim lObjResponse As SearchResponse =
                        lObjLdapConn.SendRequest(lObjSearchRequest)

                        If lObjResponse.Entries.Count > 0 Then

                            For Each lObjEntry As SearchResultEntry In lObjResponse.Entries

                                Dim lObjResultEntry As New Dictionary(Of String, String) From {
                                    {"sid", New SecurityIdentifier(lObjEntry.Attributes("objectSid")(0), 0).ToString()},
                                    {"description", ""},
                                    {"type", "2"},
                                    {"name", lObjEntry.Attributes("name")(0)}
                                }

                                If lObjEntry.Attributes.Contains("description") Then

                                    lObjResultEntry("description") =
                                        lObjEntry.Attributes("description")(0)
                                End If

                                If Not User.ldap.Cache.getListFromAttribute(lObjEntry.Attributes, "objectclass").Contains("group") Then

                                    lObjResultEntry("type") = "1"

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

                    End Try
                End If
            End If

            ml.done()

            Return lDicReturn

        End Function

        Public Function GetAccountStatus(_SID As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("SID = '{0}'", _SID))

            If Not User.IsAdmin Then

                ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
                Return -99
            End If

            Dim lIntReturn As Integer = -1

            Try

                If User.SID.Length > 0 Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "SELECT TOP(1) * FROM [ars_users] WHERE [uSID] = @uSID",
                            User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("uSID", _SID)

                    Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

                    While lObjReader.Read()

                        lIntReturn = lObjReader.GetInt32(1)
                    End While

                    lObjReader.Close()
                    User.SqlConnection.Close()
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR GetAccountStatus, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return lIntReturn
        End Function

        Public Function AddUserAssignment(_uSID As String, _aSID As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("uSID = '{0}', aSID = '{1}'", _uSID, _aSID))

            If Not User.IsAdmin Then

                ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
                Return -99
            End If

            Try

                If _uSID.Length > 0 And _aSID.Length > 0 Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "IF NOT EXISTS (SELECT uSID FROM [ars_users] WHERE uSID = @uSID)
	                            INSERT INTO [ars_users] ([uSID], [flag]) VALUES (@uSID, '1')
                             ELSE
	                            UPDATE [ars_users] SET [flag] = [flag] | 1 WHERE uSID = @uSID;

                             IF NOT EXISTS (SELECT uSID FROM [ars_assignments] WHERE uSID = @uSID AND aSID = @aSID)
	                            INSERT INTO [ars_assignments] ([uSID], [aSID], [flag]) VALUES (@uSID, @aSID, '1')
                             ELSE
	                            UPDATE [ars_assignments] SET [flag] = [flag] | 1 WHERE uSID = @uSID AND aSID = @aSID;
                            ",
                            User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)
                    lObjSqlCommand.Parameters.AddWithValue("aSID", _aSID)

                    lObjSqlCommand.ExecuteNonQuery()

                    User.SqlConnection.Close()

                    User.SqlTimestamp =
                        DateAdd(DateInterval.Day, -1, Now)

                    Return 1
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR AddUserAssignment, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return 0
        End Function

        Public Function AddGroupAssignment(_uSID As String, _aSID As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("uSID = '{0}', aSID = '{1}'", _uSID, _aSID))

            If Not User.IsAdmin Then

                ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
                Return -99
            End If

            Try

                If _uSID.Length > 0 And _aSID.Length > 0 Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "IF NOT EXISTS (SELECT uSID FROM [ars_assignments] WHERE uSID = @uSID AND aSID = @aSID)
	                            INSERT INTO [ars_assignments] ([uSID], [aSID], [flag]) VALUES (@uSID, @aSID, '2')
                             ELSE
	                            UPDATE [ars_assignments] SET [flag] = [flag] | 2 WHERE uSID = @uSID AND aSID = @aSID;
                            ",
                            User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)
                    lObjSqlCommand.Parameters.AddWithValue("aSID", _aSID)

                    lObjSqlCommand.ExecuteNonQuery()

                    User.SqlConnection.Close()

                    User.SqlTimestamp =
                        DateAdd(DateInterval.Day, -1, Now)

                    Return 1
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR AddGroupAssignment, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return 0
        End Function

        Public Function RemoveAssignment(_aID As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("Assignment = '{0}'", _aID))

            If Not User.IsAdmin Then

                ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
                Return -99
            End If

            Try

                If Not String.IsNullOrEmpty(_aID) Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "DELETE FROM [ars_assignments] WHERE aID = @aID;",
                            User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("aID", _aID)

                    lObjSqlCommand.ExecuteNonQuery()

                    User.SqlConnection.Close()

                    User.SqlTimestamp =
                        DateAdd(DateInterval.Day, -1, Now)

                    Return 1
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR RemoveAssignment, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return 0
        End Function

        Public Function GetAccountList() As List(Of AssignedUser)

            If Not User.IsAdmin Then

                Return Nothing
            End If

            Dim _Accounts As New List(Of AssignedUser)

            Try

                If User.SID.Length > 0 Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "SELECT * FROM [ars_users]",
                            User.SqlConnection)

                    Dim lObjReader As SqlDataReader = lObjSqlCommand.ExecuteReader()

                    While lObjReader.Read()

                        _Accounts.Add(New AssignedUser With {
                            .SID = lObjReader.GetString(0),
                            .flags = lObjReader.GetInt32(1)
                        })
                    End While

                    lObjReader.Close()
                    User.SqlConnection.Close()
                End If
            Catch ex As Exception

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            User.ldap.LoadLdapUser(_Accounts)

            Return _Accounts
        End Function

        Public Function RemoveUser(_uSID As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("uSID = '{0}'", _uSID))

            If Not User.IsAdmin Then

                ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
                Return -99
            End If

            Try

                If Not String.IsNullOrEmpty(_uSID) Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "DELETE FROM [ars_users] WHERE uSID = @uSID;
                             DELETE FROM [ars_assignments] WHERE (SELECT TOP(1) [uSID] FROM [ars_users] WHERE uSID = [ars_assignments].[uSID]) IS NULL AND (flag & 1 = 1);
                             DELETE M FROM [ars_assignments] M WHERE (SELECT TOP(1) [uSID] FROM [ars_assignments] WHERE [aSID] = M.[uSID] AND ([flag] & 1 = 1)) IS NULL AND (M.[flag] & 2 = 2);",
                            User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)

                    lObjSqlCommand.ExecuteNonQuery()

                    User.SqlConnection.Close()

                    Return 1
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR RemoveUser, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return 0
        End Function

        Public Function AddUser(_uSID As String) As Integer

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("uSID = '{0}'", _uSID))

            If Not User.IsAdmin Then

                ml.write(LOGLEVEL.INFO, "!! EXIT !! GetAccountStatus, NO ADMIN RIGHTS !!")
                Return -99
            End If

            Try

                If Not String.IsNullOrEmpty(_uSID) Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                            "IF NOT EXISTS (SELECT uSID FROM [ars_users] WHERE uSID = @uSID)
	                            INSERT INTO [ars_users] ([uSID], [flag]) VALUES (@uSID, '1')",
                            User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)

                    lObjSqlCommand.ExecuteNonQuery()

                    User.SqlConnection.Close()

                    Return 1
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR AddUser, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return 0
        End Function

        Public Function ToggleUserFlag(_uSID As String, _flag As Integer) As Integer

            '------------------------------------------
            '
            '  Toggle User Admin / Active Flag
            '
            '------------------------------------------

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("_uSID = '{0}', flag = '{1}'", _uSID, _flag))

            Try

                If Not String.IsNullOrEmpty(_uSID) Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                                        "UPDATE TOP(1) [ars_users] SET flag = flag ^ @flag WHERE [uSID] = @uSID",
                                        User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("uSID", _uSID)
                    lObjSqlCommand.Parameters.AddWithValue("flag", _flag)

                    lObjSqlCommand.ExecuteNonQuery()

                    User.SqlConnection.Close()

                    User.SqlTimestamp =
                        DateAdd(DateInterval.Day, -1, Now)

                    Return 1
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR ToggleFlag, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return 0
        End Function

        Public Function ToggleAssignmentFlag(_aID As String, _flag As Integer) As Integer

            '------------------------------------------
            '
            '  Toggle Assignment Flags
            '
            '------------------------------------------

            Dim ml As New MethodLogger("ArsUser",
                                       String.Format("assignment = '{0}', flag = '{1}'", _aID, _flag))

            Try

                If Not String.IsNullOrEmpty(_aID) Then

                    If Not User.SqlConnection.State = ConnectionState.Open Then

                        User.SqlConnection.Open()
                    End If

                    Dim _Assignment As Assignment =
                        User._Assignments.FirstOrDefault(Function(e) e.aID.ToString() = _aID)

                    Dim lIntType As Integer = ASS_FLAG.USER_ASSIGNMENT

                    If _flag = ASS_FLAG.GROUP_ASSIGNMENT_BASE Then

                        lIntType = ASS_FLAG.GROUP_ASSIGNMENT
                    End If

                    Dim lObjSqlCommand As New SqlCommand(
                                        "If NOT EXISTS (Select aID FROM [ars_assignments] WHERE aID = @aID )
                                            INSERT INTO [ars_assignments] ([aID], [uSID], [aSID], [flag]) VALUES (@aID, @uSID, @aSID, (@flag | 4096 | @type))
                                         ELSE
                                            UPDATE TOP(1) [ars_assignments] SET flag = flag ^ @flag WHERE [aID] = @aID",
                                        User.SqlConnection)

                    lObjSqlCommand.Parameters.AddWithValue("aID", _Assignment.aID)
                    lObjSqlCommand.Parameters.AddWithValue("uSID", User.SID)
                    lObjSqlCommand.Parameters.AddWithValue("aSID", _Assignment.aSID)
                    lObjSqlCommand.Parameters.AddWithValue("type", lIntType)
                    lObjSqlCommand.Parameters.AddWithValue("flag", _flag)

                    lObjSqlCommand.ExecuteNonQuery()

                    User.SqlConnection.Close()

                    User.SqlTimestamp =
                        DateAdd(DateInterval.Day, -1, Now)
                    Return 1
                End If
            Catch ex As Exception

                ml.write(LOGLEVEL.ERR, String.Format("ERROR ToggleAssignmentFlag, {0}", ex.Message))

                If Not User.SqlConnection.State = ConnectionState.Closed Then

                    User.SqlConnection.Close()
                End If
            End Try

            ml.done()

            Return 0
        End Function
    End Class
End Namespace
