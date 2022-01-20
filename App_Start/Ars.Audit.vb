Imports System.Collections.Concurrent
Imports System.Data.SqlClient
Imports System.Security.Principal
Imports System.Threading
Imports System.Threading.Tasks

Namespace Ars
    Public Class AuditEntry

        Public ReadOnly _ts As String
        Public ReadOnly _sourceIp As String
        Public ReadOnly _sourceUser As String
        Public ReadOnly _assignedUser As AssignedUser
        Public ReadOnly _action As String
        Public ReadOnly _result As String

        Public Sub New(assignedUser As AssignedUser, action As String, result As String)

            If HttpContext.Current IsNot Nothing Then

                _sourceIp = HttpContext.Current.Request.UserHostAddress
                _sourceUser = HttpContext.Current.User.userName
            Else

                _sourceIp = "::1"
                _sourceUser = WindowsIdentity.GetCurrent.Name
            End If

            _assignedUser = assignedUser
            _action = action
            _result = result
        End Sub
    End Class

    Public Class AuditWriter

        Private _day As Integer = 1
        Private _logfile As String = Nothing

        Private SqlConnection As SqlConnection =
                New SqlConnection(ConfigurationManager.AppSettings.Item("sqlConnectionString"))

        Sub Write(auditEntries As AuditEntry())

            Dim lObjSqlCommand As New SqlCommand("", SqlConnection)

            Try

                Dim i As Integer = 0
                Dim lStrSqlKeys As String = "INSERT INTO [ars_audit] ( [uName], [aSID], [aName], [uSource], [action], [result], [ts]) VALUES "
                Dim lStrSqlValues As String = ""
                Dim lStrDelimiter As String = ""

                If Not SqlConnection.State = ConnectionState.Open Then

                    SqlConnection.Open()
                End If

                For Each entry As AuditEntry In auditEntries

                    i += 1

                    lStrSqlValues += String.Format("{6}('{0}','{1}','{2}','{3}','{4}','{5}',GetDate())",
                                                   entry._sourceUser, entry._assignedUser.SID, entry._assignedUser.Name,
                                                   entry._sourceIp, entry._action, entry._result, lStrDelimiter)

                    lStrDelimiter = ","
                    If i > 900 Then

                        lObjSqlCommand.CommandText = lStrSqlKeys & lStrSqlValues
                        lObjSqlCommand.ExecuteNonQuery()
                        lStrDelimiter = ""
                        Debug.Print(lObjSqlCommand.CommandText)
                    End If
                Next

                lObjSqlCommand.CommandText = lStrSqlKeys & lStrSqlValues
                lObjSqlCommand.ExecuteNonQuery()

                lObjSqlCommand.Dispose()
                SqlConnection.Close()
            Catch

                If SqlConnection.State = ConnectionState.Open Then

                    SqlConnection.Close()
                End If

                Ars.log.add(1, "AUDIT", "AUDIT", 0, lObjSqlCommand.CommandText)
                lObjSqlCommand.Dispose()
            End Try
        End Sub
    End Class

    Public Class Auditor

        Private _task As Task = Task.Run(AddressOf WriteLoop)
        Private _writer As New AuditWriter()
        Private _runnning As Boolean = True
        Private _queue As New ConcurrentQueue(Of AuditEntry)
        Private _cancel As New CancellationTokenSource()

        Public Sub Dispose()

            _runnning = False
            _cancel.Cancel()
            _task.Wait()
            WriteOut()
        End Sub

        Public Sub add(_assignedUser As AssignedUser, _action As String, _result As String)

            _queue.Enqueue(New AuditEntry(_assignedUser, _action, _result))
        End Sub

        Private Async Function WriteLoop() As Task

            While _runnning

                Try

                    Await Task.Delay(1000, _cancel.Token)
                    WriteOut()
                Catch
                End Try
            End While
        End Function

        Private Sub WriteOut()

            Dim entry As AuditEntry = Nothing

            If Not _queue.TryPeek(entry) Then Return

            Dim _entries As New List(Of AuditEntry)

            While _queue.TryDequeue(entry)

                _entries.Add(entry)
            End While

            _writer.Write(_entries.ToArray())
        End Sub
    End Class

End Namespace