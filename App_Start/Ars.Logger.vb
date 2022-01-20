Imports System.Collections.Concurrent
Imports System.Threading
Imports System.Threading.Tasks
Imports System.IO
Imports System.Security.Principal

Namespace Ars
    Public Class LogEntry

        Public ReadOnly _ts As String
        Public ReadOnly _source As String
        Public ReadOnly _user As String
        Public ReadOnly _id As Integer
        Public ReadOnly _message As String
        Public ReadOnly _level As Integer

        Public Sub New(level As Integer, source As String, user As String, id As Integer, message As String)

            _ts = DateTime.Now.ToString("HH:mm:ss:fff")
            _source = source
            _user = user
            _id = id
            _message = message
            _level = level
        End Sub
    End Class

    Public Class LogWriter

        Private _day As Integer = 1
        Private _logfile As String = Nothing
        Private _str_level As New List(Of String) From {"ERROR", "INFO ", "VRBO ", "DEBUG"}
        Private _prefix As String = Nothing

        Sub New(Optional prefix As String = Nothing)

            _prefix = prefix
        End Sub

        Private ReadOnly Property logfile As String
            Get

                If Not Directory.Exists(AppContext.BaseDirectory & "log") Then

                    Try

                        Directory.CreateDirectory(AppContext.BaseDirectory & "log")
                    Catch ex As Exception
                    End Try
                End If

                If _day <> Now().DayOfYear Then

                    _day = Now().DayOfYear
                    _logfile = String.Format("{0}log\{1}{2}.log", AppContext.BaseDirectory, _prefix, DateTime.Now.ToString("yyyy-MM-dd"))
                End If

                Return _logfile
            End Get
        End Property

        Sub Write(logEntries As LogEntry())

            Try

                Dim sb As New StringBuilder()

                For Each entry As LogEntry In logEntries

                    sb.AppendLine(
                        String.Format("{0,-12}  {1,-7}  {2,-5}  {3,-16}  {4}  {5}",
                                  entry._ts, entry._id, _str_level(entry._level), entry._source, entry._user, entry._message))
                Next

                'Debug.Print(sb.ToString())
                File.AppendAllText(logfile, sb.ToString())
            Catch

            End Try
        End Sub
    End Class

    Public Class Logger

        Private _task As Task = Task.Run(AddressOf WriteLoop)
        Private _level As Integer = 0
        Private _writer As LogWriter
        Private _runnning As Boolean = True
        Private _queue As New ConcurrentQueue(Of LogEntry)
        Private _cancel As New CancellationTokenSource()
        Private _prefix As String = Nothing
        Private _eventLog As New EventLog("Application")

        Sub New(Optional level As LOGLEVEL = 0, Optional prefix As String = Nothing)

            _level = level
            _prefix = prefix
            _writer = New LogWriter(_prefix)
        End Sub

        Public Sub Dispose()

            _runnning = False
            _cancel.Cancel()
            _task.Wait()
            WriteOut()
        End Sub

        Public Sub writeEvent(_message As String)

            _eventLog.Source = "Application"
            _eventLog.WriteEntry(_message, EventLogEntryType.Information, 99)
        End Sub

        Public Sub add(logLevel As Integer, logSource As String, logUser As String, logId As Integer, logMessage As String)

            If logLevel > _level Then Return
            _queue.Enqueue(New LogEntry(logLevel, logSource, logUser, logId, logMessage))
        End Sub

        Public Property LogLevel As LOGLEVEL
            Get
                Return _level
            End Get
            Set(value As LOGLEVEL)
                _level = value
            End Set
        End Property

        Private Async Function WriteLoop() As Task

            While _runnning

                Try

                    Await Task.Delay(250, _cancel.Token)
                    WriteOut()
                Catch
                End Try
            End While
        End Function

        Private Sub WriteOut()

            Dim entry As LogEntry = Nothing

            If Not _queue.TryPeek(entry) Then Return

            Dim _entries As New List(Of LogEntry)

            While _queue.TryDequeue(entry)

                _entries.Add(entry)
            End While

            _writer.Write(_entries.ToArray())
        End Sub
    End Class

    Public Class MethodLogger

        Dim _ts As DateTime = DateTime.Now()
        Dim _id As Integer = _ts.ToString("mmssfff")
        Dim _src As String = Nothing
        Dim _usr As String = Nothing
        Dim _caller As String = Nothing

        Sub New(source As String, Optional parameter As String = Nothing, <System.Runtime.CompilerServices.CallerMemberName> Optional callerName As String = Nothing)

            _src = source
            _caller = callerName

            Dim _context As HttpContext =
                HttpContext.Current

            If _context IsNot Nothing Then

                _usr = HttpContext.Current.User.userName
            Else

                _usr = WindowsIdentity.GetCurrent.Name
            End If

            write(LOGLEVEL.DEBUG, String.Format("START {0}({1})", _caller, parameter))
        End Sub

        Sub write(_level As LOGLEVEL, _message As String)

            log.add(_level, _src, _usr, _id, _message)
        End Sub

        Sub done(Optional results As String = Nothing)

            write(LOGLEVEL.DEBUG, String.Format("END {0} in {1}ms. {2}", _caller, (DateTime.Now() - _ts).TotalMilliseconds, results))
        End Sub

        Property UserName As String
            Get

                Return _usr
            End Get
            Set(value As String)

                _usr = value
            End Set
        End Property
    End Class
End Namespace