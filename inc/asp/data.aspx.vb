Imports System.Security.Principal
Imports Newtonsoft.Json

Public Class data

    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Request.Headers("X-Requested-With") = "XMLHttpRequest" And Not Request.Headers("X-Referrer") = "ccArs" Then

            Response.AppendHeader("Refresh", "0; url=/")
            ClientScript.RegisterStartupScript(Me.GetType(), "JSScript", "<script language=JavaScript>window.location = '/';</script>")

            Response.End()
            Exit Sub
        End If

        Response.ContentType = "application/json"
        Response.Charset = "utf-8"

        Response.Cache.SetExpires(DateTime.UtcNow.AddDays(-1))
        Response.Cache.SetValidUntilExpires(False)
        Response.Cache.SetRevalidation(HttpCacheRevalidation.AllCaches)
        Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)
        Response.Cache.SetNoStore()

        If User.arsIsAuthenticated And Session("ArsUser") IsNot Nothing Then

            Dim ml As New Ars.MethodLogger("data.aspx")

            Dim _appSettings As Ars.ApplicationSettings =
                Application("ApplicationSettings")

            ml.write(Ars.LOGLEVEL.DEBUG,
                     String.Format("AD Authenticated: {0}, Azure Authenticated: {1}",
                                    User.onpremAuthenticated, User.azureAuthenticated))

            If Session("ArsUser").IsAllowed Then

                Session("ArsUser").SetHttpContext(Context)

                If Request.Form.AllKeys.Contains("e5892bf7-e9c3-4466-845a-1ebca87753f7") Then
                    'GET LOGGED ON USER INFO (vArsUser0)
                    Response.Write(
                        JsonConvert.SerializeObject(
                            Session("ArsUser"), Formatting.None, New JsonSerializerSettings With {
                                .ContractResolver = New Ars.DynamicContractResolver("user_0")
                            }
                        ))

                ElseIf Request.Form.AllKeys.Contains("9f35e966-3c8d-43ec-923f-1346246183f9") Then
                    'GET FULL USER INFO
                    Response.Write(
                        JsonConvert.SerializeObject(
                            Session("ArsUser"), Formatting.None, New JsonSerializerSettings With {
                                .ContractResolver = New Ars.DynamicContractResolver("user_2")
                            }
                        ))

                ElseIf Request.Form.AllKeys.Contains("380b841c-e14d-49f5-bdf2-0ee08d492215") Then
                    'FORCE FULL USER INFO 

                    Response.Write(
                        JsonConvert.SerializeObject(
                            Session("ArsUser").ForceAssignmentRefresh(), Formatting.None, New JsonSerializerSettings With {
                                .ContractResolver = New Ars.DynamicContractResolver("user_2")
                            }
                        ))

                ElseIf Request.Form.AllKeys.Contains("0dc44518-4efc-4393-bb5c-85b4db04015f") Then
                    'GET USER INFO WITHOUT GROUPS 
                    Response.Write(
                        JsonConvert.SerializeObject(
                            Session("ArsUser"), Formatting.None, New JsonSerializerSettings With {
                                .ContractResolver = New Ars.DynamicContractResolver("user_1")
                            }
                        ))

                ElseIf Request.Form.AllKeys.Contains("ceb052c2-355e-40fe-9877-ff6a15705703") Then
                    'FORCE USER INFO WITHOUT GROUPS 
                    Response.Write(
                        JsonConvert.SerializeObject(
                            Session("ArsUser").ForceAssignmentRefresh(), Formatting.None, New JsonSerializerSettings With {
                                .ContractResolver = New Ars.DynamicContractResolver("user_1")
                            }
                        ))

                ElseIf Request.Form.AllKeys.Contains("528e8218-b284-4475-a59e-9bad1f180208") Then
                    'REQUEST MEMBERSHIP OR EXTEND

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Add group membership.")

                    Dim lDicResult As New Dictionary(Of String, Dictionary(Of String, Object))
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("528e8218-b284-4475-a59e-9bad1f180208"))

                    If lStrAssignmentId IsNot Nothing Then

                        If Request.Form.AllKeys.Contains("t") Then

                            If IsNumeric(Request.Form.AllKeys.Contains("t")) Then

                                Dim lIntTimeEntry As Integer =
                                    Request.Form("t")

                                If lIntTimeEntry > 0 And lIntTimeEntry <= _appSettings.AssignmentHours.Length Then

                                    lDicResult.Add(lStrAssignmentId, New Dictionary(Of String, Object))

                                    Session("ArsUser").AddMemberToGroup(
                                        lStrAssignmentId, lDicResult(lStrAssignmentId),
                                        _appSettings.AssignmentHours(lIntTimeEntry - 1))

                                    ml.write(Ars.LOGLEVEL.INFO,
                                             String.Format("Assignment: {0}, AddMemberToGroup: {1}",
                                                           lStrAssignmentId,
                                                           lDicResult(lStrAssignmentId)("AddMemberToGroup")("result").ToString()))

                                End If
                            End If
                        End If
                    End If

                    Response.Write(
                        JsonConvert.SerializeObject(lDicResult))

                ElseIf Request.Form.AllKeys.Contains("cc5ddb45-c324-4208-a22a-be6f809edcd9") Then
                    'REMOVE MEMBERSHIP

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Remove group membership.")

                    Dim lDicResult As New Dictionary(Of String, Dictionary(Of String, Object))
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("cc5ddb45-c324-4208-a22a-be6f809edcd9"))

                    If lStrAssignmentId IsNot Nothing Then

                        lDicResult.Add(lStrAssignmentId, New Dictionary(Of String, Object))
                        Session("ArsUser").RemoveMemberFromGroup(lStrAssignmentId, lDicResult(lStrAssignmentId))
                        ml.write(Ars.LOGLEVEL.INFO,
                                 String.Format("Assignment: {0}, RemoveMemberFromGroup: {1}",
                                               lStrAssignmentId,
                                               lDicResult(lStrAssignmentId)("RemoveMemberFromGroup")("result").ToString()))
                    End If

                    Response.Write(
                        JsonConvert.SerializeObject(lDicResult))

                ElseIf Request.Form.AllKeys.Contains("ba9b9529-3be4-4c03-ae33-e2deca05260a[]") Then
                    'REQUEST BASE GROUPS

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Add base group memberships.")

                    Dim lDicResult As New Dictionary(Of String, Dictionary(Of String, Object))

                    For Each lStrAssignmentId In Request.Form.GetValues("ba9b9529-3be4-4c03-ae33-e2deca05260a[]")

                        Dim lIntAssignmentResult As Integer = 0

                        lStrAssignmentId =
                            Ars.Ars.arsValidateGuid(lStrAssignmentId)

                        If lStrAssignmentId IsNot Nothing Then

                            If Request.Form.AllKeys.Contains("t") Then

                                If IsNumeric(Request.Form("t")) Then

                                    Dim lIntTimeEntry As Integer =
                                        Request.Form("t")

                                    If lIntTimeEntry > 0 And lIntTimeEntry <= _appSettings.AssignmentHours.Length Then

                                        lDicResult.Add(lStrAssignmentId, New Dictionary(Of String, Object))

                                        Session("ArsUser").AddMemberToGroup(
                                            lStrAssignmentId, lDicResult(lStrAssignmentId),
                                            _appSettings.AssignmentHours(lIntTimeEntry - 1))

                                        ml.write(Ars.LOGLEVEL.INFO,
                                                 String.Format("Assignment: {0}, AddMemberToGroup: {1}",
                                                               lStrAssignmentId,
                                                               lDicResult(lStrAssignmentId)("AddMemberToGroup")("result").ToString()))
                                    End If
                                End If
                            End If
                        End If
                    Next

                    Response.Write(
                        JsonConvert.SerializeObject(lDicResult))

                ElseIf Request.Form.AllKeys.Contains("642e2ac6-c179-44d0-9d56-1fe60abfc098[]") Then
                    'REMOVE BASE GROUPS

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Remove base group memberships.")

                    Dim lDicResult As New Dictionary(Of String, Dictionary(Of String, Object))

                    For Each lStrAssignmentId In Request.Form.GetValues("642e2ac6-c179-44d0-9d56-1fe60abfc098[]")

                        Dim lIntAssignmentResult As Integer = 0

                        lStrAssignmentId =
                            Ars.Ars.arsValidateGuid(lStrAssignmentId)

                        If lStrAssignmentId IsNot Nothing Then

                            lDicResult.Add(lStrAssignmentId, New Dictionary(Of String, Object))

                            Session("ArsUser").removeMemberFromGroup(lStrAssignmentId, lDicResult(lStrAssignmentId))
                            ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}, RemoveMemberFromGroup_ResultCode: {1}",
                                                   lStrAssignmentId,
                                                   lDicResult(lStrAssignmentId)("RemoveMemberFromGroup")("result").ToString()))
                        End If
                    Next

                    Response.Write(
                        JsonConvert.SerializeObject(lDicResult))

                ElseIf Request.Form.AllKeys.Contains("4d07fa2c-9eef-465d-930f-ea652aec8ca3[]") Then
                    'REMOVE ALL GROUPS

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Remove all group memberships.")

                    Dim lDicResult As New Dictionary(Of String, Dictionary(Of String, Object))

                    For Each lStrAssignmentId In Request.Form.GetValues("4d07fa2c-9eef-465d-930f-ea652aec8ca3[]")

                        Dim lIntAssignmentResult As Integer = 0

                        lStrAssignmentId =
                            Ars.Ars.arsValidateGuid(lStrAssignmentId)

                        If lStrAssignmentId IsNot Nothing Then

                            lDicResult.Add(lStrAssignmentId, New Dictionary(Of String, Object))
                            Session("ArsUser").removeMemberFromGroup(lStrAssignmentId, lDicResult(lStrAssignmentId))

                            ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}, RemoveMemberFromGroup_ResultCode: {1}",
                                                   lStrAssignmentId,
                                                   lDicResult(lStrAssignmentId)("RemoveMemberFromGroup")("result").ToString()))
                        End If
                    Next

                    Response.Write(
                        JsonConvert.SerializeObject(lDicResult))

                ElseIf Request.Form.AllKeys.Contains("413b596d-b0af-41a2-909f-0670a2771a32") Then
                    'TOGGLE OPTION AUTO DISABLE

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Toggle auto disable.")

                    Dim lDicResult As New Dictionary(Of String, String)
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("413b596d-b0af-41a2-909f-0670a2771a32"))

                    If lStrAssignmentId IsNot Nothing Then

                        lDicResult.Add(
                            lStrAssignmentId, Session("ArsUser").toggleAutoDisable(lStrAssignmentId).ToString())

                        ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}, result: {1}",
                                                   lStrAssignmentId, lDicResult(lStrAssignmentId)))

                        Response.Write(
                            JsonConvert.SerializeObject(lDicResult))
                    End If

                ElseIf Request.Form.AllKeys.Contains("b8a6b63b-905b-4fd4-82d6-83adb8b419be") Then
                    'TOGGLE OPTION AUTO PASSWORD

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Toggle auto password.")

                    Dim lDicResult As New Dictionary(Of String, String)
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("b8a6b63b-905b-4fd4-82d6-83adb8b419be"))

                    If lStrAssignmentId IsNot Nothing Then

                        lDicResult.Add(
                            lStrAssignmentId, Session("ArsUser").toggleAutoPassword(lStrAssignmentId).ToString())

                        ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}, result: {1}",
                                                   lStrAssignmentId, lDicResult(lStrAssignmentId)))

                        Response.Write(
                            JsonConvert.SerializeObject(lDicResult))
                    End If

                ElseIf Request.Form.AllKeys.Contains("927cbd15-c1e5-4f00-85ee-431f49914b16") Then
                    'RESET PASSWORD

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Reset password.")

                    Dim lDicResult As New Dictionary(Of String, String)
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("927cbd15-c1e5-4f00-85ee-431f49914b16"))

                    If lStrAssignmentId IsNot Nothing Then

                        lDicResult.Add(
                            lStrAssignmentId, Session("ArsUser").resetPassword(lStrAssignmentId))
                        lDicResult.Add("action", "popup")

                        ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}", lStrAssignmentId))

                        Response.Write(
                            JsonConvert.SerializeObject(lDicResult))
                    End If

                ElseIf Request.Form.AllKeys.Contains("d68b5138-626b-45be-8e31-9ed01ae20e37") Then
                    'ENABLE ACCOUNT

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Enable account.")

                    Dim lDicResult As New Dictionary(Of String, Dictionary(Of String, Object))
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("d68b5138-626b-45be-8e31-9ed01ae20e37"))

                    If lStrAssignmentId IsNot Nothing Then

                        lDicResult.Add(lStrAssignmentId, New Dictionary(Of String, Object))
                        Session("ArsUser").enableAssignedUser(lStrAssignmentId, lDicResult(lStrAssignmentId))

                        ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}, SetAccountUac_ResultCode: {1}",
                                                   lStrAssignmentId, lDicResult(lStrAssignmentId)("enableAssignedUser")("result").ToString()))

                        Response.Write(
                            JsonConvert.SerializeObject(lDicResult))
                    End If

                ElseIf Request.Form.AllKeys.Contains("218a74c3-c5ed-4e87-8a25-892dedd444dc") Then
                    'DISABLE ACCOUNT

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Disable account.")

                    Dim lDicResult As New Dictionary(Of String, Dictionary(Of String, Object))
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("218a74c3-c5ed-4e87-8a25-892dedd444dc"))

                    If lStrAssignmentId IsNot Nothing Then

                        lDicResult.Add(lStrAssignmentId, New Dictionary(Of String, Object))
                        Session("ArsUser").disableAssignedUser(lStrAssignmentId, lDicResult(lStrAssignmentId))

                        ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}, SetAccountUac_ResultCode: {1}",
                                                   lStrAssignmentId, lDicResult(lStrAssignmentId)("disableAssignedUser")("result").ToString()))

                        Response.Write(
                            JsonConvert.SerializeObject(lDicResult))
                    End If

                ElseIf Request.Form.AllKeys.Contains("156d94b8-8bb3-41ac-8951-8d0ede66e075") Then
                    'Account audit log

                    ml.write(Ars.LOGLEVEL.INFO,
                             "Request account audit log.")

                    Response.Write(
                        JsonConvert.SerializeObject(
                            Session("ArsUser").GetAuditLog(lObjFormdata:=Request.Form)))

                ElseIf Request.Form.AllKeys.Contains("8c9219ab-deab-4076-8be7-55f4e7dd71a1") Then
                    'SWITCH GROUP TYPE

                    ml.write(Ars.LOGLEVEL.INFO,
                         "Toggle base group flag.")

                    Dim lDicResult As New Dictionary(Of String, String)
                    Dim lStrAssignmentId As String =
                        Ars.Ars.arsValidateGuid(Request.Form("8c9219ab-deab-4076-8be7-55f4e7dd71a1"))

                    If lStrAssignmentId IsNot Nothing Then

                        If Request.Form.AllKeys.Contains("f") Then

                            Dim lStrFlag As String =
                                Request.Form("f")

                            If IsNumeric(lStrFlag) Then

                                Dim lIntFlag As Integer = 0

                                Dim lLstValidFlags =
                                    New List(Of Integer) From {
                                        Ars.ArsGroupType.BASE,
                                        Ars.ArsGroupType.OTIONAL
                                    }

                                If Int32.TryParse(lStrFlag, lIntFlag) And lLstValidFlags.Contains(lIntFlag) Then

                                    lDicResult.Add("state",
                                            Session("ArsUSer").ToggleBaseGroupFlag(lStrAssignmentId))

                                    ml.write(Ars.LOGLEVEL.INFO,
                                        String.Format("assignment: '{0}', flag: {1}, result: {2}",
                                                        lStrAssignmentId, lIntFlag, lDicResult("state")))
                                End If

                                Response.Write(
                                    JsonConvert.SerializeObject(lDicResult))
                            End If
                        End If
                    End If
                End If

                If Session("ArsUser").IsAdmin Then
                    'ALL FUNCTIONS WITHIN THIS IF..THEN ARE
                    'ONLY ACCESSIBLE BY ADMIN ACCOUNTS

                    If Request.Form.AllKeys.Contains("ebbe0df1-2bcd-4251-895e-aea33092716e") Then
                        'SEARCH AD

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Search AD.")

                        Dim lStrClasses As String = Request.Form("ebbe0df1-2bcd-4251-895e-aea33092716e")

                        If IsNumeric(lStrClasses) Then

                            Dim lIntClasses As Integer = 0

                            If Int32.TryParse(lStrClasses, lIntClasses) And lIntClasses > 0 And lIntClasses < 4 Then

                                If Request.Form.AllKeys.Contains("s") Then

                                    Dim lLstClasses As New List(Of String)

                                    Dim lStrSearchString As String =
                                        Request.Form("s").Replace("=", "").Replace("<", "").
                                            Replace(">", "").Replace("&", "").Replace("|", "").Replace("*", "")

                                    lStrSearchString =
                                        Regex.Replace(lStrSearchString, "[ ](?=[ ])|[^-_,A-Za-z0-9 \(\)]+", "")

                                    If lStrSearchString.Length > 32 Then

                                        lStrSearchString = lStrSearchString.Substring(0, 32)
                                    End If

                                    If lStrSearchString.Length > 2 Then

                                        If lIntClasses.arsIsBitSet(1) Then

                                            lLstClasses.Add("user")
                                        End If

                                        If lIntClasses.arsIsBitSet(2) Then

                                            lLstClasses.Add("group")
                                        End If


                                        ml.write(Ars.LOGLEVEL.INFO,
                                             String.Format("classes: {0}, search: {1}",
                                                           String.Join(",", lLstClasses), lStrSearchString))

                                        Response.Write(
                                            JsonConvert.SerializeObject(
                                                Session("ArsUser").Admin.searchAd(lLstClasses, lStrSearchString)))
                                    End If
                                End If
                            End If
                        End If

                    ElseIf Request.Form.AllKeys.Contains("1f87eaa6-d780-4281-a429-856672746f19") Then
                        'GET DETAILD ADMIN USER INFO

                        'THIS STORES A INSTANCE OF USER CLASS FOR THE SELECTED USER / GROUP 
                        'IN SESSION VARIABLE 'ArsAdmin'
                        'THIS CAN BE USED IN FOLLOWUP REQUESTS

                        Dim lStrObjectSid As New SecurityIdentifier(
                            Request.Form("1f87eaa6-d780-4281-a429-856672746f19"))

                        If lStrObjectSid.IsAccountSid Then

                            If Request.Form.AllKeys.Contains("t") Then

                                Dim lStrClass As String =
                                    Request.Form("t")

                                If IsNumeric(lStrClass) Then

                                    Dim lIntClass As Integer = 0

                                    If Int32.TryParse(lStrClass, lIntClass) And lIntClass > 0 And lIntClass < 3 Then

                                        If lIntClass = 1 Then

                                            Session("ArsAdmin") =
                                                New Ars.User(lStrObjectSid.ToString(), Context, True)

                                            Response.Write(
                                                    JsonConvert.SerializeObject(
                                                        Session("ArsAdmin"), Formatting.None, New JsonSerializerSettings With {
                                                            .ContractResolver = New Ars.DynamicContractResolver("user_3")
                                                        }
                                                    ))
                                        ElseIf lIntClass = 2 Then

                                            Dim _Groups As New List(Of Ars.AssignedGroup)

                                            _Groups.Add(New Ars.AssignedGroup With {
                                                    .SID = lStrObjectSid.ToString(),
                                                    .User = Session("ArsUser")
                                                })

                                            Session("ArsUser").ldap.LoadLdapGroup(_Groups)

                                            Session("ArsAdmin") = _Groups(0)

                                            Response.Write(
                                                JsonConvert.SerializeObject(
                                                    Session("ArsAdmin"), Formatting.None, New JsonSerializerSettings With {
                                                        .ContractResolver = New Ars.DynamicContractResolver("group_0")
                                                    }
                                                ))
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    ElseIf Request.Form.AllKeys.Contains("13178d9e-1e6a-4fd3-bcd5-e13fdd8e986b") Then
                        'ADD ASSIGNMENT 
                        'USES THE 'ArsAdmin' USER INSTANCE

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Add Assignment.")

                        If Session("ArsAdmin") IsNot Nothing Then

                            Dim lObjAR As Ars.AssignmentRequest
                            Dim lDicResult As New Dictionary(Of String, String)

                            If Session("AssignmentRequest") Is Nothing Or Not String.IsNullOrEmpty(Request.Form("13178d9e-1e6a-4fd3-bcd5-e13fdd8e986b")) Then

                                lObjAR =
                                    New Ars.AssignmentRequest(
                                        Request.Form("13178d9e-1e6a-4fd3-bcd5-e13fdd8e986b"), Request.Form("t"))

                                ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Create new assignment request: sid = '{0}', type = '{1}'",
                                                   Request.Form("13178d9e-1e6a-4fd3-bcd5-e13fdd8e986b"), Request.Form("t")))
                            Else

                                lObjAR = Session("AssignmentRequest")
                                lObjAR.State = 1

                                ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Use existing assignment request: sid = '{0}', type = '{1}'",
                                                   lObjAR.SID, lObjAR.nType))
                            End If

                            If lObjAR.IsValid() Then

                                Session("AssignmentRequest") = lObjAR

                                If Session("ArsAdmin").GetType() = GetType(Ars.User) Then

                                    If lObjAR.nType = 1 Then
                                        'ADD USER AS TO MANAGE ACCOUNT

                                        Dim lIntAccountStatus =
                                            Session("ArsUser").Admin.GetAccountStatus(lObjAR.SID)

                                        If (lIntAccountStatus < 0 Or (lIntAccountStatus And 1) = 0) And lObjAR.State = 0 Then
                                            'USER NOT EXIST IN DB OR NOT ACTIVE

                                            lDicResult.Add("state", -1)

                                            ml.write(Ars.LOGLEVEL.INFO,
                                                 String.Format("Ask to add or enable account."))
                                        Else

                                            lDicResult.Add("state",
                                                           Session("ArsUser").Admin.AddUserAssignment(
                                                                lObjAR.SID, Session("ArsAdmin").SID))

                                            ml.write(Ars.LOGLEVEL.INFO,
                                                 String.Format("Add account, set '{0}' as manager for '{1}'",
                                                               lObjAR.SID, Session("ArsAdmin").SID))
                                        End If
                                    ElseIf lObjAR.nType = 2 Then
                                        'ADD USER TO BE MANAGED

                                        lDicResult.Add("state",
                                                       Session("ArsUser").Admin.AddUserAssignment(
                                                            Session("ArsAdmin").SID, lObjAR.SID))

                                        ml.write(Ars.LOGLEVEL.INFO,
                                             String.Format("Add account assignment, make '{0}' manageable by '{1}'",
                                                           Session("ArsAdmin").SID, lObjAR.SID))
                                    ElseIf lObjAR.nType = 3 Then
                                        'ADD GROUP TO BE REQUESTABLE

                                        lDicResult.Add("state",
                                                       Session("ArsUser").Admin.AddGroupAssignment(
                                                            Session("ArsAdmin").SID, lObjAR.SID))

                                        ml.write(Ars.LOGLEVEL.INFO,
                                             String.Format("Add group assignment, make '{0}' requestable for '{1}'",
                                                            lObjAR.SID, Session("ArsAdmin").SID))
                                    End If

                                ElseIf Session("ArsAdmin").GetType() = GetType(Ars.AssignedGroup) Then

                                    lDicResult.Add("state",
                                                   Session("ArsUser").Admin.AddGroupAssignment(
                                                        lObjAR.SID, Session("ArsAdmin").SID))

                                    ml.write(Ars.LOGLEVEL.INFO,
                                         String.Format("Add group assignment, make '{0}' requestable for '{1}'",
                                                        lObjAR.SID, Session("ArsAdmin").SID))
                                End If

                                Response.Write(
                                    JsonConvert.SerializeObject(lDicResult))
                            End If
                        End If

                    ElseIf Request.Form.AllKeys.Contains("ff90792a-4c63-445d-b981-ee80e6263891") Then
                        'REMOVE ASSIGNMENT

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Remove Assignment.")

                        Dim lDicResult As New Dictionary(Of String, String)
                        Dim lStrAssignmentId As String =
                            Ars.Ars.arsValidateGuid(Request.Form("ff90792a-4c63-445d-b981-ee80e6263891"))

                        If lStrAssignmentId IsNot Nothing Then

                            lDicResult.Add("state",
                                            Session("ArsUser").Admin.RemoveAssignment(
                                                    lStrAssignmentId))

                            ml.write(Ars.LOGLEVEL.INFO,
                                         String.Format("Assignment: '{0}'", lStrAssignmentId))

                            Response.Write(
                                JsonConvert.SerializeObject(lDicResult))
                        End If

                    ElseIf Request.Form.AllKeys.Contains("a6dc5917-0837-4eb8-bede-2c150969b627") Then
                        'LIST ALL ACCOUNTS

                        Response.Write(
                                JsonConvert.SerializeObject(Session("ArsUser").Admin.GetAccountList(),
                                                            Formatting.None, New JsonSerializerSettings With {
                                                                .ContractResolver = New Ars.DynamicContractResolver("user_4")
                                                            }))

                    ElseIf Request.Form.AllKeys.Contains("c5fe4461-3c40-4c5c-a61a-8afbde66485b") Then
                        'REMOVE USER

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Remove User.")

                        Dim lDicResult As New Dictionary(Of String, String)
                        Dim lStrObjectSid As New SecurityIdentifier(
                            Request.Form("c5fe4461-3c40-4c5c-a61a-8afbde66485b"))

                        If lStrObjectSid.IsAccountSid Then

                            lDicResult.Add("state",
                                            Session("ArsUser").Admin.RemoveUser(
                                                    lStrObjectSid.ToString()))

                            ml.write(Ars.LOGLEVEL.INFO,
                                         String.Format("sid: '{0}'", lStrObjectSid.ToString()))

                            Response.Write(
                                JsonConvert.SerializeObject(lDicResult))
                        End If

                    ElseIf Request.Form.AllKeys.Contains("2db55222-a29e-4656-ac82-e63fe34a2ea9") Then
                        'ADD USER

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Add User.")

                        Dim lDicResult As New Dictionary(Of String, String)
                        Dim lStrObjectSid As New SecurityIdentifier(
                            Request.Form("2db55222-a29e-4656-ac82-e63fe34a2ea9"))

                        If lStrObjectSid.IsAccountSid Then

                            lDicResult.Add("state",
                                            Session("ArsUser").Admin.AddUser(
                                                    lStrObjectSid.ToString()))

                            ml.write(Ars.LOGLEVEL.INFO,
                                         String.Format("sid: '{0}'", lStrObjectSid.ToString()))

                            Response.Write(
                                JsonConvert.SerializeObject(lDicResult))
                        End If

                    ElseIf Request.Form.AllKeys.Contains("0bd9da76-7ed3-4eb9-a49c-0b4ba4b1bc14") Then
                        'TOGGLE USER FLAG (ADMIN / ACTIVE)

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Toggle flag.")

                        Dim lDicResult As New Dictionary(Of String, String)
                        Dim lStrObjectSid As New SecurityIdentifier(
                            Request.Form("0bd9da76-7ed3-4eb9-a49c-0b4ba4b1bc14"))

                        If lStrObjectSid.IsAccountSid Then

                            If Request.Form.AllKeys.Contains("f") Then

                                Dim lStrFlag As String =
                                    Request.Form("f")

                                If IsNumeric(lStrFlag) Then

                                    Dim lIntFlag As Integer = 0

                                    If Int32.TryParse(lStrFlag, lIntFlag) And lIntFlag > 0 And lIntFlag < 3 Then

                                        If lIntFlag = 1 Then

                                            lDicResult.Add("state",
                                                            Session("ArsUser").Admin.ToggleUserFlag(
                                                                lStrObjectSid.ToString(), Ars.USER_FLAG.ADMIN))

                                            ml.write(Ars.LOGLEVEL.INFO,
                                                        String.Format("sid: '{0}', flag: {1}, result: {2}",
                                                            lStrObjectSid.ToString(), Ars.USER_FLAG.ADMIN, lDicResult("state")))

                                            Response.Write(
                                                JsonConvert.SerializeObject(lDicResult))
                                        ElseIf lIntFlag = 2 Then

                                            lDicResult.Add("state",
                                                            Session("ArsUser").Admin.ToggleUserFlag(
                                                                lStrObjectSid.ToString(), Ars.USER_FLAG.ACTIVE))

                                            ml.write(Ars.LOGLEVEL.INFO,
                                                        String.Format("sid: '{0}', flag: {1}, result: {2}",
                                                             lStrObjectSid.ToString(), Ars.USER_FLAG.ADMIN, lDicResult("state")))

                                            Response.Write(
                                                JsonConvert.SerializeObject(lDicResult))
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    ElseIf Request.Form.AllKeys.Contains("f3e7708f-e9b4-4219-8557-d5a730bbf70c") Then
                        'TOGGLE ASSIGNMENT FLAG

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Toggle assignment flag.")

                        Dim lDicResult As New Dictionary(Of String, String)
                        Dim lStrAssignmentId As String =
                            Ars.Ars.arsValidateGuid(Request.Form("f3e7708f-e9b4-4219-8557-d5a730bbf70c"))

                        If lStrAssignmentId IsNot Nothing Then

                            If Request.Form.AllKeys.Contains("f") Then

                                Dim lStrFlag As String =
                                    Request.Form("f")

                                If IsNumeric(lStrFlag) Then

                                    Dim lIntFlag As Integer = 0

                                    Dim lLstValidFlags =
                                        New List(Of Integer) From {
                                            Ars.ASS_FLAG.USER_ASSIGNMENT_AUTO_DISABLE_ACTIVE,
                                            Ars.ASS_FLAG.USER_ASSIGNMENT_AUTO_DISABLE_LOCKED,
                                            Ars.ASS_FLAG.USER_ASSIGNMENT_AUTO_PASSWORD_ACTIVE,
                                            Ars.ASS_FLAG.USER_ASSIGNMENT_AUTO_PASSWORD_LOCKED,
                                            Ars.ASS_FLAG.USER_ASSIGNMENT_MANUAL_DISABLE_LOCKED,
                                            Ars.ASS_FLAG.USER_ASSIGNMENT_RESET_PASSORD_LOCKED
                                        }

                                    If Int32.TryParse(lStrFlag, lIntFlag) And lLstValidFlags.Contains(lIntFlag) Then

                                        lDicResult.Add("state",
                                                Session("ArsAdmin").Admin.ToggleAssignmentFlag(
                                                        lStrAssignmentId, lIntFlag))

                                        ml.write(Ars.LOGLEVEL.INFO,
                                            String.Format("assignment: '{0}', flag: {1}, result: {2}",
                                                            lStrAssignmentId, lIntFlag, lDicResult("state")))
                                    End If

                                    Response.Write(
                                        JsonConvert.SerializeObject(lDicResult))
                                End If
                            End If
                        End If

                    ElseIf Request.Form.AllKeys.Contains("772005f4-edb5-4fb8-a723-2dc00af2cc09") Then
                        'ADMIN RESET PASSWORD

                        'USES THE 'ArsAdmin' USER INSTANCE

                        ml.write(Ars.LOGLEVEL.INFO,
                             "Admin reset password")

                        If Session("ArsAdmin") IsNot Nothing Then

                            Dim lDicResult As New Dictionary(Of String, String)
                            Dim lStrAssignmentId As String =
                                Ars.Ars.arsValidateGuid(Request.Form("772005f4-edb5-4fb8-a723-2dc00af2cc09"))

                            If lStrAssignmentId IsNot Nothing Then

                                lDicResult.Add(
                                    lStrAssignmentId, Session("ArsAdmin").resetPassword(lStrAssignmentId))
                                lDicResult.Add("action", "popup")

                                ml.write(Ars.LOGLEVEL.INFO,
                                     String.Format("Assignment: {0}", lStrAssignmentId))

                                Response.Write(
                                    JsonConvert.SerializeObject(lDicResult))
                            End If
                        End If

                    ElseIf Request.Form.AllKeys.Contains("cdbb6e5e-48df-419c-8f7d-9a4bc8df25b7") Then
                        'Global audit log

                        ml.write(Ars.LOGLEVEL.INFO,
                                 "Request account audit log.")

                        Response.Write(
                            JsonConvert.SerializeObject(
                                Session("ArsUser").GetAuditLog(Request.Form, 1)))
                    End If
                End If
            End If

            ml.done()
        Else

            Response.Write("{""ars"":-1}")
        End If
    End Sub

End Class