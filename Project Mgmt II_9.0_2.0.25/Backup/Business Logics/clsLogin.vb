Public Class clsLogin
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oColumn As SAPbouiCOM.Column
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath As String
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Private MatrixId1 As String
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource
#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm(ByVal aChoice As String)
        If aChoice = "Exp" Then
            EntryChoice = "Exp"
        ElseIf aChoice = "Reports" Then
            EntryChoice = "Reports"
        ElseIf aChoice = "TimeApproval" Then
            EntryChoice = "TimeApproval"
        ElseIf aChoice = "ExpApproval" Then
            EntryChoice = "ExpApproval"
        ElseIf aChoice = "LeaveApproval" Then
            EntryChoice = "LeaveApproval"
        ElseIf aChoice = "Posting" Then
            EntryChoice = "Posting"
        ElseIf aChoice = "Leave Request" Then
            EntryChoice = "Leave Request"
        ElseIf aChoice = "Acct" Then
            EntryChoice = "Acct"
        Else

            EntryChoice = "Timesheet"
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Login, frm_Login)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        databind(oForm)
        Dim blnLoginFlag, blnEntryFlag As Boolean
        blnEntryFlag = False
        blnLoginFlag = False
        If EntryChoice = "Reports" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("Reports", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True

        ElseIf EntryChoice = "Leave Request" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("Leave", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
        ElseIf EntryChoice = "ExpApproval" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("ExpApproval", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
        ElseIf EntryChoice = "TimeApproval" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("TimeApproval", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
        ElseIf EntryChoice = "LeaveApproval" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("LeaveApproval", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
        ElseIf EntryChoice = "Timesheet" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("Timesheet", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
            blnEntryFlag = True
        ElseIf EntryChoice = "Exp" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("Exp", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
            blnEntryFlag = True
        ElseIf EntryChoice = "Posting" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("Posting", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
            blnEntryFlag = True
        ElseIf EntryChoice = "Acct" Then
            oCombobox = oForm.Items.Item("8").Specific
            oCombobox.Select("Acct", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("8").Enabled = False
            blnLoginFlag = True
            blnEntryFlag = True
        Else
            oForm.Items.Item("8").Enabled = True
            blnLoginFlag = True
            blnEntryFlag = True
        End If
        If blnSourceForm = True Then
            BinLoginDetails(oForm)
            blnSourceForm = False
            oForm.Freeze(False)
        Else
            If LogonLoginDetails(oForm) = True Then


                If blnEntryFlag = True Then
                    oCombobox = oForm.Items.Item("10").Specific
                    oCombobox.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                oForm.Freeze(False)
                If blnLoginFlag = True Then
                    oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            Else
                oForm.Freeze(False)
            End If
        End If


    End Sub

#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            '' Adding 2 CFL, one for the button and one for the edit text.
            'oCFLCreationParams.MultiSelection = False
            'oCFLCreationParams.ObjectType = "Z_Expances"
            'oCFLCreationParams.UniqueID = "CFL1"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Z_Status"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "Bind Login Details"
    Private Sub BinLoginDetails(ByVal aForm As SAPbouiCOM.Form)
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select U_UID,U_Pwd, U_EmpiD,isnull(U_Approver,'N') from [@Z_Login] where U_EmpID='" & strSourceformEmpID & "'")
        If otemp.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(aForm, "3", otemp.Fields.Item(0).Value)
            oApplication.Utilities.setEdittextvalue(aForm, "5", otemp.Fields.Item(1).Value)
        End If



    End Sub
    Private Function LogonLoginDetails(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select U_UID,U_Pwd, U_EmpiD,isnull(U_Approver,'N') from [@Z_Login] where U_UID='" & oApplication.Company.UserName & "'")
        If otemp.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(aForm, "3", otemp.Fields.Item(0).Value)
            oApplication.Utilities.setEdittextvalue(aForm, "5", otemp.Fields.Item(1).Value)
            Return True
        Else
            oApplication.Utilities.setEdittextvalue(aForm, "3", "")
            oApplication.Utilities.setEdittextvalue(aForm, "5", "")
            Return False
        End If



    End Function
#End Region

#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim oCode As String
            aForm.Freeze(True)
            aForm.DataSources.UserDataSources.Add("Empid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("Pwd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("DocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("ViewType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("dtDate", SAPbouiCOM.BoDataType.dt_DATE)
            oApplication.Utilities.setUserDatabind(aForm, "3", "Empid")
            oApplication.Utilities.setUserDatabind(aForm, "5", "Pwd")
            oApplication.Utilities.setUserDatabind(aForm, "12", "dtDate")
            oCombobox = aForm.Items.Item("8").Specific
            oCombobox.ValidValues.Add("", "")
            oCombobox.ValidValues.Add("Exp", "Expenses")
            oCombobox.ValidValues.Add("Timesheet", "Time Sheet")
            oCombobox.ValidValues.Add("Reports", "Reports")
            oCombobox.ValidValues.Add("TimeApproval", "Time sheet Approval")
            oCombobox.ValidValues.Add("ExpApproval", "Expenses Approval")
            oCombobox.ValidValues.Add("LeaveApproval", "Leave Approval")
            oCombobox.ValidValues.Add("Posting", "Account Posting ")
            oCombobox.ValidValues.Add("Leave", "Leave Request")
            oCombobox.ValidValues.Add("Acct", "Expenses Accounting")

            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            aForm.Items.Item("8").DisplayDesc = True
            oCombobox = aForm.Items.Item("10").Specific
            oCombobox.ValidValues.Add("", "")
            oCombobox.ValidValues.Add("A", "Add/Edit")
            oCombobox.ValidValues.Add("V", "View")
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            aForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Items.Item("10").DisplayDesc = True
            aForm.Items.Item("9").Visible = False
            aForm.Items.Item("10").Visible = False
            aForm.Items.Item("11").Visible = False
            aForm.Items.Item("12").Visible = False
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim UID, pwd As String
        Dim otemp As SAPbobsCOM.Recordset
        Dim strAction, strdate As String
        Dim dtdate As Date
        UID = oApplication.Utilities.getEdittextvalue(aform, "3")
        pwd = oApplication.Utilities.getEdittextvalue(aform, "5")

        oCombobox = aform.Items.Item("8").Specific
        EntryChoice = oCombobox.Selected.Value
        If EntryChoice = "" Then
            oApplication.Utilities.Message("Select the Document type", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        If EntryChoice = "TimeApproval" Then
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select U_EmpiD,isnull(U_Approver,'N'),isnull(U_Superuser,'N') from [@Z_Login] where U_UID='" & UID & "' and U_PWD='" & pwd & "'")
            If otemp.RecordCount <= 0 Then
                oApplication.Utilities.Message("Invalid login details....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                If otemp.Fields.Item(1).Value <> "Y" And otemp.Fields.Item(2).Value <> "Y" Then
                    oApplication.Utilities.Message("You are not authorized to perform this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    Dim objct As New clsPrjTime
                    strApprovalType = "Time"
                    objct.LoadForm()
                    Return True
                End If
            End If

        End If

        If EntryChoice = "ExpApproval" Then
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select U_EmpiD,isnull(U_ExpApprover,'N') ,isnull(U_Superuser,'N') from [@Z_Login] where U_UID='" & UID & "' and U_PWD='" & pwd & "'")
            If otemp.RecordCount <= 0 Then
                oApplication.Utilities.Message("Invalid login details....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                If otemp.Fields.Item(1).Value <> "Y" And otemp.Fields.Item(2).Value <> "Y" Then
                    oApplication.Utilities.Message("You are not authorized to perform this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    Dim objct As New clsPrjTime
                    strApprovalType = "Exp"
                    objct.LoadForm()

                    Return True
                End If
            End If

        End If

        If EntryChoice = "LeaveApproval" Then
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select U_EmpiD,isnull(U_LEAVAPPROVER,'N'),isnull(U_Superuser,'N') from [@Z_Login] where U_UID='" & UID & "' and U_PWD='" & pwd & "'")
            If otemp.RecordCount <= 0 Then
                oApplication.Utilities.Message("Invalid login details....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                If otemp.Fields.Item(1).Value <> "Y" And otemp.Fields.Item(2).Value <> "Y" Then
                    oApplication.Utilities.Message("You are not authorized to perform this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    Dim objct As New clsPrjTime
                    strApprovalType = "Leave"
                    objct.LoadForm()
                    Return True
                End If
            End If

        End If

        oCombobox = aform.Items.Item("10").Specific
        strAction = oCombobox.Selected.Value
        If strAction = "" And EntryChoice <> "Reports" And EntryChoice <> "Leave" And EntryChoice <> "Posting" Then
            oApplication.Utilities.Message("Action missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        'If strAction = "V" Then
        If EntryChoice <> "Reports" And EntryChoice <> "Posting" And EntryChoice <> "Leave" Then
            'strdate = oApplication.Utilities.getEdittextvalue(aform, "12")
            strdate = Now.Date
            If strdate = "" Then
                oApplication.Utilities.Message("Document date is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                Try
                    dtdate = oApplication.Utilities.GetDateTimeValue(strdate)
                    ' dtdate = CDate(strdate)
                Catch ex As Exception
                    dtdate = oApplication.Utilities.GetDateTimeValue(strdate)
                End Try
            End If
        End If
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select U_EmpiD,isnull(U_Superuser,'N') from [@Z_Login] where U_UID='" & UID & "' and U_PWD='" & pwd & "'")
        If otemp.RecordCount > 0 Then
            loginEmployeeID = otemp.Fields.Item("U_EMPID").Value
            If EntryChoice = "Exp" Then
                Dim objct As New clsExpaneEntry
                If strAction = "V" Then
                    objct.LoadForm(loginEmployeeID, strAction, dtdate)
                Else
                    'objct.LoadForm(loginEmployeeID, strAction)
                    objct.LoadForm(loginEmployeeID, strAction, dtdate)
                End If
                Return True
            ElseIf EntryChoice = "Reports" Then
                Dim objct As New clsReports
                If otemp.Fields.Item(1).Value = "Y" Then
                    'If oApplication.Utilities.validateSuperuser() = True Then
                    objct.LoadForm(loginEmployeeID, "Super")
                Else
                    objct.LoadForm(loginEmployeeID, "")
                End If
                Return True
            ElseIf EntryChoice = "Leave" Then
                Dim objct As New clsLeaverequest
                If otemp.Fields.Item(1).Value = "Y" Then
                    'If oApplication.Utilities.validateSuperuser() = True Then
                    objct.LoadForm(loginEmployeeID, "Super")
                Else
                    objct.LoadForm(loginEmployeeID, "")
                End If
                Return True

            ElseIf EntryChoice = "Posting" Then
                Dim objct As New clsReports
                If otemp.Fields.Item(1).Value = "Y" Then
                    'If oApplication.Utilities.validateSuperuser() = True Then
                    objct.LoadForm(loginEmployeeID, "Super")
                Else
                    oApplication.Utilities.Message("You are not authorized to access this screen", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    ' objct.LoadForm(loginEmployeeID, "")
                End If
                Return True

            ElseIf EntryChoice = "Acct" Then
                Dim objct As New clsExpaneEntry_Account
                If otemp.Fields.Item(1).Value = "Y" Then
                    If strAction = "V" Then
                        objct.LoadForm(loginEmployeeID, strAction, dtdate)
                    Else
                        objct.LoadForm(loginEmployeeID, strAction, dtdate)
                    End If
                Else
                    oApplication.Utilities.Message("You are not authorized to access this screen", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    ' objct.LoadForm(loginEmployeeID, "")
                End If
                Return True
            Else
                Dim objct As New clsEmpTimeSheet
                If strAction = "V" Then
                    objct.LoadForm(loginEmployeeID, strAction, dtdate)
                Else
                    objct.LoadForm(loginEmployeeID, strAction, dtdate)
                End If
                Return True
            End If
        Else
            oApplication.Utilities.Message("Invalid login details....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
    End Function
#End Region
#End Region

#Region "enable Controls"
    Private Sub enableControls(ByVal aform As SAPbouiCOM.Form)
        oCombobox = aform.Items.Item("8").Specific
        If oCombobox.Selected.Value <> "Reports" And oCombobox.Selected.Value <> "Leave" And oCombobox.Selected.Value <> "Posting" And oCombobox.Selected.Value.Contains("Approval") = False Then
            aform.Items.Item("9").Visible = True
            aform.Items.Item("10").Visible = True
        Else
            aform.Items.Item("9").Visible = False
            aform.Items.Item("10").Visible = False
            aform.Items.Item("11").Visible = False
            aform.Items.Item("12").Visible = False
        End If
    End Sub
#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
        End Try
    End Sub
#End Region


    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Login Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                        End Select


                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "6" Then
                                    If validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oForm.Close()
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "10" Then
                                    oCombobox = oForm.Items.Item("8").Specific
                                    If oCombobox.Selected.Value <> "Reports" And oCombobox.Selected.Value <> "Posting" And oCombobox.Selected.Value <> "Approval" Then
                                        oCombobox = oForm.Items.Item("10").Specific
                                        If oCombobox.Selected.Value = "V" Then
                                            oForm.Items.Item("11").Visible = False
                                            oForm.Items.Item("12").Visible = False
                                        Else
                                            oForm.Items.Item("11").Visible = False
                                            oForm.Items.Item("12").Visible = False
                                            ' oForm.Items.Item("11").Visible = True
                                            ' oForm.Items.Item("12").Visible = True
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "8" Then
                                    enableControls(oForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "12" And pVal.ColUID = "V_0" Then
                                            val = oDataTable.GetValue("U_Z_EXPNAME", 0)
                                            oMatrix = oForm.Items.Item("12").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val)
                                        End If
                                        If pVal.ItemUID = "4" Then
                                            val = oDataTable.GetValue("firstName", 0)
                                            val1 = oDataTable.GetValue("empID", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", val)
                                            oApplication.Utilities.setEdittextvalue(oForm, "4", val1)

                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If

        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#End Region
End Class
