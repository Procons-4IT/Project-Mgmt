Public Class clsPrjTime
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath, strwhs, strGirdValue As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()

        'If oApplication.Utilities.validateSuperuser() = False Then
        '    oApplication.Utilities.Message("You are not authorized to perform this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Exit Sub
        'End If

        oForm = oApplication.Utilities.LoadForm(xml_Prjtime, frm_ProjectTime)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oCombobox = oForm.Items.Item("4").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow)
        Next
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("E", "Expenses")
        oCombobox.ValidValues.Add("T", "Time Sheet")
        oCombobox.ValidValues.Add("L", "Leave")
        If strApprovalType = "Time" Then
            oCombobox.Select("T", SAPbouiCOM.BoSearchKey.psk_ByValue)
        ElseIf strApprovalType = "Exp" Then
            oCombobox.Select("E", SAPbouiCOM.BoSearchKey.psk_ByValue)
        ElseIf strApprovalType = "Leave" Then
            oCombobox.Select("L", SAPbouiCOM.BoSearchKey.psk_ByValue)
        End If
        ' oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("4").DisplayDesc = True

        oCombobox = oForm.Items.Item("10").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow)
        Next
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("A", "Approved")
        oCombobox.ValidValues.Add("D", "Declined")
        oCombobox.ValidValues.Add("P", "Approval Pending")
        'oCombobox.TitleObject.Caption = "Approved"
        ' oCombobox.ValidValues.Add("V", "View Details")
        oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
        oForm.Items.Item("10").DisplayDesc = True
        oForm.Items.Item("4").Enabled = False
        oForm.DataSources.UserDataSources.Add("dtFrom", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("dtto", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("fromEMP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("ToEmp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "11", "dtFrom")
        oApplication.Utilities.setUserDatabind(oForm, "14", "dtto")
        oApplication.Utilities.setUserDatabind(oForm, "17", "fromEMP")
        oApplication.Utilities.setUserDatabind(oForm, "19", "ToEmp")
        oEditText = oForm.Items.Item("17").Specific
        oEditText.ChooseFromListUID = "CFL_2"
        oEditText.ChooseFromListAlias = "empID"
        oEditText = oForm.Items.Item("19").Specific
        oEditText.ChooseFromListUID = "CFL_3"
        oEditText.ChooseFromListAlias = "empID"
        oForm.Freeze(False)
    End Sub
#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As String
        Dim strReportType, strViewType, strFromDate, strToDate, strFromEMP, strToEMP, strSQL, strSQL1, strStatus, strempcondition, strCondidtion, strDateCondition As String
        Dim dtFromdate, dtTodate As Date
        oCombobox = aForm.Items.Item("4").Specific
        strReportType = oCombobox.Selected.Value
        oCombobox = aForm.Items.Item("10").Specific
        strViewType = oCombobox.Selected.Value
        strFromDate = oApplication.Utilities.getEdittextvalue(aForm, "11")
        strToDate = oApplication.Utilities.getEdittextvalue(aForm, "14")
        strFromEMP = oApplication.Utilities.getEdittextvalue(aForm, "17")
        strToEMP = oApplication.Utilities.getEdittextvalue(aForm, "19")
        If strReportType = "" Then
            oApplication.Utilities.Message("Document type missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        End If
        If strViewType = "" Then
            oApplication.Utilities.Message("Report type missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        End If
        Dim strDateCondition1 As String = ""
        If strFromDate <> "" And strToDate <> "" Then
            dtFromdate = oApplication.Utilities.GetDateTimeValue(strFromDate)
            dtTodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            'strDateCondition = " Convert(varchar(10),T0.[U_Z_DOCDATE],103) between '" & dtFromdate.ToString("dd/MM/yyyy") & "' and '" & dtTodate.ToString("dd/MM/yyyy") & "'"
            strDateCondition = " T0.[U_Z_DOCDATE] > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and   T0.[U_Z_DOCDATE] < ='" & dtTodate.ToString("yyyy-MM-dd") & "'"
            strDateCondition1 = " T0.[U_Z_DOCDATE] > = '" & dtFromdate.ToString("yyyy-dd-MM") & "' and   T0.[U_Z_DOCDATE] < ='" & dtTodate.ToString("yyyy-dd-MM") & "'"

        ElseIf strFromDate <> "" And strToDate = "" Then
            dtFromdate = oApplication.Utilities.GetDateTimeValue(strFromDate)
            dtTodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            strDateCondition = " T0.[U_Z_DOCDATE] > = '" & dtFromdate.ToString("yyyy-MM-dd") & "'"
            strDateCondition1 = " T0.[U_Z_DOCDATE] > = '" & dtFromdate.ToString("yyyy-dd-MM") & "'"
        ElseIf strFromDate = "" And strToDate <> "" Then
            dtFromdate = oApplication.Utilities.GetDateTimeValue(strFromDate)
            dtTodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            strDateCondition = " T0.[U_Z_DOCDATE] < = '" & dtTodate.ToString("yyyy-MM-dd") & "'"
            strDateCondition1 = " T0.[U_Z_DOCDATE] < = '" & dtTodate.ToString("yyyy-dd-MM") & "'"
        Else
            strDateCondition = " 1=1"
            strDateCondition1 = " 1=1"
        End If

        Dim oTestRs As SAPbobsCOM.Recordset
        Dim s As String
        Try
            s = "Select * from [@Z_OEXP] T0 where T0.U_Z_DocDate > ='2011-01-04' and T0.U_Z_DocDate  <='2011-30-04'"
            oTestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTestRs.DoQuery(s)
            strDateCondition = strDateCondition1
        Catch ex As Exception
            s = "Select * from [@Z_OTIM] T0 where " & strDateCondition
            oTestRs.DoQuery(s)
        End Try
        Dim strExpEmpCondition As String = ""


        If strFromEMP <> "" And strToEMP <> "" Then
            strempcondition = " T0.[U_Z_EMPCODE] between '" & strFromEMP & "' and '" & strToEMP & "'"
            strExpEmpCondition = " T1.[U_Z_EMPID] between '" & strFromEMP & "' and '" & strToEMP & "'"
        ElseIf strFromEMP <> "" And strToEMP = "" Then
            strempcondition = " T0.[U_Z_EMPCODE] >= '" & strFromEMP & "'"
            strExpEmpCondition = " T1.[U_Z_EMPID] >= '" & strFromEMP & "'"
        ElseIf strFromEMP = "" And strToEMP <> "" Then

            strempcondition = " T0.[U_Z_EMPCODE] <='" & strToEMP & "'"
            strExpEmpCondition = " T1.[U_Z_EMPID] <='" & strToEMP & "'"
        Else
            strempcondition = " 1=1"
            strExpEmpCondition = " 1=1"
        End If


        Dim strCon1 As String
        If strViewType = "A" Then
            If strReportType = "E" Then
                strCondidtion = " Where T1.U_Z_APPROVED='A' and " & strExpEmpCondition & " and " & strDateCondition
            Else
                strCondidtion = " Where T1.U_Z_APPROVED='A' and " & strempcondition & " and " & strDateCondition
            End If

            strCon1 = " Where T0.U_Z_APPROVED='A' and " & strempcondition & " and " & strDateCondition
        ElseIf strViewType = "D" Then
            If strReportType = "E" Then
                strCondidtion = " Where T1.U_Z_APPROVED='D' and " & strExpEmpCondition & " and " & strDateCondition
            Else
                strCondidtion = " Where T1.U_Z_APPROVED='D' and " & strempcondition & " and " & strDateCondition
            End If

            strCon1 = " Where T0.U_Z_APPROVED='D' and " & strempcondition & " and " & strDateCondition
        ElseIf strViewType = "P" Then
            If strReportType = "E" Then
                strCondidtion = " Where T1.U_Z_APPROVED='P' and " & strExpEmpCondition & " and " & strDateCondition
            Else
                strCondidtion = " Where T1.U_Z_APPROVED='P' and " & strempcondition & " and " & strDateCondition
            End If

            strCon1 = " Where T0.U_Z_APPROVED='P' and " & strempcondition & " and " & strDateCondition
        Else
            If strReportType = "E" Then
                strCondidtion = " Where " & strExpEmpCondition & " and " & strDateCondition
            Else
                strCondidtion = " Where " & strempcondition & " and " & strDateCondition
            End If

            strCon1 = " Where " & strempcondition & " and " & strDateCondition
        End If
        Dim strEMPIDS As String = oApplication.Utilities.getEmpIDforMangers(oApplication.Company.UserName)
        If strEMPIDS = "" Then
            strEMPIDS = "99999"
        End If
        strSQL = ""

        Dim blnSuperUser As Boolean
        Dim intUser As String = oApplication.Utilities.getLoggedonEmployee()
        If oApplication.Utilities.CheckSuperUser(intUser) = False Then
            blnSuperUser = False
        Else
            blnSuperUser = True
        End If

        Dim strProjCondition As String
        strProjCondition = "Select U_Z_PRJCODE from [@Z_HPRJ] where U_Z_EMPID='" & intUser & "'"
        strProjCondition = "Select PRJCODE from OPRJ where U_Z_EMPID='" & intUser & "'"

        Dim oTest, otest1 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select isnull(firstName,'') + ' ' + isnull(middleName,'')+' '+isnull(lastName,''),empID from OHEM ")
        For intLoop As Integer = 0 To oTest.RecordCount - 1
            otest1.DoQuery("Update [@Z_EXP1] set U_Z_EMPNAME='" & oTest.Fields.Item(0).Value & "' where U_Z_EMPID=" & oTest.Fields.Item(1).Value)
            otest1.DoQuery("Update [@Z_OTIM] set U_Z_EMPNAME='" & oTest.Fields.Item(0).Value & "' where U_Z_EMPCODE='" & oTest.Fields.Item(1).Value & "'")
            oTest.MoveNext()
        Next


        If blnSuperUser = True Then
            If strReportType = "E" Then
                strSQL = "SELECT T1.[U_Z_REFCODE],T1.[U_Z_EMPID] 'U_Z_EMPCODE', T1.[U_Z_EMPNAME], T0.[U_Z_DOCDATE], T1.[U_Z_EXPNAME], T1.[U_Z_EXPTYPE], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME], T1.[U_Z_CURRENCY], T1.[U_Z_AMOUNT],Convert(Decimal,(replace(T1.[U_Z_LocAmt],substring(T1.[U_Z_LocAmt],0,4),''))) 'U_Z_LOCAMT',Convert(Decimal,(replace(T1.[U_Z_SysAmt],substring(T1.[U_Z_SysAmt],0,4),''))) 'U_Z_SYSAMT', T1.[CODE],T1.[U_Z_Remarks], T1.[U_Z_DisRule],T1.[U_Z_APPROVED],T1.U_Z_DOCNUM FROM [dbo].[@Z_OEXP]  T0 "
                strSQL = strSQL & " inner join  [dbo].[@Z_EXP1]  T1 on T0.Code=T1.U_Z_RefCode " & strCondidtion & " order by T1.[U_Z_EMPID],T0.[U_Z_DOCDATE]"
            ElseIf strReportType = "T" Then
                strSQL = "SELECT T1.[U_Z_REFCODE],isnull(T0.U_Z_TYPE,'E') 'TType', isnull(T0.U_Z_CNTID,'E') 'U_Z_CNTID', T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T0.[U_Z_DOCDATE], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME], isnull(T1.[U_Z_TYPE],'I') 'U_Z_TYPE' ,T1.[U_Z_BDGQTY],T1.[U_Z_QUANTITY],T1.[U_Z_MEASURE], T1.[U_Z_DATE], T1.[U_Z_HOURS], T1.[CODE], T1.[U_Z_Remarks],T1.[U_Z_APPROVED] FROM [dbo].[@Z_OTIM]  T0 inner Join   [dbo].[@Z_TIM1]  T1 on T0.Code=T1.[U_Z_REFCODE] " & strCondidtion & " and isnull(T1.U_Z_EmpApproval,'P')='C'   order by T0.[U_Z_EMPCODE],T0.[U_Z_DOCDATE]"
            Else
                strSQL = "Select * from [@Z_OLEV] T0 " & strCon1
            End If
        Else
            If strReportType = "E" Then
                strSQL = "SELECT T1.[U_Z_REFCODE],T1.[U_Z_EMPID] 'U_Z_EMPCODE', T1.[U_Z_EMPNAME], T0.[U_Z_DOCDATE], T1.[U_Z_EXPNAME], T1.[U_Z_EXPTYPE], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME], T1.[U_Z_CURRENCY], T1.[U_Z_AMOUNT],Convert(Decimal,(replace(T1.[U_Z_LocAmt],substring(T1.[U_Z_LocAmt],0,4),''))) 'U_Z_LOCAMT',Convert(Decimal,(replace(T1.[U_Z_SysAmt],substring(T1.[U_Z_SysAmt],0,4),''))) 'U_Z_SYSAMT', T1.[CODE],T1.[U_Z_Remarks],T1.[U_Z_DisRule], T1.[U_Z_APPROVED],T1.U_Z_DOCNUM FROM [dbo].[@Z_OEXP]  T0 "
                strSQL = strSQL & " inner join  [dbo].[@Z_EXP1]  T1 on T0.Code=T1.U_Z_RefCode " & strCondidtion & " and T1.U_Z_EMPID in (" & strEMPIDS & ") and T1.U_Z_PRJCODE in (" & strProjCondition & ")  order by T1.[U_Z_EMPID],T0.[U_Z_DOCDATE]"
            ElseIf strReportType = "T" Then
                strSQL = "SELECT T1.[U_Z_REFCODE],isnull(T0.U_Z_TYPE,'E') 'TType', isnull(T0.U_Z_CNTID,'E') 'U_Z_CNTID', T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T0.[U_Z_DOCDATE], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME], isnull(T1.[U_Z_TYPE],'I') 'U_Z_TYPE' ,T1.[U_Z_BDGQTY],T1.[U_Z_QUANTITY],T1.[U_Z_MEASURE], T1.[U_Z_DATE], T1.[U_Z_HOURS], T1.[CODE], T1.[U_Z_Remarks],T1.[U_Z_APPROVED] FROM [dbo].[@Z_OTIM]  T0 inner Join   [dbo].[@Z_TIM1]  T1 on T0.Code=T1.[U_Z_REFCODE] " & strCondidtion & " and T0.[U_Z_EMPCODE] in (" & strEMPIDS & ")  and T1.U_Z_PRJCODE in (" & strProjCondition & ")  and isnull(T1.U_Z_EmpApproval,'P')='C'   order by T0.[U_Z_EMPCODE],T0.[U_Z_DOCDATE]"
            Else
                strSQL = "Select * from [@Z_OLEV] T0 " & strCon1 & "  and T0.U_Z_EMPCODE in (" & strEMPIDS & ")"
            End If
        End If

      
        Return strSQL
    End Function
#End Region

#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_PrjTime
                    If pVal.BeforeAction = False Then
                        Dim oTe As New clsLogin
                        oTe.LoadForm("TimeApproval")
                    End If
                Case mnu_ExpApproval
                    If pVal.BeforeAction = False Then
                        Dim oTe As New clsLogin
                        oTe.LoadForm("ExpApproval")
                    End If
                Case mnu_LeaveApproval
                    Dim oTe As New clsLogin
                    oTe.LoadForm("LeaveApproval")

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        'EnableControls(oForm)
                    End If

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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ProjectTime Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                
                        End Select


                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                ' oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                               
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "20"
                                        Dim stQuery As String
                                        stQuery = validation(oForm)
                                        If stQuery <> "" Then
                                            Dim oItem As New clsApproval
                                            Dim frmSource As SAPbouiCOM.Form
                                            frmSource = oForm
                                            oItem.LoadForm(stQuery, oForm)
                                            frmSource.Close()
                                        End If

                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

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
                                        If pVal.ItemUID = "17" Or pVal.ItemUID = "19" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
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
