Public Class clsLeaverequest
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Grid
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBoxColumn
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

    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_LeaveReqest, frm_Leaverequest)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        'oForm.PaneLevel = 1
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        databind(oForm, Now.Date)
        oForm.Freeze(False)
    End Sub

    Public Sub LoadForm(ByVal aEmpid As String, ByVal aOption As String, Optional ByVal aDate As String = "")
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dtdate As Date
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select * from OHEM where empID='" & aEmpid & "'")
        If oTemp.RecordCount > 0 Then
            Dim strName As String
            strName = oTemp.Fields.Item("firstName").Value & " " & oTemp.Fields.Item("LastName").Value
            oForm = oApplication.Utilities.LoadForm(xml_LeaveReqest, frm_Leaverequest)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            ' AddChooseFromList(oForm)
            'oForm.PaneLevel = 1
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oApplication.Utilities.setEdittextvalue(oForm, "6", aEmpid)
            oApplication.Utilities.setEdittextvalue(oForm, "8", strName)
            If aOption = "A" Then
                databind(oForm, aDate)
            Else
                '  dtdate = aDate
                databind_ViewEntry(oForm, dtdate)
            End If
            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        End If
        oForm.Freeze(False)
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

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_Expances"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region


#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form, ByVal adate As Date)
        Try
            Dim oCode As String
            Dim strSQL As String
            Dim dtdate1 As Date
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            aForm.Freeze(True)
            strSQL = "Select * from [@Z_OEXP] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "'" ' and convert(varchar(10),U_Z_DocDate,105)='" & adate.ToString("dd-MM-yyyy") & "'"
            otemp.DoQuery(strSQL)
            oMatrix = aForm.Items.Item("9").Specific
            oMatrix.DataTable.ExecuteQuery(strSQL)
            FormatGrid(oMatrix, oApplication.Utilities.getEdittextvalue(aForm, "6"))
            'oColumn.ValidValues.Add("A", "Approved")
            'oColumn.ValidValues.Add("D", "Declined")
            'oColumn.ValidValues.Add("P", "Approval Pending")
            'oColumn.TitleObject.Caption = "Approval Status"
            'oColumn.DisplayDesc = True
            'oColumn = oMatrix.Columns.Item("V_6")
            '' LoadCurrency(oColumn)
            'oColumn.DisplayDesc = True

            
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub databind_ViewEntry(ByVal aForm As SAPbouiCOM.Form, ByVal adate As Date)
        Try
            Dim oCode As String
            Dim strSQL, strEmpID As String
            Dim dtdate1 As Date
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            aForm.Freeze(True)
            ' strEmpID = oApplication.Utilities.getEdittextvalue(aForm, "4")
            strSQL = "Select * from [@Z_OLEV] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "6") & "'" ' and convert(varchar(10),U_Z_DocDate,105)='" & adate.ToString("dd-MM-yyyy") & "'"
            otemp.DoQuery(strSQL)
            oMatrix = aForm.Items.Item("9").Specific
            oMatrix.DataTable.ExecuteQuery(strSQL)
            FormatGrid(oMatrix, oApplication.Utilities.getEdittextvalue(aForm, "6"))
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Format Grid"
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid, ByVal aempid As String)
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        aGrid.Columns.Item(0).Visible = False
        aGrid.Columns.Item(1).Visible = False
        aGrid.Columns.Item(2).Visible = False
        aGrid.Columns.Item(3).Visible = False
        aGrid.Columns.Item(4).TitleObject.Caption = "Requested Date"
        aGrid.Columns.Item(5).TitleObject.Caption = "Leave Type"
        aGrid.Columns.Item(5).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        otemp.DoQuery("Select U_Z_Name,U_Z_Name from [@Z_LeaveType] ")
        oCombobox = aGrid.Columns.Item(5)
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 0 To otemp.RecordCount - 1
            oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
            otemp.MoveNext()
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        aGrid.Columns.Item(6).TitleObject.Caption = "From Date"
        aGrid.Columns.Item(7).TitleObject.Caption = "To date"
        aGrid.Columns.Item(8).TitleObject.Caption = "Number of Days"
        aGrid.Columns.Item(8).Editable = False
        aGrid.Columns.Item(9).TitleObject.Caption = "Reason"
        aGrid.Columns.Item(10).TitleObject.Caption = "Approved"
        aGrid.Columns.Item(10).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oCombobox = aGrid.Columns.Item(10)
        oCombobox.ValidValues.Add("A", "Approved")
        oCombobox.ValidValues.Add("D", "Declined")
        oCombobox.ValidValues.Add("P", "Approval Pending")

        oCombobox.TitleObject.Caption = "Approval Status"
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        aGrid.Columns.Item(10).Editable = False
        aGrid.Columns.Item(11).TitleObject.Caption = "Remarks"
        aGrid.Columns.Item(11).Editable = False
        aGrid.Columns.Item(12).TitleObject.Caption = "Replacement"
        aGrid.Columns.Item(12).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        otemp.DoQuery("Select firstName + lastname,firstName + lastname from [OHEM] where empid<>" & aempid & "  order by empID ")
        oCombobox = aGrid.Columns.Item(12)
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 0 To otemp.RecordCount - 1
            oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
            otemp.MoveNext()
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        'oCombobox = aGrid.Columns.Item(12)
        'oCombobox.SetSelectedValue(aGrid.DataTable.Rows.Count - 1, oCombobox.ValidValues.Item("P"))
        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dtFrom, dtto As Date
        'strProject = oApplication.Utilities.getEdittextvalue(aform, "4")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oMatrix = aform.Items.Item("9").Specific
        For intRow As Integer = 0 To oMatrix.DataTable.Rows.Count - 1
            strCode = oMatrix.DataTable.GetValue("Code", intRow)
            If 1 = 1 Then
                oCombobox = oMatrix.Columns.Item("U_Z_TYPE")
                If oCombobox.GetSelectedValue(intRow).Value <> "" Then
                    Dim strFrom, strTo As String
                    '  MsgBox(oMatrix.DataTable.GetValue("U_Z_FROMDATE", intRow))
                    strFrom = oMatrix.DataTable.GetValue("U_Z_FROMDATE", intRow)
                    strTo = oMatrix.DataTable.GetValue("U_Z_TODATE", intRow)

                    If strFrom = "" Then
                        oMatrix.Columns.Item("U_Z_FROMDATE").Click(intRow, False, 1)
                        oApplication.Utilities.Message("From date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strTo = "" Then
                        oMatrix.Columns.Item("U_Z_TODATE").Click(intRow, False, 1)
                        oApplication.Utilities.Message("To date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    dtFrom = oMatrix.DataTable.GetValue("U_Z_FROMDATE", intRow)
                    dtto = oMatrix.DataTable.GetValue("U_Z_TODATE", intRow)
                    Dim intDiffer As Integer
                    intDiffer = DateDiff(DateInterval.Day, dtFrom, dtto)
                    If intDiffer < 0 Then
                        oApplication.Utilities.Message("To date should be greater than or equal to from date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("U_Z_TODATE").Click(intRow, False, 1)
                        Return False
                    End If

                    intDiffer = intDiffer + 1
                    Dim dtTemp As Date
                    Dim intWeekFrom, intWeekTo, intNumberofDays As Integer
                    Dim stTemp As SAPbobsCOM.Recordset
                    stTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    stTemp.DoQuery("SELECT T0.[WndFrm], T0.[WndTo] FROM OHLD T0  INNER JOIN OADM T1 ON T0.HldCode = T1.HldCode")
                    intNumberofDays = 0
                    If stTemp.RecordCount > 0 Then
                        intWeekFrom = stTemp.Fields.Item(0).Value
                        intWeekTo = stTemp.Fields.Item(1).Value
                        dtTemp = dtFrom
                        Dim strTemp As Date
                        strTemp = dtFrom.ToString
                        strTemp = oApplication.Utilities.getDateStrin(dtFrom)
                        dtTemp = oApplication.Utilities.GetDateTimeValue(strTemp)
                        Dim sttemp2 As SAPbobsCOM.Recordset
                        sttemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        For intDate As Integer = 1 To intDiffer
                            If dtTemp <= dtto Then
                                If dtTemp.DayOfWeek + 1 = intWeekFrom Or dtTemp.DayOfWeek + 1 = intWeekTo Then
                                    intNumberofDays = intNumberofDays + 1
                                Else
                                    sttemp2.DoQuery("SELECT T0.[HldCode], T0.[WndFrm], T0.[WndTo], T0.[isCurYear], T1.[StrDate], T1.[EndDate] FROM OHLD T0  INNER JOIN HLD1 T1 ON T0.HldCode = T1.HldCode INNER JOIN OADM T2 ON T0.HldCode = T2.HldCode where '" & dtTemp.ToString("yyyy-MM-dd") & "' between T1.[Strdate] and T1.[EndDate]")
                                    If sttemp2.RecordCount > 0 Then
                                        intNumberofDays = intNumberofDays + 1
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                            dtTemp = dtTemp.AddDays(1)
                        Next
                    End If

                   
                    intDiffer = intDiffer - intNumberofDays
                    oMatrix.DataTable.SetValue("U_Z_DAYS", intRow, intDiffer)
                    If oMatrix.DataTable.GetValue("U_Z_REASON", intRow) = "" Then
                        oMatrix.Columns.Item("U_Z_REASON").Click(intRow, False, 1)
                        oApplication.Utilities.Message("Leave request reason is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
        ' MsgBox(Year(dtFrom))
                End If
            End If
        Next
        Return True
    End Function
#End Region
#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID


                Case mnu_Leaverequest
                    If pVal.BeforeAction = False Then
                        ' LoadForm()
                        Dim oTe As New clsLogin
                        oTe.LoadForm("Leave Request")

                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("6").Enabled = False
                        oForm.Items.Item("6").Enabled = False
                    End If
                Case mnu_ADD_ROW

                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        ' MsgBox(RowtoDelete1)

                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = True Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("6").Enabled = True
                        ' oForm.Items.Item("8").Enabled = True
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = True Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("6").Enabled = True
                        'oForm.Items.Item("8").Enabled = True
                    End If
                Case mnu_Delete
                    If pVal.BeforeAction = True Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If pVal.BeforeAction = True Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm
                        If oForm.TypeEx = frm_Budget Then
                            If oApplication.SBO_Application.MessageBox("Do you want to remove this project details?", , "Continue", "Cancel") = 1 Then
                                Dim objEditText As SAPbouiCOM.EditText
                                Dim strDocNum As String
                                objEditText = oForm.Items.Item("4").Specific
                                strDocNum = objEditText.String
                            End If
                        Else
                            BubbleEvent = False
                            Exit Sub
                        End If
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

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strDocEntry As String
        aForm.Freeze(True)
        If validation(aForm) = False Then
            aForm.Freeze(False)
            Return False
        End If
        If AddToUDT_Table(aForm) = False Then
            aForm.Freeze(False)
            Return False
        End If
        aForm.Freeze(False)
        ' AssignLineNo(aForm)
        Return True
    End Function



#Region "AddtoUDT"

    
    Private Function AddToUDT_Table(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strDocref, strEmpID, strLineCode, stdocdate, strEmpName, strEmployeename, stremptype, strprojectname, strPrjCode, strAmount As String
        Dim dtDate As Date
        Dim intHours, dblAmount As Double
        Dim oTempRec, otemp As SAPbobsCOM.Recordset
        Dim ousertable As SAPbobsCOM.UserTable
        Dim ocheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oedittext As SAPbouiCOM.EditTextColumn
        Dim dblPercentage As Double
        Dim blnexits As Boolean = False
        Dim blnLines As Boolean = False
        Dim dtFrom, dtTo, dtRequestdate As Date
        Dim oBPGrid As SAPbouiCOM.Grid
        Dim strRef1 As String
        oBPGrid = aform.Items.Item("9").Specific

        ousertable = oApplication.Company.UserTables.Item("Z_OLEV")
        strEmpID = oApplication.Utilities.getEdittextvalue(aform, "6")
        strEmpName = oApplication.Utilities.getEdittextvalue(aform, "8")
        For intRow As Integer = 0 To oBPGrid.DataTable.Rows.Count - 1
            strCode = oBPGrid.DataTable.GetValue("Code", intRow)
            oCombobox = oBPGrid.Columns.Item("U_Z_TYPE")
            If strCode = "" And oCombobox.GetSelectedValue(intRow).Value <> "" Then
                strCode = oApplication.Utilities.getMaxCode("@Z_OLEV", "Code")
                dtRequestdate = oBPGrid.DataTable.GetValue(4, intRow)
                'dtRequestdate = oBPGrid.DataTable.GetValue("Z_TYPE", intRow)
                dtFrom = oBPGrid.DataTable.GetValue(6, intRow)
                dtTo = oBPGrid.DataTable.GetValue(7, intRow)
                ousertable.Code = strCode
                ousertable.Name = strCode
                ousertable.UserFields.Fields.Item("U_Z_EMPCODE").Value = strEmpID
                ousertable.UserFields.Fields.Item("U_Z_EMPNAME").Value = strEmpName
                ousertable.UserFields.Fields.Item("U_Z_DocDate").Value = dtRequestdate

                ousertable.UserFields.Fields.Item("U_Z_TYPE").Value = oCombobox.GetSelectedValue(intRow).Value
                ousertable.UserFields.Fields.Item("U_Z_FROMDATE").Value = dtFrom
                ousertable.UserFields.Fields.Item("U_Z_TODATE").Value = dtTo
                ousertable.UserFields.Fields.Item("U_Z_DAYS").Value = oBPGrid.DataTable.GetValue(8, intRow)
                ousertable.UserFields.Fields.Item("U_Z_REASON").Value = oBPGrid.DataTable.GetValue(9, intRow)
                oCombobox = oBPGrid.Columns.Item("U_Z_SUBEMP")
                Try
                    ousertable.UserFields.Fields.Item("U_Z_SUBEMP").Value = oCombobox.GetSelectedValue(intRow).Description
                Catch ex As Exception
                    ousertable.UserFields.Fields.Item("U_Z_SUBEMP").Value = ""
                End Try

                ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "P"
                If ousertable.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Else
                If ousertable.GetByKey(strCode) Then
                    ' strCode = oApplication.Utilities.getMaxCode("@Z_OLEV", "Code")
                    dtRequestdate = oBPGrid.DataTable.GetValue(4, intRow)
                    dtFrom = oBPGrid.DataTable.GetValue(6, intRow)
                    dtTo = oBPGrid.DataTable.GetValue(7, intRow)
                    ousertable.Code = strCode
                    ousertable.Name = strCode
                    ousertable.UserFields.Fields.Item("U_Z_EMPCODE").Value = strEmpID
                    ousertable.UserFields.Fields.Item("U_Z_EMPNAME").Value = strEmpName
                    ousertable.UserFields.Fields.Item("U_Z_DocDate").Value = dtRequestdate
                    ousertable.UserFields.Fields.Item("U_Z_TYPE").Value = oCombobox.GetSelectedValue(intRow).Value
                    ousertable.UserFields.Fields.Item("U_Z_FROMDATE").Value = dtFrom
                    ousertable.UserFields.Fields.Item("U_Z_TODATE").Value = dtTo
                    ousertable.UserFields.Fields.Item("U_Z_DAYS").Value = oBPGrid.DataTable.GetValue(8, intRow)
                    ousertable.UserFields.Fields.Item("U_Z_REASON").Value = oBPGrid.DataTable.GetValue(9, intRow)
                    oCombobox = oBPGrid.Columns.Item("U_Z_APPROVED")
                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = oCombobox.GetSelectedValue(intRow).Value
                    oCombobox = oBPGrid.Columns.Item("U_Z_SUBEMP")
                    Try
                        'MsgBox(oCombobox.GetSelectedValue(intRow).Description)
                        ousertable.UserFields.Fields.Item("U_Z_SUBEMP").Value = oCombobox.GetSelectedValue(intRow).Description
                    Catch ex As Exception
                        ousertable.UserFields.Fields.Item("U_Z_SUBEMP").Value = ""
                    End Try

                    If ousertable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
        Next
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec.DoQuery("Delete from [@Z_OLEV] where name like '%D'")
        oApplication.Utilities.Message("Operation completed successfuly", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Leaverequest Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "9" And pVal.CharPressed <> 9 Then
                                    oMatrix = oForm.Items.Item("9").Specific
                                    oCombobox = oMatrix.Columns.Item(10)
                                    Try
                                        If oCombobox.GetSelectedValue(pVal.Row).Value <> "P" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try
                                    


                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("9").Specific
                                If pVal.ItemUID = "9" Then
                                    oCombobox = oMatrix.Columns.Item(10)
                                    Try

                                    
                                        If oCombobox.GetSelectedValue(pVal.Row).Value <> "P" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "9" Then
                                    oCombobox = oMatrix.Columns.Item(10)
                                    Try
                                        If oCombobox.GetSelectedValue(pVal.Row).Value <> "P" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                        End Select


                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                If pVal.ItemUID = "9" And pVal.ColUID = "U_Z_TODATE" And pVal.CharPressed = 9 Then
                                    oMatrix = oForm.Items.Item("9").Specific
                                    Dim dtFrom, dtTo As Date
                                    Dim strFrom, strTo As String
                                    'dtFrom = oMatrix.DataTable.GetValue("U_Z_FROMDATE", pVal.Row)
                                    'dtTo = oMatrix.DataTable.GetValue("U_Z_TODATE", pVal.Row)

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then

                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False

                                        Exit Sub
                                    Else
                                        Dim obj As New clsLogin
                                        blnSourceForm = True
                                        strSourceformEmpID = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                        oForm.Close()
                                        obj.LoadForm("Leave Request")
                                    End If
                                ElseIf pVal.ItemUID = "4" Then
                                    oMatrix = oForm.Items.Item("9").Specific
                                    If oMatrix.DataTable.Rows.Count - 1 <= 0 Then
                                        oMatrix.DataTable.Rows.Add()
                                        oCombobox = oMatrix.Columns.Item("U_Z_APPROVED")
                                        oCombobox.SetSelectedValue(oMatrix.DataTable.Rows.Count - 1, oCombobox.ValidValues.Item("P"))
                                    End If
                                    If oMatrix.DataTable.GetValue(2, oMatrix.DataTable.Rows.Count - 1).ToString <> "" Then
                                        oMatrix.DataTable.Rows.Add()
                                        oCombobox = oMatrix.Columns.Item("U_Z_APPROVED")
                                        oCombobox.SetSelectedValue(oMatrix.DataTable.Rows.Count - 1, oCombobox.ValidValues.Item("P"))

                                    End If
                                    oMatrix.Columns.Item("U_Z_DOCDATE").Click(oMatrix.DataTable.Rows.Count - 1)
                                ElseIf pVal.ItemUID = "5" Then
                                    oMatrix = oForm.Items.Item("9").Specific
                                    For intRow As Integer = 0 To oMatrix.DataTable.Rows.Count - 1
                                        If oMatrix.Rows.IsSelected(intRow) Then
                                            oCombobox = oMatrix.Columns.Item("U_Z_APPROVED")
                                            If oCombobox.GetSelectedValue(intRow).Value = "P" Then
                                                Dim otemp As SAPbobsCOM.Recordset
                                                otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                otemp.DoQuery("Update [@Z_OLEV] set name=name +'D' where Code='" & oMatrix.DataTable.GetValue("Code", intRow) & "'")
                                                oMatrix.DataTable.Rows.Remove(intRow)
                                                Exit Sub
                                            End If

                                        End If
                                    Next
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
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
