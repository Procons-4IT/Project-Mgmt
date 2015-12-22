Public Class clsApproval
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
    Private strChocie As String = ""
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm(ByVal aQuery As String, ByVal aForm As SAPbouiCOM.Form)
        Dim strViewtype As String
        If aForm.TypeEx = frm_ProjectTime Then
            oCombobox = aForm.Items.Item("4").Specific
            strChocie = oCombobox.Selected.Value
            oCombobox = aForm.Items.Item("10").Specific
            strViewtype = oCombobox.Selected.Value
        End If
        oApplication.Utilities.LoadForm(xml_Details, frm_Approal)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        If oForm.TypeEx = frm_Approal Then
            oForm.Freeze(True)
            If strChocie = "E" Then
                oForm.Title = "Expenses Approval"
            ElseIf strChocie = "T" Then
                oForm.Title = "Time Sheet Approval"
            ElseIf (strChocie = "L") Then
                oForm.Title = "Leave Request Approval"
            End If
            oGrid = oForm.Items.Item("1").Specific
            oGrid.DataTable.ExecuteQuery(aQuery)
            FormatGrid(oGrid, strChocie)
            oForm.Freeze(False)
        End If
    End Sub
#Region "FormatGrid"
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid, ByVal aChoice As String)
        Dim oColumn As SAPbouiCOM.Column
        Select Case aChoice
            Case "E"
                aGrid.Columns.Item("U_Z_EMPCODE").TitleObject.Caption = "Employee Code"
                oEditTextColumn = oGrid.Columns.Item("U_Z_EMPCODE")
                oEditTextColumn.LinkedObjectType = "171"
                aGrid.Columns.Item("U_Z_EMPCODE").Editable = False
                aGrid.Columns.Item("U_Z_EMPNAME").TitleObject.Caption = "Employee Name"
                aGrid.Columns.Item("U_Z_EMPNAME").Editable = False
                aGrid.Columns.Item("U_Z_DOCDATE").TitleObject.Caption = "Document Date"
                aGrid.Columns.Item("U_Z_DOCDATE").Editable = False
                aGrid.Columns.Item("U_Z_EXPNAME").TitleObject.Caption = "Expenses Name"
                aGrid.Columns.Item("U_Z_EXPNAME").Editable = False
                aGrid.Columns.Item("U_Z_EXPTYPE").TitleObject.Caption = "Expenses Type"
                aGrid.Columns.Item("U_Z_EXPTYPE").Editable = False
                aGrid.Columns.Item("U_Z_EXPTYPE").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                ocombo = aGrid.Columns.Item("U_Z_EXPTYPE")
                ocombo.ValidValues.Add("N", "Normal")
                ocombo.ValidValues.Add("P", "Project")
                ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                aGrid.Columns.Item("U_Z_PRJCODE").TitleObject.Caption = "Project Code"
                oEditTextColumn = oGrid.Columns.Item("U_Z_PRJCODE")
                oEditTextColumn.LinkedObjectType = "63"
                aGrid.Columns.Item("U_Z_PRJCODE").Editable = False

                aGrid.Columns.Item("U_Z_PRJNAME").TitleObject.Caption = "Project name"
                aGrid.Columns.Item("U_Z_PRJNAME").Editable = False
                oEditTextColumn = aGrid.Columns.Item("U_Z_AMOUNT")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item("U_Z_AMOUNT").TitleObject.Caption = "Txn Currency Amount"
                aGrid.Columns.Item("U_Z_AMOUNT").Editable = True
                aGrid.Columns.Item("U_Z_REFCODE").TitleObject.Caption = "Document No"
                aGrid.Columns.Item("U_Z_REFCODE").Editable = False
                aGrid.Columns.Item("U_Z_REFCODE").Visible = True

                aGrid.Columns.Item("CODE").TitleObject.Caption = "Code"
                aGrid.Columns.Item("CODE").Editable = False
                aGrid.Columns.Item("CODE").Visible = False
                aGrid.Columns.Item("U_Z_APPROVED").TitleObject.Caption = "Approval Status"
                aGrid.Columns.Item("U_Z_APPROVED").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                ocombo = aGrid.Columns.Item("U_Z_APPROVED")
                ocombo.ValidValues.Add("P", "Pending")
                ocombo.ValidValues.Add("A", "Approved")
                ocombo.ValidValues.Add("D", "Declined")
                ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

                aGrid.Columns.Item("U_Z_APPROVED").Editable = True
                aGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                aGrid.Columns.Item("U_Z_CURRENCY").TitleObject.Caption = "Currency"
                aGrid.Columns.Item("U_Z_CURRENCY").Editable = False
                'ocombo = aGrid.Columns.Item("U_Z_APPROVED")
                'ocombo.ValidValues.Add("A", "Approved")
                'ocombo.ValidValues.Add("N", "Non-Approved")
                'ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            Case "T"
                aGrid.Columns.Item("U_Z_EMPCODE").TitleObject.Caption = "Employee Code"
                aGrid.Columns.Item("U_Z_EMPCODE").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_EMPCODE")
                oEditTextColumn.LinkedObjectType = "171"
                aGrid.Columns.Item("U_Z_EMPNAME").TitleObject.Caption = "Employee Name"
                aGrid.Columns.Item("U_Z_EMPNAME").Editable = False
                aGrid.Columns.Item("U_Z_DOCDATE").TitleObject.Caption = "Document Date"
                aGrid.Columns.Item("U_Z_DOCDATE").Editable = False

                aGrid.Columns.Item("U_Z_PRJCODE").TitleObject.Caption = "Project Code"
                aGrid.Columns.Item("U_Z_PRJCODE").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_PRJCODE")
                oEditTextColumn.LinkedObjectType = "63"
                aGrid.Columns.Item("U_Z_PRJNAME").TitleObject.Caption = "Project name"
                aGrid.Columns.Item("U_Z_PRJNAME").Editable = False
                aGrid.Columns.Item("U_Z_PRCNAME").TitleObject.Caption = "Phase "
                aGrid.Columns.Item("U_Z_PRCNAME").Editable = False
                aGrid.Columns.Item("U_Z_ACTNAME").TitleObject.Caption = "Activity"
                aGrid.Columns.Item("U_Z_ACTNAME").Editable = False
                aGrid.Columns.Item("U_Z_DATE").TitleObject.Caption = "Date"
                aGrid.Columns.Item("U_Z_DATE").Editable = False
                aGrid.Columns.Item("U_Z_DATE").Visible = False
                oEditTextColumn = aGrid.Columns.Item("U_Z_HOURS")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item("U_Z_HOURS").TitleObject.Caption = "No of Hours"
                aGrid.Columns.Item("U_Z_HOURS").Editable = True


                oEditTextColumn = aGrid.Columns.Item("U_Z_QUANTITY")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item("U_Z_QUANTITY").TitleObject.Caption = "Quantity"
                aGrid.Columns.Item("U_Z_QUANTITY").Visible = False

                oEditTextColumn = aGrid.Columns.Item("U_Z_BDGQTY")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item("U_Z_BDGQTY").TitleObject.Caption = "Budgeted.Hours"
                aGrid.Columns.Item("U_Z_BDGQTY").Editable = False

                aGrid.Columns.Item("U_Z_MEASURE").TitleObject.Caption = "Measure"
                aGrid.Columns.Item("U_Z_MEASURE").Visible = False

                aGrid.Columns.Item("U_Z_TYPE").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                ocombo = aGrid.Columns.Item("U_Z_TYPE")
                ocombo.ValidValues.Add("I", "Item")
                ocombo.ValidValues.Add("R", "Resource")
                ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                aGrid.Columns.Item("U_Z_TYPE").TitleObject.Caption = "Activity Type"
                aGrid.Columns.Item("U_Z_TYPE").Editable = False


                aGrid.Columns.Item("U_Z_REFCODE").TitleObject.Caption = "Document No"
                aGrid.Columns.Item("U_Z_REFCODE").Editable = False
                aGrid.Columns.Item("U_Z_REFCODE").Visible = True
                aGrid.Columns.Item("CODE").TitleObject.Caption = "Code"
                aGrid.Columns.Item("CODE").Editable = False
                aGrid.Columns.Item("CODE").Visible = False

                aGrid.Columns.Item("U_Z_APPROVED").TitleObject.Caption = "Approval Status"
                aGrid.Columns.Item("U_Z_APPROVED").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                ocombo = aGrid.Columns.Item("U_Z_APPROVED")
                ocombo.ValidValues.Add("P", "Pending")
                ocombo.ValidValues.Add("A", "Approved")
                ocombo.ValidValues.Add("D", "Declined")
                ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                aGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                'ocombo = aGrid.Columns.Item("U_Z_APPROVED")
                'ocombo.ValidValues.Add("A", "Approved")
                'ocombo.ValidValues.Add("N", "Non-Approved")
                'ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            Case "L"
                aGrid.Columns.Item(0).Visible = False
                aGrid.Columns.Item(1).Visible = False
                aGrid.Columns.Item(2).Visible = True
                aGrid.Columns.Item(2).TitleObject.Caption = "Employee Code"
                aGrid.Columns.Item(2).Editable = False
                oEditTextColumn = oGrid.Columns.Item(2)
                oEditTextColumn.LinkedObjectType = "171"
                aGrid.Columns.Item(3).Visible = True
                aGrid.Columns.Item(3).TitleObject.Caption = "Employee Name"
                aGrid.Columns.Item(3).Editable = False
                aGrid.Columns.Item(4).TitleObject.Caption = "Requested Date"
                aGrid.Columns.Item(4).Editable = False
                aGrid.Columns.Item(5).TitleObject.Caption = "Leave Type"
                aGrid.Columns.Item(5).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                Dim otemp As SAPbobsCOM.Recordset
                otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemp.DoQuery("Select U_Z_Name,U_Z_Name from [@Z_LeaveType] ")
                ocombo = aGrid.Columns.Item(5)
                ocombo.ValidValues.Add("", "")
                For intRow As Integer = 0 To otemp.RecordCount - 1
                    ocombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                    otemp.MoveNext()
                Next
                ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                aGrid.Columns.Item(5).Editable = False
                aGrid.Columns.Item(6).TitleObject.Caption = "From Date"
                aGrid.Columns.Item(6).Editable = False
                aGrid.Columns.Item(7).TitleObject.Caption = "To date"
                aGrid.Columns.Item(7).Editable = False
                aGrid.Columns.Item(8).TitleObject.Caption = "Number of Days"
                aGrid.Columns.Item(8).Editable = False
                aGrid.Columns.Item(8).Editable = False
                aGrid.Columns.Item(9).TitleObject.Caption = "Reason"
                aGrid.Columns.Item(9).Editable = False
                aGrid.Columns.Item(10).TitleObject.Caption = "Approved"
                aGrid.Columns.Item(10).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                ocombo = aGrid.Columns.Item(10)
                ocombo.ValidValues.Add("P", "Pending")
                ocombo.ValidValues.Add("A", "Approved")
                ocombo.ValidValues.Add("D", "Declined")
                ocombo.TitleObject.Caption = "Approval Status"
                ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                aGrid.Columns.Item(10).Editable = True
                aGrid.Columns.Item(11).TitleObject.Caption = "Remarks"
                aGrid.Columns.Item(11).Editable = True
                aGrid.Columns.Item(12).TitleObject.Caption = "SubLevel Employee"
                'aGrid.Columns.Item(12).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                'aGrid.Columns.Item(12).Editable = True
                'otemp.DoQuery("Select empid,firstName + lastname from [OHEM] order by empID ")
                'oCombobox = aGrid.Columns.Item(12)
                'oCombobox.ValidValues.Add("", "")
                'For intRow As Integer = 0 To otemp.RecordCount - 1
                '    oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                '    otemp.MoveNext()
                'Next
                'oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                aGrid.Columns.Item(12).Editable = False
                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        End Select
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        aGrid.AutoResizeColumns()
    End Sub
#End Region

#Region "Add to UDT"
    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim ousertable As SAPbobsCOM.UserTable
        Dim strCode, strField, strQuery, strtable, dblValue As String
        Dim oCheckboxcolumn As SAPbouiCOM.CheckBoxColumn
        Dim otemprs As SAPbobsCOM.Recordset
        aForm.Freeze(True)
        If aForm.Title = "Expenses Approval" Then
            strtable = "[@Z_EXP1]"
            strQuery = "U_Z_AMOUNT"
            strField = "U_Z_Amount"
        ElseIf aForm.Title = "Time Sheet Approval" Then
            strtable = "[@Z_TIM1]"
            strQuery = "U_Z_HOURS"
            strField = "U_Z_Hours"
        ElseIf aForm.Title = "Leave Request Approval" Then
            strtable = "[@Z_OLEV]"
            strQuery = "U_Z_DAYS"
            strField = "U_Z_DAYS"
        End If
        otemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aForm.Items.Item("1").Specific
        Dim strChoice, strremarks As String
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            ocombo = oGrid.Columns.Item("U_Z_APPROVED")
            Try
                strCode = oGrid.DataTable.GetValue("CODE", intRow)
            Catch ex As Exception

                strCode = oGrid.DataTable.GetValue("Code", intRow)
            End Try

            If strCode <> "" Then
                strChoice = ocombo.GetSelectedValue(intRow).Value
                dblValue = oGrid.DataTable.GetValue(strQuery, intRow)
                Dim strValue As String
                If CompanyDecimalSeprator <> "." Then
                    strValue = dblValue.ToString
                    strValue = strValue.Replace(CompanyDecimalSeprator, ".")
                Else
                    strValue = dblValue.ToString
                End If
                Try
                    strremarks = oGrid.DataTable.GetValue("U_Z_Remarks", intRow)
                Catch ex As Exception
                    strremarks = oGrid.DataTable.GetValue("U_Z_REMARKS", intRow)
                End Try
                '  strSQL = "Update " & strtable & " set U_Z_Remarks='" & strremarks & "', U_Z_Approved='" & strChoice & "'," & strField & "='" & dblValue & "' where Code='" & strCode & "'"
                strSQL = "Update " & strtable & " set U_Z_Remarks='" & strremarks & "', U_Z_Approved='" & strChoice & "'," & strField & "=" & strValue & " where Code='" & strCode & "'"
                otemprs.DoQuery(strSQL)
            End If
        Next
        aForm.Freeze(False)
        Return True
    End Function
#End Region

#Region "Load Details"
    Private Sub loaddetails(ByVal aform As SAPbouiCOM.Form, ByVal intRow As Integer)
        oGrid = aform.Items.Item("1").Specific
        If aform.Title = "Expenses Approval" Then
            Dim obj As New clsExpaneEntry
            If intRow >= 0 Then
                obj.DataView(oGrid.DataTable.GetValue("U_Z_EMPCODE", intRow), oGrid.DataTable.GetValue("U_Z_EMPNAME", intRow), oGrid.DataTable.GetValue("U_Z_DOCDATE", intRow))
            End If
        Else
            Dim obj As New clsEmpTimeSheet

            If intRow >= 0 Then
                obj.DataView(oGrid.DataTable.GetValue("U_Z_EMPCODE", intRow), oGrid.DataTable.GetValue("U_Z_EMPNAME", intRow), oGrid.DataTable.GetValue("U_Z_DOCDATE", intRow))
            End If

        End If
    End Sub
#End Region

#Region "Validations"
    Private Sub Selectall(ByVal aForm As SAPbouiCOM.Form, ByVal blnValue As Boolean)
        Dim ocheckboxcolumn As SAPbouiCOM.CheckBoxColumn
        Dim ovalue As SAPbouiCOM.ValidValue
        oGrid = aForm.Items.Item("1").Specific
        aForm.Freeze(True)
        For introw As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            ocombo = oGrid.Columns.Item("U_Z_APPROVED")
            ocombo.SetSelectedValue(introw, ocombo.ValidValues.Item(1))
        Next
        aForm.Freeze(False)
    End Sub
#End Region

#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try

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
            If pVal.FormTypeEx = frm_Approal Then
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
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then
                                    loaddetails(oForm, pVal.Row)
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        If AddtoUDT(oForm) = True Then
                                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            If oForm.Title = "Expenses Approval" Then
                                                Dim objct As New clsPrjTime
                                                strApprovalType = "Exp"
                                                objct.LoadForm()
                                            ElseIf oForm.Title.Contains("Time") Then
                                                Dim objct As New clsPrjTime
                                                strApprovalType = "Time"
                                                objct.LoadForm()
                                            ElseIf oForm.Title.Contains("Leave") Then
                                                Dim objct As New clsPrjTime
                                                strApprovalType = "Leave"
                                                objct.LoadForm()
                                            End If
                                            oForm.Close()

                                            
                                        End If
                                    Case "4"
                                        Selectall(oForm, True)
                                    Case "5"
                                        Selectall(oForm, False)
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)


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
