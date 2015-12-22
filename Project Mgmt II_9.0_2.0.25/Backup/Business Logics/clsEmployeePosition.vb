Public Class clsEmployeePosition
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private objMatrix As SAPbouiCOM.Matrix
    Private objForm As SAPbouiCOM.Form
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombo As SAPbouiCOM.ComboBox
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
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_EmployeePostion) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_EmpoyeePostion, frm_EmployeePostion)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        BindData(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        oForm.Freeze(False)
    End Sub
    Private Sub BindData(ByVal aform As SAPbouiCOM.Form)
        Try
            oCombo = aform.Items.Item("5").Specific
            Dim otest As SAPbobsCOM.Recordset
            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery("SELECT * from CUFD where upper(TableID)='OHEM' and upper(AliasID)='Z_RATE'")
            If otest.RecordCount <= 0 Then
                otest.DoQuery("Update OADM set U_Z_EMPRATE='R'")
            End If
            otest.DoQuery("Select isnull(U_Z_EMPRATE,'R') from OADM")
            oGrid = aform.Items.Item("1").Specific
            '   oCombo.Select(otest.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Try
                If oCombo.Selected.Value = "P" Then
                    oGrid.DataTable.ExecuteQuery("SELECT T0.[empID], T0.[firstName] + ' ' + T0.[lastName], T0.[U_Z_RATE], T0.[U_Z_RATE] * 8 'U_HR_RATE' FROM OHEM T0 where T0.U_Z_ISPROJECT='Y' and T0.Active='Y' order by T0.[empID]")
                Else
                    oGrid.DataTable.ExecuteQuery("SELECT T0.[empID], T0.[firstName] + ' ' + T0.[lastName], T0.[U_DAILY_RATE], T0.[U_HR_RATE] FROM OHEM T0 where T0.U_Z_ISPROJECT='Y' and T0.Active='Y' order by T0.[empID]")
                End If
            Catch ex As Exception
                oGrid.DataTable.ExecuteQuery("SELECT T0.[empID], T0.[firstName] + ' ' + T0.[lastName], T0.[U_DAILY_RATE], T0.[U_HR_RATE] FROM OHEM T0 where T0.U_Z_ISPROJECT='Y' and T0.Active='Y' order by T0.[empID]")
            End Try

            'oGrid.DataTable.ExecuteQuery("SELECT T0.[posID], T0.[name], T0.[U_DAILY_RATE], T0.[U_HR_RATE] FROM OHPS T0 order by T0.[posID]")
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee ID"
            oGrid.Columns.Item(0).Editable = False
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).TitleObject.Caption = "Daily Rate"
            oGrid.Columns.Item(2).Editable = True
            oGrid.Columns.Item(3).TitleObject.Caption = "Hourly Rate"
            oGrid.Columns.Item(3).Editable = True
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PopulateDate(ByVal aform As SAPbouiCOM.Form)
        Try
            oCombo = aform.Items.Item("5").Specific
            Dim otest As SAPbobsCOM.Recordset
            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery("SELECT * from CUFD where upper(TableID)='OHEM' and upper(AliasID)='Z_RATE'")
            If otest.RecordCount <= 0 Then
                otest.DoQuery("Update OADM set U_Z_EMPRATE='R'")
            End If
            If oCombo.Selected.Value = "P" Then
                oGrid.DataTable.ExecuteQuery("SELECT T0.[empID], T0.[firstName] + ' ' + T0.[lastName], T0.[U_RATE], T0.[U_RATE] * 8 'U_HR_RATE' FROM OHEM T0 where T0.U_Z_ISPROJECT='Y' and T0.Activt='Y' order by T0.[empID]")
            Else
                oGrid.DataTable.ExecuteQuery("SELECT T0.[empID], T0.[firstName] + ' ' + T0.[lastName], T0.[U_DAILY_RATE], T0.[U_HR_RATE] FROM OHEM T0 where T0.U_Z_ISPROJECT='Y' and T0.Activt='Y' order by T0.[empID]")
            End If
            oGrid = aform.Items.Item("1").Specific
            'oGrid.DataTable.ExecuteQuery("SELECT T0.[posID], T0.[name], T0.[U_DAILY_RATE], T0.[U_HR_RATE] FROM OHPS T0 order by T0.[posID]")
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee ID"
            oGrid.Columns.Item(0).Editable = False
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).TitleObject.Caption = "Daily Rate"
            oGrid.Columns.Item(2).Editable = True
            oGrid.Columns.Item(3).TitleObject.Caption = "Hourly Rate"
            oGrid.Columns.Item(3).Editable = True
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None

        Catch ex As Exception

        End Try
    End Sub


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_EmpPosition
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND, mnu_ADD_ROW, mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_Activity Then
                        If pVal.BeforeAction = False Then

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
            If pVal.FormTypeEx = frm_EmployeePostion Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                              
                        End Select
                    Case False
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            If pVal.ItemUID = "5" Then
                                BindData(oForm)
                            End If
                        End If
                        If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.ItemUID = "3")) Then
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Dim oTemp As SAPbobsCOM.Recordset
                            Dim strsql As String
                            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'T0.[posID], T0.[name], T0.[U_DAILY_RATE], T0.[U_HR_RATE]
                            oGrid = oForm.Items.Item("1").Specific
                            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                strsql = "Update OHEM set U_Daily_Rate=" & oGrid.DataTable.GetValue(2, intRow) & ",U_HR_Rate=" & oGrid.DataTable.GetValue(3, intRow)
                                strsql = strsql & " where empID=" & oGrid.DataTable.GetValue(0, intRow)
                                oTemp.DoQuery(strsql)
                            Next
                            Try
                                oCombo = oForm.Items.Item("5").Specific
                                oTemp.DoQuery("Update OADM set U_Z_EMPRATE='" & oCombo.Selected.Value & "'")
                            Catch ex As Exception

                            End Try

                            BindData(oForm)



                        End If

                End Select
            End If

        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

End Class
