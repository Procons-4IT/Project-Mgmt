Public Class clsEmpTimeSheet
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private oColumn As SAPbouiCOM.Column
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Matrix
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath, strwhs, strGirdValue As String
    Private InvForConsumedItems As Integer
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
        oForm = oApplication.Utilities.LoadForm(xml_ExpEntry, frm_ExpEntry)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.PaneLevel = 1
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
            oForm = oApplication.Utilities.LoadForm(xml_EmpTime, frm_Timesheet)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            If oForm.TypeEx = frm_Timesheet Then
                oForm.Freeze(True)
                AddChooseFromList(oForm)
                oForm.PaneLevel = 1
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oApplication.Utilities.setEdittextvalue(oForm, "4", aEmpid)
                oApplication.Utilities.setEdittextvalue(oForm, "6", strName)
                oForm.Items.Item("4").Enabled = False
                ' oForm.Items.Item("6").Enabled = False
                If aOption = "A" Then
                    databind(oForm, aDate)
                Else
                    dtdate = aDate
                    ' databind_View(oForm, dtdate)
                    databind_yViewEntry(oForm, dtdate)
                End If
            End If
            oForm.Freeze(False)
        End If
    End Sub

    Public Sub DataView(ByVal aEmpID As String, ByVal aName As String, ByVal dtDate As Date)
        Try
            Dim oCode As String
            Dim strSQL As String
            Dim otemp As SAPbobsCOM.Recordset
            oForm = oApplication.Utilities.LoadForm(xml_EmpTime, frm_Timesheet)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            AddChooseFromList(oForm)
            oForm.PaneLevel = 1
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

            oApplication.Utilities.setEdittextvalue(oForm, "4", aEmpID)
            oApplication.Utilities.setEdittextvalue(oForm, "6", aName)
            databind_View(oForm, dtDate)
            oForm.Items.Item("4").Enabled = False
            oForm.Freeze(False)

            ' oForm.Items.Item("6").Enabled = False
        Catch ex As Exception

        End Try
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
            oCFLCreationParams.ObjectType = "Z_PRJ"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_STATUS"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "X"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region


#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form, ByVal aDate As Date)
        Try
            Dim oCode As String
            Dim strSQL As String
            Dim dtdate1 As Date
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            aForm.Freeze(True)
            '   strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & Now.Date.ToString("dd-MM-yyyy") & "'"
            strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & aDate.ToString("dd-MM-yyyy") & "'"
            'strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and Convert(varchar(10),U_Z_DocDate,105)='" & aDate & "'"
            otemp.DoQuery(strSQL)
            If otemp.RecordCount > 0 Then
                oCode = otemp.Fields.Item("Code").Value
                dtdate1 = otemp.Fields.Item("U_Z_DocDate").Value
                dtdate1 = aDate
            Else

                oCode = oApplication.Utilities.getMaxCode("@Z_OTIM", "Code")
                dtdate1 = aDate
            End If
            oApplication.Utilities.setEdittextvalue(aForm, "11", oCode)
            oApplication.Utilities.setEdittextvalue(aForm, "13", aDate)

            oMatrix = aForm.Items.Item("12").Specific
            oEditText = aForm.Items.Item("4").Specific
            oEditText.ChooseFromListUID = "CFL2"
            oEditText.ChooseFromListAlias = "empID"
            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "U_Z_PRJCODE"
            oColumn = oMatrix.Columns.Item("V_1")

            oColumn = oMatrix.Columns.Item("V_8")
            Try
                oColumn.ValidValues.Add("A", "Approved")
                oColumn.ValidValues.Add("D", "Declined")
                oColumn.ValidValues.Add("P", "Approval Pending")

            Catch ex As Exception
            End Try
            oColumn.TitleObject.Caption = "Approval Status"
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_9")
            Try
                oColumn.ValidValues.Add("P", "Pending")
                oColumn.ValidValues.Add("C", "Confirmed")
            Catch ex As Exception
            End Try
            oColumn = oMatrix.Columns.Item("V_9")
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_5")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("Qty")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            ' oColumn.Visible = False

            oColumn = oMatrix.Columns.Item("BdgQty")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("BdgQty")
            oColumn.Visible = False
            oMatrix.AddRow()
            oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Try
                oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                Dim strdate As String
                strdate = oApplication.Utilities.getDateStrin(dtdate1)
                If strdate <> "" Then
                    oEditText.String = strdate
                End If

            Catch ex As Exception
            End Try
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single


            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "Select * from [@Z_TIM1] where U_Z_RefCode='" & oCode & "'"
            otemp.DoQuery(strSQL)
            For introw As Integer = 0 To otemp.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRJCODE").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRCNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRJNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, otemp.Fields.Item("U_Z_ACTNAME").Value)

                oApplication.Utilities.SetMatrixValues(oMatrix, "Qty", oMatrix.RowCount, otemp.Fields.Item("U_Z_QUANTITY").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "Measure", oMatrix.RowCount, otemp.Fields.Item("U_Z_MEASURE").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "BdgQty", oMatrix.RowCount, otemp.Fields.Item("U_Z_BDGQTY").Value)

                Try
                    oCombobox = oMatrix.Columns.Item("Type").Cells.Item(oMatrix.RowCount).Specific
                    oCombobox.Select(otemp.Fields.Item("U_Z_TYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try
                
                Try
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Dim strdate As String
                    strdate = oApplication.Utilities.getDateStrin(dtdate1)
                    If strdate <> "" Then
                        oEditText.String = strdate
                    End If

                Catch ex As Exception
                End Try

                'oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_DATE").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, otemp.Fields.Item("U_Z_HOURS").Value)
                oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_Approved").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oCombobox.Select(otemp.Fields.Item("U_Z_EmpApproval").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                Catch ex As Exception
                    oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End Try


                oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, otemp.Fields.Item("U_Z_REMARKS").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, otemp.Fields.Item("Code").Value)
                oMatrix.AddRow()
                otemp.MoveNext()
            Next
            oColumn = oMatrix.Columns.Item("BdgQty")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Qty")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Measure")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Type")
            oColumn.Visible = False
            oMatrix.Columns.Item("V_4").Editable = True
            oMatrix.AutoResizeColumns()
            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub databind_yViewEntry(ByVal aForm As SAPbouiCOM.Form, ByVal aDate As Date)
        Try
            Dim oCode As String
            Dim strSQL, strEmpid As String
            Dim dtdate1 As Date
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            aForm.Freeze(True)
            strEmpid = oApplication.Utilities.getEdittextvalue(aForm, "4")
            '   strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & Now.Date.ToString("dd-MM-yyyy") & "'"
            strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & aDate.ToString("dd-MM-yyyy") & "'"
            'strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and Convert(varchar(10),U_Z_DocDate,105)='" & aDate & "'"
            otemp.DoQuery(strSQL)
            If otemp.RecordCount > 0 Then
                oCode = otemp.Fields.Item("Code").Value
                dtdate1 = otemp.Fields.Item("U_Z_DocDate").Value
                dtdate1 = aDate
            Else
                oCode = oApplication.Utilities.getMaxCode("@Z_OTIM", "Code")
                dtdate1 = aDate
            End If
            oApplication.Utilities.setEdittextvalue(aForm, "11", oCode)
            oApplication.Utilities.setEdittextvalue(aForm, "13", aDate)
            oMatrix = aForm.Items.Item("12").Specific
            oEditText = aForm.Items.Item("4").Specific
            oEditText.ChooseFromListUID = "CFL2"
            oEditText.ChooseFromListAlias = "empID"
            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "U_Z_PRJCODE"
            oColumn = oMatrix.Columns.Item("V_1")
            oColumn = oMatrix.Columns.Item("V_8")
            Try
                oColumn.ValidValues.Add("A", "Approved")
                oColumn.ValidValues.Add("D", "Declined")
                oColumn.ValidValues.Add("P", "Approval Pending")

            Catch ex As Exception
            End Try
            oColumn.TitleObject.Caption = "Approval Status"
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_9")
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_5")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oColumn = oMatrix.Columns.Item("Qty")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("BdgQty")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oMatrix.AddRow()
            oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)

            Try
                oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                Dim strdate As String
                strdate = oApplication.Utilities.getDateStrin(dtdate1)
                If strdate <> "" Then
                    oEditText.String = strdate
                End If

            Catch ex As Exception
            End Try
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single


            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'strSQL = "Select * from [@Z_TIM1] where U_Z_RefCode='" & oCode & "'"
            strSQL = "Select * from [@Z_TIM1] where isnull(U_Z_APPROVED,'P')='P' and U_Z_RefCode in (Select Code from [@Z_OTIM] where  U_Z_EMPCODE='" & strEmpid & " ') order by U_Z_DATE,Code"
            otemp.DoQuery(strSQL)

            For introw As Integer = 0 To otemp.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRJCODE").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRCNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRJNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, otemp.Fields.Item("U_Z_ACTNAME").Value)


                oApplication.Utilities.SetMatrixValues(oMatrix, "Qty", oMatrix.RowCount, otemp.Fields.Item("U_Z_QUANTITY").Value)
                'oApplication.Utilities.SetMatrixValues(oMatrix, "Measure", oMatrix.RowCount, otemp.Fields.Item("U_Z_MEASURE").Value) 
                oApplication.Utilities.SetMatrixValues(oMatrix, "BdgQty", oMatrix.RowCount, otemp.Fields.Item("U_Z_HOURS").Value)
                '   oApplication.Utilities.SetMatrixValues(oMatrix, "BdgQty", oMatrix.RowCount, otemp.Fields.Item("U_Z_BDGQTY").Value)
                Try
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Dim strdate As String
                    dtdate1 = otemp.Fields.Item("U_Z_Date").Value
                    strdate = oApplication.Utilities.getDateStrin(dtdate1)
                    If strdate <> "" Then
                        oEditText.String = strdate
                    End If

                Catch ex As Exception
                End Try


                Try
                    oCombobox = oMatrix.Columns.Item("Type").Cells.Item(oMatrix.RowCount).Specific
                    oCombobox.Select(otemp.Fields.Item("U_Z_TYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try


                oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, otemp.Fields.Item("U_Z_HOURS").Value)
                oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_Approved").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oCombobox.Select(otemp.Fields.Item("U_Z_EmpApproval").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                Catch ex As Exception
                    oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)

                End Try

                oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, otemp.Fields.Item("U_Z_REMARKS").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, otemp.Fields.Item("code").Value)

                oMatrix.AddRow()
                otemp.MoveNext()
            Next
            oColumn = oMatrix.Columns.Item("BdgQty")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Qty")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Measure")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Type")
            oColumn.Visible = False
            oMatrix.Columns.Item("V_4").Editable = True
            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
            oMatrix.AutoResizeColumns()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub databind_View(ByVal aForm As SAPbouiCOM.Form, ByVal dtDate As Date)
        Try
            Dim oCode As String
            Dim strSQL As String
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            aForm.Freeze(True)
            strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & dtDate.ToString("dd-MM-yyyy") & "'"
            otemp.DoQuery(strSQL)
            If otemp.RecordCount > 0 Then
                oCode = otemp.Fields.Item("Code").Value
            Else
                oCode = oApplication.Utilities.getMaxCode("@Z_OTIM", "Code")
            End If
            oApplication.Utilities.setEdittextvalue(aForm, "11", oCode)
            oApplication.Utilities.setEdittextvalue(aForm, "13", dtDate)
            oMatrix = aForm.Items.Item("12").Specific
            oEditText = aForm.Items.Item("4").Specific
            oEditText.ChooseFromListUID = "CFL2"
            oEditText.ChooseFromListAlias = "empID"
            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "U_Z_PRJCODE"
            oColumn = oMatrix.Columns.Item("V_1")

            oColumn = oMatrix.Columns.Item("V_8")
            Try
                oColumn.ValidValues.Add("A", "Approved")
                oColumn.ValidValues.Add("D", "Declined")
                oColumn.ValidValues.Add("P", "Approval Pending")

            Catch ex As Exception
            End Try
            oColumn.TitleObject.Caption = "Approval Status"
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_9")
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_5")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("Qty")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("BdgQty")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oMatrix.AddRow()

            oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)

            'oCombobox = oMatrix.Columns.Item("V_52").Cells.Item(1).Specific
            'oCombobox.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)

            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strSQL = "Select * from [@Z_TIM1] where U_Z_RefCode='" & oCode & "'"
            otemp.DoQuery(strSQL)
            For introw As Integer = 0 To otemp.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRJCODE").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRCNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, otemp.Fields.Item("U_Z_PRJNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, otemp.Fields.Item("U_Z_ACTNAME").Value)

                oApplication.Utilities.SetMatrixValues(oMatrix, "Qty", oMatrix.RowCount, otemp.Fields.Item("U_Z_QUANTITY").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "Measure", oMatrix.RowCount, otemp.Fields.Item("U_Z_MEASURE").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "BdgQty", oMatrix.RowCount, otemp.Fields.Item("U_Z_BDGQTY").Value)
                Try
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Dim strdate As String
                    strdate = oApplication.Utilities.getDateStrin(dtDate)
                    If strdate <> "" Then
                        oEditText.String = strdate
                    End If
                Catch ex As Exception
                End Try

                Try
                    oCombobox = oMatrix.Columns.Item("Type").Cells.Item(oMatrix.RowCount).Specific
                    oCombobox.Select(otemp.Fields.Item("U_Z_TYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try

                oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, otemp.Fields.Item("U_Z_HOURS").Value)
                oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_Approved").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oCombobox.Select(otemp.Fields.Item("U_Z_EmpApproval").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                Catch ex As Exception
                    oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End Try
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, otemp.Fields.Item("U_Z_REMARKS").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, otemp.Fields.Item("code").Value)

                oMatrix.AddRow()
                otemp.MoveNext()
            Next
            oColumn = oMatrix.Columns.Item("BdgQty")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Qty")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Measure")
            oColumn.Visible = False
            oColumn = oMatrix.Columns.Item("Type")
            oColumn.Visible = False
            aForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Items.Item("12").Enabled = False
            aForm.Items.Item("21").Enabled = False
            aForm.Items.Item("22").Enabled = False
            aForm.Items.Item("1").Enabled = False
            oMatrix.AutoResizeColumns()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("12").Specific
            For Count As Integer = 1 To oMatrix.VisualRowCount
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", Count, Count)
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub


#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("12").Specific
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
                oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Try
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Dim strdate As String
                    Dim dtdate1 As Date
                    strdate = oApplication.Utilities.getEdittextvalue(aForm, "13")
                    dtdate1 = oApplication.Utilities.GetDateTimeValue(strdate)
                    If strdate <> "" Then
                        oEditText.String = dtdate1
                    End If
                Catch ex As Exception
                End Try
            End If
            oEditText = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
            Try
                If oEditText.Value <> "" Then
                    oMatrix.AddRow()
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, "")
                    oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                    oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    ' oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, "")
                    Try
                        oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                        Dim strdate As String
                        Dim dtdate1 As Date
                        strdate = oApplication.Utilities.getEdittextvalue(aForm, "13")
                        dtdate1 = oApplication.Utilities.GetDateTimeValue(strdate)
                        If strdate <> "" Then
                            oEditText.String = dtdate1
                        End If

                    Catch ex As Exception
                    End Try
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'oMatrix.AddRow()
                End If
                oMatrix.Columns.Item("V_8").Editable = False
                AssignLineNo(aForm)
            Catch ex As Exception
            End Try
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub
    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Dim strLineCode, strLineApproval, strEmpApproval As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMatrix = aForm.Items.Item("12").Specific
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(introw).Specific
                strLineApproval = oCombobox.Selected.Value
                oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(introw).Specific
                strEmpApproval = oCombobox.Selected.Value
                '  If oCombobox.Selected.Value = "P" Then
                If strLineApproval = "P" Then
                    strLineCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_6", introw)
                    If strLineCode <> "" Then
                        oTemp.DoQuery("Update [@Z_TIM1] set Name=name +'D' where code='" & strLineCode & "'")
                    End If
                    oMatrix.DeleteRow(introw)
                    AssignLineNo(aForm)
                    Exit Sub
                Else
                    If strEmpApproval = "P" Then
                        strLineCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_6", introw)
                        If strLineCode <> "" Then
                            oTemp.DoQuery("Update [@Z_TIM1] set Name=name +'D' where code='" & strLineCode & "'")
                        End If
                        oMatrix.DeleteRow(introw)
                        AssignLineNo(aForm)
                        Exit Sub
                    End If
                    oApplication.Utilities.Message("You can not delete Approved / Declined time sheet entries", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub

                End If
            End If
        Next

    End Sub
#End Region

#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject As String
        Dim oTemp As SAPbobsCOM.Recordset
        strProject = oApplication.Utilities.getEdittextvalue(aform, "4")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

        End If

        oMatrix = aform.Items.Item("12").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            strCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)
            If strCode <> "" Then
                strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Phase is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_3").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
                strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Activity is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_4").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
                strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_5", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_5").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
                If CDbl(strcode1) <= 0 Then
                    oApplication.Utilities.Message("Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_5").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If

                Dim dblBdgqty, dblQty As Double
                dblBdgqty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "BdgQty", intRow))
                'dblQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "Qty", intRow))
                'If dblBdgqty > 0 Then
                '    If dblQty <= 0 Then
                '        oApplication.Utilities.Message("Quantity is missing. Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        oMatrix.Columns.Item("Qty").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                '        Return False
                '    End If
                'End If


            Else
                oMatrix.DeleteRow(intRow)
            End If
        Next

        oMatrix = aform.Items.Item("12").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Time Sheet details are missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        'AssignLineNo(aform)
        Return True
    End Function
#End Region
#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_EmpTime
                    If pVal.BeforeAction = False Then
                        ' LoadForm()
                        Dim oTe As New clsLogin
                        oTe.LoadForm("Timesheet")

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
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        ' MsgBox(RowtoDelete1)
                        '  RefereshDeleteRow(oForm)
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
                            'If oApplication.SBO_Application.MessageBox("Do you want to remove this project details?", , "Continue", "Cancel") = 1 Then
                            '    Dim objEditText As SAPbouiCOM.EditText
                            '    Dim strDocNum As String
                            '    objEditText = oForm.Items.Item("4").Specific
                            '    strDocNum = objEditText.String
                            'End If
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
        If oApplication.SBO_Application.MessageBox("Do you want to save the timesheet entries?", , "Continue", "Cancel") = 2 Then
            Return False
        End If
        If validation(aForm) = False Then
            Return False
        End If
        If AddToUDT_Table(aForm) = False Then
            Return False
        End If
        ' AssignLineNo(aForm)
        Return True
    End Function



#Region "AddtoUDT"

    Private Function addto_HeaderTable(ByVal aEmpId As String, ByVal aEmpName As String, ByVal aDate As Date) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim oCode As String
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "Select * from [@Z_OTIM] where U_Z_EMPCODE='" & aEmpId & "' and convert(varchar(10),U_Z_DocDate,105)='" & aDate.ToString("dd-MM-yyyy") & "'"
        otemp.DoQuery(strSQL)
        If otemp.RecordCount > 0 Then
            oCode = otemp.Fields.Item("Code").Value
        Else
            oCode = oApplication.Utilities.getMaxCode("@Z_OTIM", "Code")
            Dim ousertable1 As SAPbobsCOM.UserTable
            ousertable1 = oApplication.Company.UserTables.Item("Z_OTIM")
            ousertable1.Code = oCode
            ousertable1.Name = oCode
            ousertable1.UserFields.Fields.Item("U_Z_EMPCODE").Value = aEmpId
            ousertable1.UserFields.Fields.Item("U_Z_EMPNAME").Value = aEmpName
            ousertable1.UserFields.Fields.Item("U_Z_DocDate").Value = aDate
            If ousertable1.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return ""
            Else
                Return oCode
            End If
        End If
        Return oCode
    End Function
    Private Function AddToUDT_Table(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strDocref, strEmpID, strLineCode, stdocdate, strProjectName, strEmpName, strProject, strProcess, strdate, strActivtiy, strhours, stremptype, strPrjCode, strAmount As String
        Dim dtDate As Date
        Dim intHours, dblAmount As Double
        Dim oTempRec, otemp As SAPbobsCOM.Recordset
        Dim ousertable As SAPbobsCOM.UserTable
        Dim ocheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oedittext As SAPbouiCOM.EditTextColumn
        Dim dblPercentage As Double
        Dim blnexits As Boolean = False
        Dim blnLines As Boolean = False
        Dim dtFrom, dtTo As Date
        Dim oBPGrid As SAPbouiCOM.Matrix

        oBPGrid = aform.Items.Item("12").Specific
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strEmpID = oApplication.Utilities.getEdittextvalue(aform, "4")
        strEmpName = oApplication.Utilities.getEdittextvalue(aform, "6")
        strDocref = oApplication.Utilities.getEdittextvalue(aform, "11")
        stdocdate = oApplication.Utilities.getEdittextvalue(aform, "13")
        dtDate = oApplication.Utilities.GetDateTimeValue(stdocdate)
        ousertable = oApplication.Company.UserTables.Item("Z_OTIM")
        blnLines = True
        If blnLines = False Then
            ' oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            For intLoop As Integer = 1 To oBPGrid.RowCount
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                ousertable = oApplication.Company.UserTables.Item("Z_TIM1")
                strPrjCode = oApplication.Utilities.getMatrixValues(oBPGrid, "V_1", intLoop)
                If strPrjCode <> "" Then
                    strProcess = oApplication.Utilities.getMatrixValues(oBPGrid, "V_3", intLoop)
                    strActivtiy = oApplication.Utilities.getMatrixValues(oBPGrid, "V_4", intLoop)
                    strProjectName = oApplication.Utilities.getMatrixValues(oBPGrid, "V_2", intLoop)
                    Dim oEdit As SAPbouiCOM.EditText
                    oEdit = oMatrix.Columns.Item("V_0").Cells.Item(intLoop).Specific
                    strdate = oEdit.String
                    dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                    strhours = oApplication.Utilities.getMatrixValues(oBPGrid, "V_5", intLoop)
                    strLineCode = oApplication.Utilities.getMatrixValues(oBPGrid, "V_6", intLoop)
                    dblAmount = oApplication.Utilities.getDocumentQuantity(strhours)
                    strDocref = addto_HeaderTable(strEmpID, strEmpName, dtDate)
                    If strDocref <> "" Then
                        If strLineCode = "" Then
                            strCode = oApplication.Utilities.getMaxCode("@Z_TIM1", "Code")
                            ousertable.Code = strCode
                            ousertable.Name = strCode
                            ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = strPrjCode
                            ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = strProjectName
                            ousertable.UserFields.Fields.Item("U_Z_PRCNAME").Value = strProcess
                            ousertable.UserFields.Fields.Item("U_Z_ACTNAME").Value = strActivtiy
                            If strdate <> "" Then
                                ousertable.UserFields.Fields.Item("U_Z_DATE").Value = dtDate
                            End If
                            ousertable.UserFields.Fields.Item("U_Z_BDGQTY").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "BdgQty", intLoop))
                            ousertable.UserFields.Fields.Item("U_Z_QUANTITY").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "Qty", intLoop))
                            ousertable.UserFields.Fields.Item("U_Z_MEASURE").Value = oApplication.Utilities.getMatrixValues(oBPGrid, "Measure", intLoop)
                            ousertable.UserFields.Fields.Item("U_Z_HOURS").Value = dblAmount
                            ' oCombobox = oBPGrid.Columns.Item("Type").Cells.Item(intLoop).Specific
                            ousertable.UserFields.Fields.Item("U_Z_TYPE").Value = "R" ' oCombobox.Selected.Value
                            Dim strEmpConfirmaion As String
                            oCombobox = oBPGrid.Columns.Item("V_9").Cells.Item(intLoop).Specific
                            Try
                                ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = oCombobox.Selected.Value
                                strEmpConfirmaion = oCombobox.Selected.Value
                            Catch ex As Exception
                                ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = "P"
                                strEmpConfirmaion = "P"
                            End Try
                            oCombobox = oBPGrid.Columns.Item("V_8").Cells.Item(intLoop).Specific
                            Dim otest As SAPbobsCOM.Recordset
                            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            otest.DoQuery("SELECT T0.[U_Z_PRJCODE], T0.[U_Z_APPROVAL],* FROM [dbo].[@Z_HPRJ]  T0 where T0.U_Z_PRJCODE='" & strPrjCode & "'")
                            If otest.Fields.Item("U_Z_APPROVAL").Value = "N" Then
                                If strEmpConfirmaion = "C" Then
                                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "A" ' oCombobox.Selected.Value
                                Else
                                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "P" ' oCombobox.Selected.Value
                                End If
                            Else
                                ousertable.UserFields.Fields.Item("U_Z_Approved").Value = oCombobox.Selected.Value
                            End If
                            ousertable.UserFields.Fields.Item("U_Z_RefCode").Value = strDocref
                            ousertable.UserFields.Fields.Item("U_Z_Remarks").Value = oApplication.Utilities.getMatrixValues(oBPGrid, "V_7", intLoop)
                            If ousertable.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If

                        Else
                            If ousertable.GetByKey(strLineCode) Then
                                ousertable.Code = strLineCode
                                ousertable.Name = strLineCode
                                ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = strPrjCode
                                ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = strProjectName
                                ousertable.UserFields.Fields.Item("U_Z_PRCNAME").Value = strProcess
                                ousertable.UserFields.Fields.Item("U_Z_ACTNAME").Value = strActivtiy
                                If strdate <> "" Then
                                    dtDate = dtDate ' oApplication.Utilities.GetDateTimeValue(strdate)
                                    ousertable.UserFields.Fields.Item("U_Z_DATE").Value = dtDate
                                End If
                                ousertable.UserFields.Fields.Item("U_Z_BDGQTY").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "BdgQty", intLoop))
                                ousertable.UserFields.Fields.Item("U_Z_QUANTITY").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "Qty", intLoop))
                                ousertable.UserFields.Fields.Item("U_Z_MEASURE").Value = oApplication.Utilities.getMatrixValues(oBPGrid, "Measure", intLoop)
                                '    oCombobox = oBPGrid.Columns.Item("Type").Cells.Item(intLoop).Specific
                                ousertable.UserFields.Fields.Item("U_Z_TYPE").Value = "R"
                                'oCombobox.Selected.Value
                                ousertable.UserFields.Fields.Item("U_Z_HOURS").Value = dblAmount
                                Dim strEmpConfirmaion As String
                                oCombobox = oBPGrid.Columns.Item("V_9").Cells.Item(intLoop).Specific
                                Try
                                    ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = oCombobox.Selected.Value
                                    strEmpConfirmaion = oCombobox.Selected.Value
                                Catch ex As Exception
                                    ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = "P"
                                    strEmpConfirmaion = "P"
                                End Try
                                oCombobox = oBPGrid.Columns.Item("V_8").Cells.Item(intLoop).Specific
                                Dim otest As SAPbobsCOM.Recordset
                                otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                otest.DoQuery("SELECT T0.[U_Z_PRJCODE], T0.[U_Z_APPROVAL],* FROM [dbo].[@Z_HPRJ]  T0 where T0.U_Z_PRJCODE='" & strPrjCode & "'")
                                If otest.Fields.Item("U_Z_APPROVAL").Value = "N" Then
                                    If strEmpConfirmaion = "C" Then
                                        ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "A" ' oCombobox.Selected.Value
                                    Else
                                        ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "P" ' oCombobox.Selected.Value
                                    End If
                                Else
                                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = oCombobox.Selected.Value
                                End If
                                ousertable.UserFields.Fields.Item("U_Z_RefCode").Value = strDocref
                                ousertable.UserFields.Fields.Item("U_Z_Remarks").Value = oApplication.Utilities.getMatrixValues(oBPGrid, "V_7", intLoop)
                                If ousertable.Update <> 0 Then
                                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            oTempRec.DoQuery("Delete from [@Z_TIM1] where name like '%D' and U_Z_RefCode='" & strDocref & "'")
        End If
        oApplication.Utilities.Message("Operation completed successfuly", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function

#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Timesheet Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.CharPressed <> 9 And (pVal.ColUID = "Qty" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5" Or pVal.ColUID = "V_6" Or pVal.ColUID = "V_7") Then
                                    Dim stCode, stCode1 As String
                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_6", pVal.Row)
                                    If stCode <> "" Then
                                        oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                        stCode = oCombobox.Selected.Value

                                        oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                                        stCode1 = oCombobox.Selected.Value
                                        If stCode <> "P" And stCode1 = "C" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And (pVal.ColUID = "Qty" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5" Or pVal.ColUID = "V_6" Or pVal.ColUID = "V_7") Then
                                    Dim stCode, stCode1 As String

                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_6", pVal.Row)

                                    oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                    stCode = oCombobox.Selected.Value
                                    'If stCode <> "P" Then
                                    '    BubbleEvent = False
                                    '    Exit Sub
                                    'End If
                                    oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                                    stCode1 = oCombobox.Selected.Value
                                    If stCode <> "P" And stCode1 = "C" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.Row > 0 And pVal.Row <= oMatrix.RowCount Then
                                    rowtoDelete1 = pVal.Row
                                    Me.MatrixId1 = "12"
                                End If

                                If pVal.ItemUID = "12" And (pVal.ColUID = "Qty" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5" Or pVal.ColUID = "V_6" Or pVal.ColUID = "V_7") Then
                                    Dim stCode, stCode1 As String
                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_6", pVal.Row)
                                    If stCode <> "" Then
                                        oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                        stCode = oCombobox.Selected.Value
                                        'If stCode <> "P" Then
                                        '    BubbleEvent = False
                                        '    Exit Sub
                                        'End If
                                        oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                                        stCode1 = oCombobox.Selected.Value
                                        If stCode <> "P" And stCode1 = "C" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If pVal.ItemUID = "12" And (pVal.ColUID = "Qty" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5" Or pVal.ColUID = "V_6" Or pVal.ColUID = "V_7") Then
                                    Dim stCode, stCode1 As String
                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_6", pVal.Row)
                                    If stCode <> "" Then
                                        oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                        stCode = oCombobox.Selected.Value
                                        'If stCode <> "P" Then
                                        '    BubbleEvent = False
                                        '    Exit Sub
                                        'End If
                                        oCombobox = oMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific
                                        stCode1 = oCombobox.Selected.Value
                                        If stCode <> "P" And stCode1 = "C" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If

                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        Dim obj As New clsLogin
                                        blnSourceForm = True
                                        strSourceformEmpID = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                        oForm.Close()
                                        obj.LoadForm("TimeSheet")
                                    End If
                                ElseIf pVal.ItemUID = "22" Then
                                    AddRow(oForm)
                                ElseIf pVal.ItemUID = "21" Then
                                    deleterow(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" And pVal.ColUID = "V_3" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    oGrid = oForm.Items.Item("12").Specific
                                    strwhs = oApplication.Utilities.getMatrixValues(oGrid, "V_1", pVal.Row)
                                    If strwhs = "" Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getMatrixValues(oGrid, "V_3", pVal.Row)
                                    If oApplication.Utilities.CheckModule_Activity(strwhs, "[@Z_PRJ1]", strGirdValue, "U_Z_MODNAME") = False Then
                                        oApplication.Utilities.SetMatrixValues(oGrid, "V_3", pVal.Row, "")
                                    Else
                                        Dim strProject, strActivtiy, strBusiness, strsql As String
                                        Dim oTe As SAPbobsCOM.Recordset
                                        oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        strProject = oApplication.Utilities.getMatrixValues(oGrid, "V_1", pVal.Row)
                                        strBusiness = oApplication.Utilities.getMatrixValues(oGrid, "V_3", pVal.Row)
                                        strActivtiy = oApplication.Utilities.getMatrixValues(oGrid, "V_4", pVal.Row)

                                        strsql = "Select * from [@Z_PRJ1]" & " where U_Z_MODNAME='" & strBusiness & "' and U_Z_ActName='" & strActivtiy & "' and docentry=(Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & strProject & "')"
                                        oTe.DoQuery(strsql)
                                        If oTe.RecordCount > 0 Then
                                            oForm.Freeze(True)
                                            oCombobox = oGrid.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                            oCombobox.Select(oTe.Fields.Item("U_Z_TYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            '   oApplication.Utilities.SetMatrixValues(oGrid, "BdgQty", pVal.Row, oTe.Fields.Item("U_Z_Quantity").Value)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "BdgQty", oMatrix.RowCount, oTe.Fields.Item("U_Z_HOURS").Value)
                                            oApplication.Utilities.SetMatrixValues(oGrid, "Measure", pVal.Row, oTe.Fields.Item("U_Z_Measure").Value)
                                            oForm.Freeze(False)
                                        End If
                                        Exit Sub
                                    End If
                                    If strwhs <> "" Then
                                        clsChooseFromList.ItemUID = "12"
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = pVal.Row
                                        clsChooseFromList.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "MODULE"
                                        clsChooseFromList.ItemCode = strwhs
                                        clsChooseFromList.sourceItemCode = ""
                                        clsChooseFromList.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                        clsChooseFromList.sourceColumID = pVal.ColUID
                                        clsChooseFromList.sourcerowId = pVal.Row
                                        clsChooseFromList.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If
                                If pVal.ItemUID = "12" And pVal.ColUID = "V_4" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strProcess As String
                                    oGrid = oForm.Items.Item("12").Specific
                                    strwhs = oApplication.Utilities.getMatrixValues(oGrid, "V_1", pVal.Row)
                                    If strwhs = "" Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getMatrixValues(oGrid, "V_4", pVal.Row)
                                    strProcess = oApplication.Utilities.getMatrixValues(oGrid, "V_3", pVal.Row)
                                    If oApplication.Utilities.CheckActivity(strwhs, "[@Z_PRJ1]", strGirdValue, strProcess, "U_Z_ACTNAME", "U_Z_MODNAME") = False Then
                                        oApplication.Utilities.SetMatrixValues(oGrid, "V_4", pVal.Row, "")
                                    Else
                                        Dim strProject, strActivtiy, strBusiness, strsql As String
                                        Dim oTe As SAPbobsCOM.Recordset

                                        oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        strProject = oApplication.Utilities.getMatrixValues(oGrid, "V_1", pVal.Row)
                                        strBusiness = oApplication.Utilities.getMatrixValues(oGrid, "V_3", pVal.Row)
                                        strActivtiy = oApplication.Utilities.getMatrixValues(oGrid, "V_4", pVal.Row)

                                        strsql = "Select * from [@Z_PRJ1]" & " where U_Z_MODNAME='" & strBusiness & "' and U_Z_ActName='" & strActivtiy & "' and docentry=(Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & strProject & "')"
                                        oTe.DoQuery(strSQL)
                                        If oTe.RecordCount > 0 Then
                                            oForm.Freeze(True)
                                            '  MsgBox(oTe.Fields.Item("U_Z_TYPE").Value)

                                            oCombobox = oGrid.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                            oCombobox.Select(oTe.Fields.Item("U_Z_TYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            ' oApplication.Utilities.SetMatrixValues(oGrid, "BdgQty", pVal.Row, oTe.Fields.Item("U_Z_Quantity").Value)
                                            oApplication.Utilities.SetMatrixValues(oGrid, "BdgQty", pVal.Row, oTe.Fields.Item("U_Z_HOURS").Value)
                                            oApplication.Utilities.SetMatrixValues(oGrid, "Measure", pVal.Row, oTe.Fields.Item("U_Z_Measure").Value)
                                            oForm.Freeze(False)
                                        End If
                                        Exit Sub
                                    End If
                                    If strwhs <> "" Then
                                        clsChooseFromList.ItemUID = "12"
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = pVal.Row
                                        clsChooseFromList.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "ACTIVITY"
                                        clsChooseFromList.ItemCode = strwhs
                                        clsChooseFromList.Documentchoice = strProcess  ' oApplication.Utilities.GetDocType(oForm)
                                        clsChooseFromList.sourceColumID = pVal.ColUID
                                        clsChooseFromList.sourcerowId = pVal.Row
                                        clsChooseFromList.sourceItemCode = ""
                                        clsChooseFromList.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val3 As String
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
                                        If pVal.ItemUID = "12" And pVal.ColUID = "V_1" Then
                                            val = oDataTable.GetValue("U_Z_PRJCODE", 0)
                                            val1 = oDataTable.GetValue("U_Z_PRJNAME", 0)
                                            val3 = oDataTable.GetValue("U_Z_APPROVAL", 0)
                                            oMatrix = oForm.Items.Item("12").Specific
                                            oCombobox = oMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific
                                            If val3 <> "N" Then
                                                oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            Else
                                                oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            End If

                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", pVal.Row, val1)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, val)
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
