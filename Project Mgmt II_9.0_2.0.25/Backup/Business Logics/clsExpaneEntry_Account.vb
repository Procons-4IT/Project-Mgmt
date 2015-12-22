Public Class clsExpaneEntry_Account
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

    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_AcctEntry, frm_ACCEntry)
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
            oForm = oApplication.Utilities.LoadForm(xml_AcctEntry, frm_ACCEntry)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            AddChooseFromList(oForm)
            'oForm.PaneLevel = 1
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oApplication.Utilities.setEdittextvalue(oForm, "4", aEmpid)
            oApplication.Utilities.setEdittextvalue(oForm, "6", strName)
            oForm.Items.Item("4").Enabled = False
            If aOption = "A" Then
                databind(oForm, aDate)
            Else
                dtdate = aDate
                'databind_View(oForm, dtdate)
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

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "171"
            oCFLCreationParams.UniqueID = "CFL3"
            oCFL = oCFLs.Add(oCFLCreationParams)



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
            strSQL = "Select * from [@Z_OEXP] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & adate.ToString("dd-MM-yyyy") & "'"
            otemp.DoQuery(strSQL)
            If otemp.RecordCount > 0 Then
                oCode = otemp.Fields.Item("Code").Value
                dtdate1 = otemp.Fields.Item("U_Z_DocDate").Value
                dtdate1 = adate
            Else
                oCode = oApplication.Utilities.getMaxCode("@Z_OEXP", "Code")
                dtdate1 = adate
            End If
            oApplication.Utilities.setEdittextvalue(aForm, "11", oCode)
            'oApplication.Utilities.setEdittextvalue(aForm, "13", Now.Date)
            oApplication.Utilities.setEdittextvalue(aForm, "13", adate)
            oMatrix = aForm.Items.Item("12").Specific
            oEditText = aForm.Items.Item("4").Specific
            oEditText.ChooseFromListUID = "CFL2"
            oEditText.ChooseFromListAlias = "empID"
            oColumn = oMatrix.Columns.Item("V_0")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "U_Z_EXPNAME"

            oColumn = oMatrix.Columns.Item("V_18")
            oColumn.ChooseFromListUID = "CFL3"
            oColumn.ChooseFromListAlias = "empID"

            oColumn = oMatrix.Columns.Item("V_2")
            otemp.DoQuery("Select PrjCode,PrjName from OPRJ order by PrjCode")
            oColumn.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oColumn.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oColumn.DisplayDesc = True

            oColumn = oMatrix.Columns.Item("V_3")
            oColumn.DisplayDesc = True
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_7")
            oColumn.ValidValues.Add("A", "Approved")
            oColumn.ValidValues.Add("D", "Declined")
            oColumn.ValidValues.Add("P", "Approval Pending")
            oColumn.TitleObject.Caption = "Approval Status"
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_6")
            LoadCurrency(oColumn)
            oColumn.DisplayDesc = True

            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oMatrix.AddRow()
            oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Try
                '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_DATE").Value)
                oEditText = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                Dim strdate As String
                strdate = oApplication.Utilities.getDateStrin(dtdate1)
                If strdate <> "" Then
                    oEditText.String = strdate
                End If

            Catch ex As Exception
            End Try
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strSQL = "Select * from [@Z_EXP1] where U_Z_RefCode='" & oCode & "'"
            otemp.DoQuery(strSQL)
            For introw As Integer = 0 To otemp.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, otemp.Fields.Item("U_Z_EXPNAME").Value)
                oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_EXPTYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_PRJCODE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_Approved").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, otemp.Fields.Item("U_Z_AMOUNT").Value)
                'Dim dblAmt As Double
                'dblAmt = otemp.Fields.Item("U_Z_AMOUNT").Value
                'oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, dblAmt.ToString)
                oCombobox = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_CURRENCY").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, otemp.Fields.Item("U_Z_REMARKS").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, otemp.Fields.Item("Code").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, otemp.Fields.Item("U_Z_Ref1").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_20", oMatrix.RowCount, otemp.Fields.Item("U_Z_MODNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_21", oMatrix.RowCount, otemp.Fields.Item("U_Z_ACTNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_18", oMatrix.RowCount, otemp.Fields.Item("U_Z_EMPID").Value)
                Try
                    '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_DATE").Value)
                    oEditText = oMatrix.Columns.Item("V_-2").Cells.Item(oMatrix.RowCount).Specific
                    Dim strdate As String
                    strdate = oApplication.Utilities.getDateStrin(dtdate1)
                    If strdate <> "" Then
                        oEditText.String = strdate
                    End If

                Catch ex As Exception
                End Try

                oMatrix.AddRow()
                otemp.MoveNext()
            Next
            oMatrix.Columns.Item("V_5").Editable = True
            oMatrix.Columns.Item("V_8").Editable = True
            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
            oMatrix.AutoResizeColumns()
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
            strEmpID = oApplication.Utilities.getEdittextvalue(aForm, "4")
            strSQL = "Select * from [@Z_OEXP] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & adate.ToString("dd-MM-yyyy") & "'"
            otemp.DoQuery(strSQL)
            If otemp.RecordCount > 0 Then
                oCode = otemp.Fields.Item("Code").Value
                dtdate1 = otemp.Fields.Item("U_Z_DocDate").Value
                dtdate1 = adate
            Else
                oCode = oApplication.Utilities.getMaxCode("@Z_OEXP", "Code")
                dtdate1 = adate
            End If
            oApplication.Utilities.setEdittextvalue(aForm, "11", oCode)
            'oApplication.Utilities.setEdittextvalue(aForm, "13", Now.Date)
            oApplication.Utilities.setEdittextvalue(aForm, "13", adate)
            oMatrix = aForm.Items.Item("12").Specific
            oEditText = aForm.Items.Item("4").Specific
            oEditText.ChooseFromListUID = "CFL2"
            oEditText.ChooseFromListAlias = "empID"
            oColumn = oMatrix.Columns.Item("V_0")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "U_Z_EXPNAME"

            oColumn = oMatrix.Columns.Item("V_18")
            oColumn.ChooseFromListUID = "CFL3"
            oColumn.ChooseFromListAlias = "empID"
            oColumn = oMatrix.Columns.Item("V_2")
            otemp.DoQuery("Select PrjCode,PrjName from OPRJ order by PrjCode")
            oColumn.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oColumn.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oColumn.DisplayDesc = True

            oColumn = oMatrix.Columns.Item("V_3")
            oColumn.DisplayDesc = True
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_7")
            oColumn.ValidValues.Add("A", "Approved")
            oColumn.ValidValues.Add("D", "Declined")
            oColumn.ValidValues.Add("P", "Approval Pending")
            oColumn.TitleObject.Caption = "Approval Status"
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_6")
            LoadCurrency(oColumn)
            oColumn.DisplayDesc = True

            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oMatrix.AddRow()
            oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Try
                '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_DATE").Value)
                oEditText = oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific
                Dim strdate As String
                strdate = oApplication.Utilities.getDateStrin(dtdate1)
                If strdate <> "" Then
                    oEditText.String = strdate
                End If

            Catch ex As Exception
            End Try
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            '  strSQL = "Select * from [@Z_EXP1] where U_Z_RefCode='" & oCode & "'"
            strSQL = "Select * from [@Z_EXP1] where isnull(U_Z_APPROVED,'P')='P' and U_Z_RefCode in (Select Code from [@Z_OEXP] where  U_Z_EMPCODE='" & strEmpID & " ') order by U_Z_DATE,Code"
            otemp.DoQuery(strSQL)
            For introw As Integer = 0 To otemp.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, otemp.Fields.Item("U_Z_EXPNAME").Value)
                oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_EXPTYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_PRJCODE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_Approved").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, otemp.Fields.Item("U_Z_AMOUNT").Value)
                'Dim dblAmt As Double
                'dblAmt = otemp.Fields.Item("U_Z_AMOUNT").Value
                'oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, dblAmt.ToString)
                oCombobox = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select(otemp.Fields.Item("U_Z_CURRENCY").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, otemp.Fields.Item("U_Z_REMARKS").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, otemp.Fields.Item("Code").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, otemp.Fields.Item("U_Z_Ref1").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_20", oMatrix.RowCount, otemp.Fields.Item("U_Z_MODNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_21", oMatrix.RowCount, otemp.Fields.Item("U_Z_ACTNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_18", oMatrix.RowCount, otemp.Fields.Item("U_Z_EMPID").Value)
                Try
                    '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_DATE").Value)
                    oEditText = oMatrix.Columns.Item("V_-2").Cells.Item(oMatrix.RowCount).Specific
                    Dim strdate As String
                    dtdate1 = otemp.Fields.Item("U_Z_DATE").Value
                    dtdate1 = dtdate1

                    strdate = oApplication.Utilities.getDateStrin(dtdate1)
                    If strdate <> "" Then
                        oEditText.String = strdate
                    End If

                Catch ex As Exception
                End Try

                oMatrix.AddRow()
                otemp.MoveNext()
            Next
            oMatrix.Columns.Item("V_5").Editable = True
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
            '  strSQL = "Select * from [@Z_OEXP] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & dtDate.ToString("dd-MM-yyyy") & "'"
            'strSQL = "Select * from [@Z_OEXP] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and U_Z_DocDate='" & dtDate.ToString("yyyy-MM-dd") & "'"
            strSQL = "Select * from [@Z_OEXP] where U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and convert(varchar(10),U_Z_DocDate,105)='" & dtDate.ToString("dd-MM-yyyy") & "'"
            otemp.DoQuery(strSQL)
            If otemp.RecordCount > 0 Then
                oCode = otemp.Fields.Item("Code").Value
            Else

                oCode = oApplication.Utilities.getMaxCode("@Z_OEXP", "Code")
            End If
            oApplication.Utilities.setEdittextvalue(aForm, "11", oCode)
            oApplication.Utilities.setEdittextvalue(aForm, "13", dtDate)
            oMatrix = aForm.Items.Item("12").Specific
            oEditText = aForm.Items.Item("4").Specific
            oEditText.ChooseFromListUID = "CFL2"
            oEditText.ChooseFromListAlias = "empID"
            oColumn = oMatrix.Columns.Item("V_0")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "U_Z_EXPNAME"
            oColumn = oMatrix.Columns.Item("V_18")
            oColumn.ChooseFromListUID = "CFL3"
            oColumn.ChooseFromListAlias = "empID"

            oColumn = oMatrix.Columns.Item("V_2")
            otemp.DoQuery("Select PrjCode,PrjName from OPRJ order by PrjCode")
            oColumn.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oColumn.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oColumn.DisplayDesc = True

            oColumn = oMatrix.Columns.Item("V_3")
            oColumn.DisplayDesc = True
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_7")
            oColumn.ValidValues.Add("A", "Approved")
            oColumn.ValidValues.Add("D", "Declined")
            oColumn.ValidValues.Add("P", "Approval Pending")
            oColumn.TitleObject.Caption = "Approved"
            oColumn.DisplayDesc = True
            oColumn = oMatrix.Columns.Item("V_6")
            LoadCurrency(oColumn)
            oColumn.DisplayDesc = True

            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oMatrix.AddRow()
            oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(1).Specific
            oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(1).Specific
            oCombobox.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strSQL = "Select * from [@Z_EXP1] where U_Z_RefCode='" & oCode & "'"
            otemp.DoQuery(strSQL)
            For introw As Integer = 0 To otemp.RecordCount - 1
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, otemp.Fields.Item("U_Z_EXPNAME").Value)
                oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oCombobox.Select(otemp.Fields.Item("U_Z_EXPTYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception
                End Try
                oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oCombobox.Select(otemp.Fields.Item("U_Z_PRJCODE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try

                oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oCombobox.Select(otemp.Fields.Item("U_Z_Approved").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try

                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, otemp.Fields.Item("U_Z_AMOUNT").Value)
                oCombobox = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oCombobox.Select(otemp.Fields.Item("U_Z_CURRENCY").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try

                oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, otemp.Fields.Item("U_Z_REMARKS").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, otemp.Fields.Item("Code").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, otemp.Fields.Item("U_Z_Ref1").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_20", oMatrix.RowCount, otemp.Fields.Item("U_Z_MODNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_21", oMatrix.RowCount, otemp.Fields.Item("U_Z_ACTNAME").Value)
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_18", oMatrix.RowCount, otemp.Fields.Item("U_Z_EMPID").Value)
                Try
                    oEditText = oMatrix.Columns.Item("V_-2").Cells.Item(oMatrix.RowCount).Specific
                    Dim strdate As String
                    strdate = oApplication.Utilities.getDateStrin(dtDate)
                    If strdate <> "" Then
                        oEditText.String = strdate
                    End If
                Catch ex As Exception
                End Try
                oMatrix.AddRow()
                otemp.MoveNext()
            Next
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


    Public Sub DataView(ByVal aEmpID As String, ByVal aName As String, ByVal dtDate As Date)
        Try
            Dim oCode As String
            Dim strSQL As String
            Dim otemp As SAPbobsCOM.Recordset
            oForm = oApplication.Utilities.LoadForm(xml_ExpEntry, frm_ExpEntry)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            AddChooseFromList(oForm)
            oForm.PaneLevel = 1
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oApplication.Utilities.setEdittextvalue(oForm, "4", aEmpID)
            oApplication.Utilities.setEdittextvalue(oForm, "6", aName)
            oForm.Items.Item("4").Enabled = False
            ' oForm.Items.Item("6").Enabled = False
            databind_View(oForm, dtDate)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Load Currencies"
    Private Sub LoadCurrency(ByVal aCombo As SAPbouiCOM.Column)
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("SELECT T0.[CurrCode], T0.[CurrName] FROM OCRN T0 order by T0.[CurrCode]")
        Try
            For intRow As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
                aCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
        Catch ex As Exception

        End Try
        Try
            aCombo.ValidValues.Add("", "")
            For intRow As Integer = 0 To oTemp.RecordCount - 1
                aCombo.ValidValues.Add(oTemp.Fields.Item(0).Value, oTemp.Fields.Item(1).Value)
                oTemp.MoveNext()
            Next

        Catch ex As Exception

        End Try

    End Sub
#End Region
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("12").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_EXP1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()

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
            Dim strEmp As String = oApplication.Utilities.getEdittextvalue(aForm, "4")
            oMatrix = aForm.Items.Item("12").Specific
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
                oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Try
                    oEditText = oMatrix.Columns.Item("V_-2").Cells.Item(oMatrix.RowCount).Specific
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
            '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_18", oMatrix.RowCount, strEmp)
            Try
                oEditText = oMatrix.Columns.Item("V_-2").Cells.Item(oMatrix.RowCount).Specific
                If oEditText.Value <> "" Then
                    oMatrix.AddRow()
                    oMatrix.ClearRowData(oMatrix.RowCount)
                    'oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, "")
                    'oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                    'oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, "")
                    'oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, "0")
                    'oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                    'oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    'oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                    'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    'oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                    'oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    'oCombobox = oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific
                    'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    Try
                        '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otemp.Fields.Item("U_Z_DATE").Value)
                        oEditText = oMatrix.Columns.Item("V_-2").Cells.Item(oMatrix.RowCount).Specific
                        Dim strdate As String
                        Dim dtdate1 As Date
                        strdate = oApplication.Utilities.getEdittextvalue(aForm, "13")
                        dtdate1 = oApplication.Utilities.GetDateTimeValue(strdate)
                        If strdate <> "" Then
                            oEditText.String = dtdate1
                        End If

                    Catch ex As Exception
                    End Try
                    ' oApplication.Utilities.SetMatrixValues(oMatrix, "V_18", oMatrix.RowCount, strEmp)
                    oMatrix.Columns.Item("V_-2").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            Catch ex As Exception
                ' oMatrix.AddRow()
            End Try
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub
    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Dim strLineCode, strApproved As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMatrix = aForm.Items.Item("12").Specific
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(introw).Specific
                If oCombobox.Selected.Value = "P" Then
                    strLineCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", introw)
                    If strLineCode <> "" Then
                        oTemp.DoQuery("Update [@Z_EXP1] set Name=name +'D' where code='" & strLineCode & "'")
                    End If
                    oMatrix.DeleteRow(introw)
                Else
                    oApplication.Utilities.Message("You can not delete the Approved / Declined entry details", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                Exit Sub
            End If
        Next

    End Sub

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)

    End Sub
#End Region
#End Region

#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject As String
        Dim oTemp As SAPbobsCOM.Recordset
        strProject = oApplication.Utilities.getEdittextvalue(aform, "4")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            'oTemp.DoQuery("Select * from [@Z_OEXP] where U_Z_EMPCODE='" & strProject & "'")
            'If oTemp.RecordCount > 0 Then
            '    oApplication.Utilities.Message("Employee Expense details already exists...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
        End If

        oMatrix = aform.Items.Item("12").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            strCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
            If strCode <> "" Then

                oCombobox = oMatrix.Columns.Item("V_6").Cells.Item(intRow).Specific
                If oCombobox.Selected.Value = "" Then
                    oApplication.Utilities.Message("Transaction Currency is missing. Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Amount is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If


                If CDbl(strcode1) <= 0 Then
                    oApplication.Utilities.Message("Amount is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_18", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Employee ID is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(intRow).Specific
                Try
                    strcode1 = oCombobox.Selected.Value
                Catch ex As Exception
                    strcode1 = "N"
                End Try
                If strcode1 = "P" Then
                    oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(intRow).Specific
                    If oCombobox.Selected.Value = "" Then
                        oApplication.Utilities.Message("Project code is missing. Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If

            Else
                oMatrix.DeleteRow(intRow)
            End If
        Next

        oMatrix = aform.Items.Item("12").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Expances details are missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                Case mnu_ACCClaim
                    If pVal.BeforeAction = False Then
                        ' LoadForm()
                        Dim oTe As New clsLogin
                        oTe.LoadForm("Acct")

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
                        RefereshDeleteRow(oForm)
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
        strSQL = "Select * from [@Z_OEXP] where U_Z_EMPCODE='" & aEmpId & "' and convert(varchar(10),U_Z_DocDate,105)='" & aDate.ToString("dd-MM-yyyy") & "'"
        oTemp.DoQuery(strSQL)
        If oTemp.RecordCount > 0 Then
            oCode = oTemp.Fields.Item("Code").Value
        Else
            oCode = oApplication.Utilities.getMaxCode("@Z_OEXP", "Code")
            Dim ousertable1 As SAPbobsCOM.UserTable
            ousertable1 = oApplication.Company.UserTables.Item("Z_OEXP")
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
        Dim dtFrom, dtTo As Date
        Dim oBPGrid As SAPbouiCOM.Matrix
        Dim strRef1 As String

        oBPGrid = aform.Items.Item("12").Specific
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strEmpID = oApplication.Utilities.getEdittextvalue(aform, "4")
        strEmployeename = oApplication.Utilities.getEdittextvalue(aform, "6")
        strDocref = oApplication.Utilities.getEdittextvalue(aform, "11")
        stdocdate = oApplication.Utilities.getEdittextvalue(aform, "13")
        dtDate = oApplication.Utilities.GetDateTimeValue(stdocdate)
        ousertable = oApplication.Company.UserTables.Item("Z_OEXP")

        blnLines = True
        If blnLines = False Then
            ' oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            For intLoop As Integer = 1 To oBPGrid.RowCount
                Dim strRemarks, strCurrency, strdate, strMod, strAct As String
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                ousertable = oApplication.Company.UserTables.Item("Z_EXP1")
                strEmpName = oApplication.Utilities.getMatrixValues(oBPGrid, "V_0", intLoop)
                If strEmpName <> "" Then
                    strLineCode = oApplication.Utilities.getMatrixValues(oBPGrid, "V_4", intLoop)
                    strRemarks = oApplication.Utilities.getMatrixValues(oBPGrid, "V_5", intLoop)
                    oCombobox = oBPGrid.Columns.Item("V_6").Cells.Item(intLoop).Specific
                    strCurrency = oCombobox.Selected.Value
                    oCombobox = oBPGrid.Columns.Item("V_3").Cells.Item(intLoop).Specific
                    stremptype = oCombobox.Selected.Value
                    If stremptype = "P" Then
                        oCombobox = oBPGrid.Columns.Item("V_2").Cells.Item(intLoop).Specific
                        strPrjCode = oCombobox.Selected.Value
                        strprojectname = oCombobox.Selected.Description
                    Else
                        strPrjCode = ""
                        strprojectname = ""
                    End If

                    Dim oEdit As SAPbouiCOM.EditText
                    oEdit = oMatrix.Columns.Item("V_-2").Cells.Item(intLoop).Specific
                    strdate = oEdit.String
                    dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                    strAmount = oApplication.Utilities.getMatrixValues(oBPGrid, "V_1", intLoop)
                    dblAmount = oApplication.Utilities.getDocumentQuantity(strAmount)
                    strDocref = addto_HeaderTable(strEmpID, strEmployeename, dtDate)
                    strRef1 = oApplication.Utilities.getMatrixValues(oBPGrid, "V_8", intLoop)
                    strMod = oApplication.Utilities.getMatrixValues(oBPGrid, "V_20", intLoop)
                    strAct = oApplication.Utilities.getMatrixValues(oBPGrid, "V_21", intLoop)
                    Dim strEmpid1 As String = oApplication.Utilities.getMatrixValues(oBPGrid, "V_18", intLoop)
                    Dim ot As SAPbobsCOM.Recordset
                    ot = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If strEmpid1 = "" Then
                        strEmpid1 = strEmpID
                    End If
                    ot.DoQuery("Select * from OHEM where ""empID""=" & strEmpid1)
                    strEmployeename = ot.Fields.Item("firstName").Value


                    If strDocref <> "" Then
                        If strLineCode = "" Then
                            strCode = oApplication.Utilities.getMaxCode("@Z_EXP1", "Code")
                            ousertable.Code = strCode
                            ousertable.Name = strCode
                            ousertable.UserFields.Fields.Item("U_Z_EXPNAME").Value = strEmpName
                            ousertable.UserFields.Fields.Item("U_Z_EXPTYPE").Value = stremptype
                            ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = strPrjCode
                            ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = strprojectname
                            ousertable.UserFields.Fields.Item("U_Z_CURRENCY").Value = strCurrency
                            If strdate <> "" Then
                                ousertable.UserFields.Fields.Item("U_Z_DATE").Value = dtDate
                            End If
                            ousertable.UserFields.Fields.Item("U_Z_AMOUNT").Value = dblAmount
                            ousertable.UserFields.Fields.Item("U_Z_RefCode").Value = strDocref
                            ousertable.UserFields.Fields.Item("U_Z_REMARKS").Value = strRemarks
                            ousertable.UserFields.Fields.Item("U_Z_REF1").Value = strRef1
                            ousertable.UserFields.Fields.Item("U_Z_ACTNAME").Value = strAct
                            ousertable.UserFields.Fields.Item("U_Z_MODNAME").Value = strMod
                            ousertable.UserFields.Fields.Item("U_Z_EMPID").Value = strEmpid1 'oApplication.Utilities.getMatrixValues(oBPGrid, "V_18", intLoop)
                            ousertable.UserFields.Fields.Item("U_Z_EMPNAME").Value = strEmployeename
                            If ousertable.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            If ousertable.GetByKey(strLineCode) Then
                                strRemarks = oApplication.Utilities.getMatrixValues(oBPGrid, "V_5", intLoop)

                                ousertable.Code = strLineCode
                                ousertable.Name = strLineCode
                                ousertable.UserFields.Fields.Item("U_Z_EXPNAME").Value = strEmpName
                                ousertable.UserFields.Fields.Item("U_Z_EXPTYPE").Value = stremptype
                                ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = strPrjCode
                                ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = strprojectname
                                ousertable.UserFields.Fields.Item("U_Z_CURRENCY").Value = strCurrency
                                If strdate <> "" Then
                                    ousertable.UserFields.Fields.Item("U_Z_DATE").Value = dtDate
                                End If
                                ousertable.UserFields.Fields.Item("U_Z_AMOUNT").Value = dblAmount
                                ousertable.UserFields.Fields.Item("U_Z_RefCode").Value = strDocref
                                ousertable.UserFields.Fields.Item("U_Z_REMARKS").Value = strRemarks
                                ousertable.UserFields.Fields.Item("U_Z_REF1").Value = strRef1
                                ousertable.UserFields.Fields.Item("U_Z_ACTNAME").Value = strAct
                                ousertable.UserFields.Fields.Item("U_Z_MODNAME").Value = strMod
                                ousertable.UserFields.Fields.Item("U_Z_EMPID").Value = strEmpid1 ' oApplication.Utilities.getMatrixValues(oBPGrid, "V_18", intLoop)
                                ousertable.UserFields.Fields.Item("U_Z_EMPNAME").Value = strEmployeename
                                If ousertable.Update <> 0 Then
                                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            oApplication.Utilities.updatelocalamount(strDocref)

            oTempRec.DoQuery("Delete from [@Z_EXP1] where name like '%D' and U_Z_RefCode='" & strDocref & "'")
        End If
        oApplication.Utilities.Message("Operation completed successfuly", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ACCEntry Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.CharPressed <> 9 And (pVal.ColUID = "V_-2" Or pVal.ColUID = "V_20" Or pVal.ColUID = "V_21" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5") Then
                                    Dim stCode As String
                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", pVal.Row)
                                    oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                    stCode = oCombobox.Selected.Value
                                    If stCode <> "P" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.CharPressed <> 9 And (pVal.ColUID = "V_-2" Or pVal.ColUID = "V_20" Or pVal.ColUID = "V_21" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4" Or pVal.ColUID = "V_5") Then
                                    Dim stCode As String
                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", pVal.Row)
                                    oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                    stCode = oCombobox.Selected.Value
                                    'If stCode <> "P" And stCode <> "" Then
                                    '    BubbleEvent = False
                                    '    Exit Sub
                                    'End If
                                    'oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                    'stCode = oCombobox.Selected.Value
                                    'If stCode <> "P" Then
                                    '    BubbleEvent = False
                                    '    Exit Sub
                                    'End If

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And (pVal.ColUID = "V_2" Or pVal.ColUID = "V_3") Then
                                    Dim stCode As String
                                    Try
                                        stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", pVal.Row)
                                        oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                        stCode = oCombobox.Selected.Value
                                        If stCode <> "P" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try


                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.Row > 0 And pVal.Row <= oMatrix.RowCount Then
                                    rowtoDelete1 = pVal.Row
                                    Me.MatrixId1 = "12"
                                End If
                                If pVal.ItemUID = "12" And (pVal.ColUID = "V_-2" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_6" Or pVal.ColUID = "V_5") Then
                                    Dim stCode As String
                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", pVal.Row)
                                    oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                    stCode = oCombobox.Selected.Value
                                    If stCode <> "P" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If pVal.ItemUID = "12" And (pVal.ColUID = "V_-2" Or pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_2" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_6") Then
                                    Dim stCode As String
                                    stCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", pVal.Row)
                                    oCombobox = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                    stCode = oCombobox.Selected.Value
                                    If stCode <> "P" Then
                                        BubbleEvent = False
                                        Exit Sub
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
                                        obj.LoadForm("Acct")
                                    End If
                                ElseIf pVal.ItemUID = "22" Then
                                    AddRow(oForm)
                                ElseIf pVal.ItemUID = "21" Then
                                    deleterow(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" And (pVal.ColUID = "V_20" Or pVal.ColUID = "V_21") And pVal.CharPressed = 9 Then
                                    If 1 = 1 Then
                                        Dim objChooseForm As SAPbouiCOM.Form
                                        Dim objChoose As New clsChooseFromList
                                        Dim strwhs, strProject, strGirdValue As String
                                        oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        Dim strprojectname, strbusinessprocess, strActvity, strRef As String
                                        oMatrix = oForm.Items.Item("12").Specific
                                        oCombobox = oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific
                                        If oCombobox.Selected.Value = "P" Then
                                            oCombobox = oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific
                                            strProject = oCombobox.Selected.Value
                                            strprojectname = oCombobox.Selected.Description
                                            strbusinessprocess = oApplication.Utilities.getMatrixValues(oMatrix, "V_20", pVal.Row)
                                            strActvity = oApplication.Utilities.getMatrixValues(oMatrix, "V_21", pVal.Row)
                                            If strProject = "" Then
                                                Exit Sub
                                            End If
                                            strGirdValue = oApplication.Utilities.getMatrixValues(oMatrix, "V_21", pVal.Row)
                                            If oApplication.Utilities.CheckModule_Activity(strProject, "[@Z_PRJ1]", strGirdValue, "U_Z_ACTNAME") = False Then
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_21", pVal.Row, "")
                                            Else
                                                Exit Sub
                                            End If
                                            If strProject <> "" Then
                                                clsChooseFromList.ItemUID = pVal.ItemUID
                                                clsChooseFromList.SourceFormUID = FormUID
                                                clsChooseFromList.SourceLabel = pVal.Row
                                                clsChooseFromList.sourceItemCode = "" ' oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                                clsChooseFromList.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                                clsChooseFromList.choice = "ACTIVITY3"
                                                clsChooseFromList.ItemCode = strProject
                                                clsChooseFromList.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                                clsChooseFromList.sourceColumID = pVal.ColUID
                                                clsChooseFromList.sourcerowId = pVal.Row
                                                clsChooseFromList.BinDescrUID = ""
                                                oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                                objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                                objChoose.databound(objChooseForm)
                                            End If
                                        End If
                                    End If
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
                                        If pVal.ItemUID = "12" And pVal.ColUID = "V_18" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            oMatrix = oForm.Items.Item("12").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_18", pVal.Row, val)
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
