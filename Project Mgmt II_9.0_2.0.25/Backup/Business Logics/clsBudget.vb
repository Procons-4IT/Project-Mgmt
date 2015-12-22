Public Class clsBudget
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oColumn As SAPbouiCOM.Column
    Private oCheckBox As SAPbouiCOM.CheckBox
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
    Private RowtoDelete As Integer
    Private MatrixId As String
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource
#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_Budget, frm_Budget)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "4"
        FillProjectCode(oForm)
        AddChooseFromList(oForm)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu(mnu_Duplicate_Row, True)
        oForm.PaneLevel = 1
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next

        'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        'For count = 1 To oDataSrc_Line.Size - 1
        '    oDataSrc_Line.SetValue("LineId", count - 1, count)
        'Next
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        ' oForm.Items.Item("6").Enabled = False
        databind(oForm)
        oForm.Freeze(False)
    End Sub

#Region "Fill Project Code"
    Private Sub FillProjectCode(ByVal aForm As SAPbouiCOM.Form)
        Dim oTempRec As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCombobox = aForm.Items.Item("4").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("", "")
        oTempRec.DoQuery("Select PrjCode,Prjname from OPRJ order by PrjCode")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            oCombobox.ValidValues.Add(oTempRec.Fields.Item("PrjCode").Value, oTempRec.Fields.Item("PrjName").Value)
            oTempRec.MoveNext()
        Next
    End Sub
#End Region
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
            oCFLCreationParams.ObjectType = "Z_Module"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "Z_Module"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFLCreationParams.ObjectType = "Z_Activity"
            oCFLCreationParams.UniqueID = "CFL3"
            oCFL = oCFLs.Add(oCFLCreationParams)

            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "CFL4"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "171"
            oCFLCreationParams.UniqueID = "CFL5"
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AddChooseFromList_Conditions(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCombobox = objForm.Items.Item("11").Specific
            oCFL = oCFLs.Item("CFL11")
            oCons = oCFL.GetConditions()
            oCon = oCons.Item(0)
            oCon.Alias = "U_Z_CompNo"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = oCombobox.Selected.Value
            oCFL.SetConditions(oCons)
            ' oCon = oCons.Add
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("12").Specific
            oColumn = oMatrix.Columns.Item("V_0")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "U_Z_MODNAME"

            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("V_2")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("V_8")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("Qty")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


            oMatrix = aForm.Items.Item("12").Specific
            oColumn = oMatrix.Columns.Item("V_3")
            oColumn.ChooseFromListUID = "CFL3"
            oColumn.ChooseFromListAlias = "U_Z_ActName"

            oMatrix = aForm.Items.Item("12").Specific
            oColumn = oMatrix.Columns.Item("V_6")
            oColumn.ChooseFromListUID = "CFL4"
            oColumn.ChooseFromListAlias = "DocEntry"

            'oColumn = oMatrix.Columns.Item("V_4")
            'Dim oComborec As SAPbobsCOM.Recordset
            'oComborec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ''  oComborec.DoQuery("SELECT T0.[posID], T0.[name] FROM OHPS T0 order by T0.[posID]  ")
            'oComborec.DoQuery("SELECT T0.[empID],  T0.[firstName] + ' ' + T0.[lastName] FROM OHEM T0 order by T0.[empID]  ")
            'For introw As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
            '    oColumn.ValidValues.Remove(introw)
            'Next
            'oColumn.ValidValues.Add("", "")
            'For introw As Integer = 0 To oComborec.RecordCount - 1
            '    Try
            '        oColumn.ValidValues.Add(oComborec.Fields.Item(0).Value, oComborec.Fields.Item(1).Value)
            '    Catch ex As Exception
            '    End Try
            '    oComborec.MoveNext()
            'Next

            oColumn.DisplayDesc = True
            oMatrix = aForm.Items.Item("12").Specific
            oColumn = oMatrix.Columns.Item("V_4")
            oColumn.ChooseFromListUID = "CFL5"
            oColumn.ChooseFromListAlias = "empID"
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
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
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
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
        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("12").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
                'Case "2"
                '    oMatrix = aForm.Items.Item("13").Specific
                '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        End Select
        Try
            aForm.Freeze(True)
            
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If
            oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
            Try
                If oEditText.Value <> "" Then
                    oMatrix.AddRow()
                    Select Case aForm.PaneLevel
                        Case "1"
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_11", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, "0")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, "0")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, "0")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_9", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_10", oMatrix.RowCount, "")
                            oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
                            oCheckBox.Checked = False
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "BOQ", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "Qty", oMatrix.RowCount, "0")
                    End Select
                End If
            Catch ex As Exception
                aForm.Freeze(False)
                If oMatrix.RowCount <= 0 Then
                    oMatrix.AddRow()
                End If
            End Try

            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)

        End Try
    End Sub



    Private Sub DuplicateRow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("12").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        End Select
        Try
            aForm.Freeze(True)

            For intRow As Integer = 1 To oMatrix.RowCount
                If oMatrix.IsRowSelected(intRow) Then
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific
                    Try
                        If oEditText.Value <> "" Then
                            oMatrix.AddRow()
                            Select Case aForm.PaneLevel
                                Case "1"
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific.value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, oMatrix.Columns.Item("V_3").Cells.Item(intRow).Specific.value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, oMatrix.Columns.Item("V_4").Cells.Item(intRow).Specific.value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_11", oMatrix.RowCount, oMatrix.Columns.Item("V_11").Cells.Item(intRow).Specific.value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, oMatrix.Columns.Item("V_1").Cells.Item(intRow).Specific.Value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, oMatrix.Columns.Item("V_2").Cells.Item(intRow).Specific.Value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, oMatrix.Columns.Item("V_8").Cells.Item(intRow).Specific.Value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_9", oMatrix.RowCount, "")
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "BOQ", oMatrix.RowCount, "")
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_10", oMatrix.RowCount, "")
                                    oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.introw).Specific
                                    If oCheckBox.Checked = True Then
                                        oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
                                        oCheckBox.Checked = True
                                    Else
                                        oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
                                        oCheckBox.Checked = False
                                    End If
                                    'oCheckBox.Checked = False
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, oMatrix.Columns.Item("V_6").Cells.Item(intRow).Specific.Value)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, oMatrix.Columns.Item("V_7").Cells.Item(intRow).Specific.Value)
                                    oCombobox = oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific
                                    oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End Select
                        End If
                    Catch ex As Exception
                        aForm.Freeze(False)
                        If oMatrix.RowCount <= 0 Then
                            oMatrix.AddRow()
                        End If
                    End Try
                End If
            Next


            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)

        End Try
    End Sub
    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("12").Specific
                'Case "2"
                '    oMatrix = aForm.Items.Item("13").Specific
        End Select
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oMatrix.DeleteRow(introw)
            End If
        Next

    End Sub

    Private Function ValidateDeletion(ByVal aForm As SAPbouiCOM.Form) As Boolean
        If intSelectedMatrixrow <= 0 Then
            Return True
        End If
        oMatrix = frmSourceMatrix
        Dim strPrjCode, strActivity, strProcess, strMessage As String
        Dim otemp As SAPbobsCOM.Recordset
        strMessage = ""
        If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intSelectedMatrixrow) <> "" Then
            oCombobox = aForm.Items.Item("4").Specific
            strPrjCode = oCombobox.Selected.Value
            strProcess = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intSelectedMatrixrow)
            strActivity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intSelectedMatrixrow)
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("select * from [@Z_OEXP] T0 Inner Join [@Z_EXP1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strPrjCode & "'")
            strMessage = "Project Code=" & strPrjCode
            'If otemp.RecordCount > 0 Then
            '    oApplication.Utilities.Message("Expenses already entered for this : " & strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            otemp.DoQuery("select * from [@Z_OTIM] T0 Inner Join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strPrjCode & "' and U_Z_PRCNAME='" & strProcess & "' and U_Z_ACTNAME='" & strActivity & "'")
            If otemp.RecordCount > 0 Then
                strMessage = "Project Code : " & strPrjCode & " , Phase : " & strProcess & " , Activity : " & strActivity
                oApplication.Utilities.Message("Time Sheet already entered for this " & strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Return True


    End Function

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        If Me.MatrixId = "12" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
            'Else
            '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        End If
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        If intSelectedMatrixrow <= 0 Then
            Exit Sub

        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        oMatrix = frmSourceMatrix
        oMatrix.FlushToDataSource()
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(oMatrix.RowCount)
        End If
    End Sub
#End Region
#End Region

#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject, strActivity, strActivity1 As String
        Dim oTemp As SAPbobsCOM.Recordset
        oCombobox = aform.Items.Item("4").Specific
        strProject = oCombobox.Selected.Value
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If strProject = "" Then
                oApplication.Utilities.Message("Project code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oTemp.DoQuery("Select * from [@Z_HPRJ] where U_Z_PRJCODE='" & strProject & "'")
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Project code already exists...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Dim dblBudget, dblLineBudget As Double
        Dim strHours As String
        Dim dblHours, dblRate As Double
        dblBudget = 0
        dblLineBudget = 0
        dblRate = 0
        oMatrix = aform.Items.Item("12").Specific
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        For intRow As Integer = 1 To oMatrix.RowCount
            strCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
            If strCode <> "" Then
                strActivity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intRow)
                strHours = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", intRow)
                If strHours <> "" Then
                    dblHours = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_2", intRow))
                    If (dblHours > 0) Then
                        'dblHours = dblHours * 8
                        ' oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", intRow, dblHours)
                        dblHours = dblHours * 8
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oEditText = oMatrix.Columns.Item("V_4").Cells.Item(intRow).Specific
                        If oEditText.String <> "" Then
                            oTest.DoQuery("Select isnull(U_Daily_rate,0) from OHEM where empID=" & oEditText.String)
                            dblRate = oTest.Fields.Item(0).Value
                        Else
                            dblRate = 1
                        End If
                        dblRate = dblRate * oApplication.Utilities.getDocumentQuantity(strHours)
                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", intRow, dblHours)
                        'oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", intRow, dblRate)
                    End If
                End If

                strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)
                oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific
                If oCombobox.Selected.Value = "R" Then


                    If strcode1 = "" Then
                        oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Return False
                    End If
                    If CInt(strcode1) <= 0 Then
                        oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Return False
                    End If
                End If
                If strActivity = "" Then
                    oApplication.Utilities.Message("Activity detail is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific
                If oCombobox.Selected.Value = "I" Then
                    If oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "Qty", intRow)) <= 0 Then
                        oApplication.Utilities.Message("Required Quantity should be greater than or equal to Zero. Line number:" & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("Qty").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Return False
                    End If
                    If oApplication.Utilities.getMatrixValues(oMatrix, "BOQ", intRow) = "" Then
                        '  oApplication.Utilities.Message("BOQ Not entered for the line number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        '  Return False
                    End If
                End If

                If oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_8", intRow)) < 0 Then
                    '   oApplication.Utilities.Message("Estimated Cost should be greater than or equal to Zero. Line number:" & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    ' oMatrix.Columns.Item("V_8").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    ' Return False
                End If

                oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(intRow).Specific
                If oCheckBox.Checked = True Then
                    If oApplication.Utilities.getMatrixValues(oMatrix, "V_6", intRow) = "" Then
                        'oApplication.Utilities.Message("Sales Order number is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Return False
                    End If
                End If


                'For intLoop As Integer = intRow + 1 To oMatrix.RowCount
                '    strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intLoop)
                '    strActivity1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intLoop)
                '    If strcode1 <> "" Then
                '        If strCode.ToUpper = strcode1.ToUpper And strActivity.ToUpper = strActivity1.ToUpper Then
                '            'oApplication.Utilities.Message("Process and Activity details already exists : " & strCode & "-" & strActivity, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '            'Return False
                '        End If
                '    End If
                'Next
                Dim strdays, strdate As String
                strdays = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", intRow)
                dblLineBudget = dblLineBudget + oApplication.Utilities.getDocumentQuantity(strdays)


                strdate = oApplication.Utilities.getMatrixValues(oMatrix, "CmpDate", intRow)
                oCombobox = oMatrix.Columns.Item("V_12").Cells.Item(intRow).Specific
                If oCombobox.Selected.Value = "C" Then
                    If strdate = "" Then
                        oApplication.Utilities.Message("Completion date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("CmpDate").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Return False
                    End If
                End If

                Dim dtFrom, dtTo, dtDate As Date
                strdate = oApplication.Utilities.getEdittextvalue(aform, "19")
                If strdate = "" Then
                    oApplication.Utilities.Message("Project Budget From Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Items.Item("19").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Else
                    dtFrom = oApplication.Utilities.GetDateTimeValue(strdate)
                End If
                strdate = oApplication.Utilities.getEdittextvalue(aform, "21")
                If strdate = "" Then
                    oApplication.Utilities.Message("Project Budget End Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Else
                    dtTo = oApplication.Utilities.GetDateTimeValue(strdate)
                End If

                If dtFrom > dtTo Then
                    oApplication.Utilities.Message("Project Budget End date should be Greater than or equal to From date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
                strdate = oApplication.Utilities.getMatrixValues(oMatrix, "V_9", intRow)
                If strdate = "" Then
                    oApplication.Utilities.Message("From date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_9").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                Else
                    dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                    If dtDate >= dtFrom And dtDate <= dtTo Then
                    Else
                        oApplication.Utilities.Message("From date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_9").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Return False
                    End If

                End If

                strdate = oApplication.Utilities.getMatrixValues(oMatrix, "V_10", intRow)
                If strdate = "" Then
                    oApplication.Utilities.Message("End date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_10").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    Return False
                Else
                    dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                    If dtDate >= dtFrom And dtDate <= dtTo Then
                    Else
                        oApplication.Utilities.Message("End date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_10").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Return False
                    End If

                End If
            Else
                oMatrix.DeleteRow(intRow)
            End If
        Next

        oMatrix = aform.Items.Item("12").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Process details are missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        dblBudget = CDbl(oApplication.Utilities.getEdittextvalue(aform, "8"))
        If dblBudget <> dblLineBudget Then
            If oApplication.SBO_Application.MessageBox("Total man days does not match with Line man days. Do you want to save this document ? ", , "Continue", "Cancel") = 2 Then
                Return False
            Else


            End If
        End If
        
        Return True
    End Function
#End Region
#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Budget
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("6").Enabled = False
                    End If
                Case mnu_ADD_ROW


                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        AddRow(oForm)
                    End If
                Case mnu_Duplicate_Row
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        DuplicateRow(oForm)
                        BubbleEvent = False
                        Exit Sub
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    Else
                        If ValidateDeletion(oForm) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = False
                        'oForm.Items.Item("8").Enabled = True
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = True
                        oForm.Items.Item("8").Enabled = True
                    End If
                Case mnu_Delete
                    If pVal.BeforeAction = True Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm
                        If oForm.TypeEx = frm_Budget Then
                            If oApplication.SBO_Application.MessageBox("Do you want to remove this project details?", , "Continue", "Cancel") = 1 Then
                                Dim objEditText As SAPbouiCOM.EditText
                                Dim strDocNum As String
                                oCombobox = oForm.Items.Item("4").Specific
                                strDocNum = oCombobox.Selected.Value
                                Dim orec As SAPbobsCOM.Recordset
                                orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orec.DoQuery("select * from [@Z_OEXP] T0 Inner Join [@Z_EXP1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strDocNum & "'")
                                If orec.RecordCount > 0 Then
                                    oApplication.Utilities.Message("You can not remove this project details . Expenses already exists for this project.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                orec.DoQuery("select * from [@Z_OTIM T0 Inner Join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strDocNum & "'")
                                If orec.RecordCount > 0 Then
                                    oApplication.Utilities.Message("You can not remove this project details . Time Sheet already exists for this project.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Else
                                BubbleEvent = False
                                Exit Sub
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                Dim STRcODE As String
                ' ProjectDetailstoSAP("", "Add")
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                Dim strProjectCode As String
                ' strProjectCode = oApplication.Utilities.getEdittextvalue(oForm, "4")
                'If strProjectCode <> "" Then
                '    ' ProjectDetailstoSAP(strProjectCode, "Update")
                'End If
            End If
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
        AssignLineNo(aForm)
        Return True
    End Function


#Region "Update / Add Project Details in SAP"
    Private Function ProjectDetailstoSAP(ByVal strProjectCode As String, ByVal strChoice As String) As Boolean
        Dim oTempRS As SAPbobsCOM.Recordset
        Dim intDocEntry As Integer
        Dim strProject, strProjectName As String
        oApplication.Utilities.ExecuteSQL(oTempRS, "Select * from [@Z_HPRJ] where U_Z_PRJCODE='" & strProjectCode & "'")
        intDocEntry = 0
        strProject = ""
        strProjectName = ""
        If oTempRS.RecordCount > 0 Then
            intDocEntry = oTempRS.Fields.Item("DocEntry").Value
            strProject = oTempRS.Fields.Item("U_Z_PRJCODE").Value
            strProjectName = oTempRS.Fields.Item("U_Z_PRJNAME").Value
        End If
        If strChoice <> "Delete" Then
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery("Select * from OPRJ where PrjCode='" & strProject & "'")
            If oTempRS.RecordCount > 0 Then
                strChoice = "Update"
            Else
                strChoice = "Add"
            End If

        End If
        If strChoice = "Add" Then
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            oCmpSrv = oApplication.Company.GetCompanyService
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery("Select * from [@Z_HPRJ] order by DocEntry Desc")
            intDocEntry = oTempRS.Fields.Item("DocEntry").Value
            strProject = oTempRS.Fields.Item("U_Z_PRJCODE").Value
            strProjectName = oTempRS.Fields.Item("U_Z_PRJNAME").Value
            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
            project = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)
            project.Code = strProject
            project.Name = strProjectName
            projectService.AddProject(project)
        ElseIf strChoice = "Update" Then
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            Dim projectParams As SAPbobsCOM.IProjectParams
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery("Select * from [@Z_HPRJ] where U_Z_PrjCode='" & strProjectCode & "'")

            oCmpSrv = oApplication.Company.GetCompanyService
            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
            'Get a project
            projectParams = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)
            projectParams.Code = strProject
            project = projectService.GetProject(projectParams)
            'Update the project
            project.Name = strProjectName
            projectService.UpdateProject(project)
        ElseIf strChoice = "Delete" Then
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            Dim projectParams As SAPbobsCOM.IProjectParams
            oCmpSrv = oApplication.Company.GetCompanyService
            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
            'Get a project
            projectParams = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)
            projectParams.Code = strProject
            project = projectService.GetProject(projectParams)
            'delete the project
            Try
                projectService.DeleteProject(projectParams)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
        End If
        Return True

    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Budget Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.ColUID = "V_6" Then
                                    oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                                    If oCheckBox.Checked = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "12" And pVal.ColUID = "BOQ" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oCombobox = oForm.Items.Item("17").Specific
                                If oCombobox.Selected.Value = "C" Or oCombobox.Selected.Value = "H" Then
                                    If pVal.ItemUID = "12" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "12" And (pVal.ColUID = "Qty" Or pVal.ColUID = "Measure") And pVal.CharPressed <> 9 Then
                                    oCombobox = oMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                    If oCombobox.Selected.Value = "R" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "12"
                                    frmSourceMatrix = oMatrix
                                End If

                                If pVal.ItemUID = "12" And pVal.ColUID = "V_6" Then
                                    oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                                    If oCheckBox.Checked = False Then
                                        BubbleEvent = False
                                        Exit Sub

                                    End If
                                End If
                                ' oMatrix = oForm.Items.Item("13").Specific
                                'If pVal.ItemUID = "13" And pVal.Row > 0 And pVal.Row <= oMatrix.RowCount Then
                                '    Me.RowtoDelete = pVal.Row
                                '    Me.MatrixId = "13"
                                '    intSelectedMatrixrow = pVal.Row
                                '    frmSourceMatrix = oMatrix
                                'End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "4" Then
                                    oCombobox = oForm.Items.Item("4").Specific
                                    oApplication.Utilities.setEdittextvalue(oForm, "6", oCombobox.Selected.Description)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" And pVal.ColUID <> "V_-1" Then
                                    If 1 = 1 Then
                                        Dim objChooseForm As SAPbouiCOM.Form
                                        Dim objChoose As New ClsBOQ
                                        Dim strproject, strprojectname, strbusinessprocess, strActvity, strRef As String
                                        oMatrix = oForm.Items.Item("12").Specific
                                        oCombobox = oForm.Items.Item("4").Specific
                                        strproject = oCombobox.Selected.Value ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                        strprojectname = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                        strbusinessprocess = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                        strActvity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", pVal.Row)
                                        strRef = oApplication.Utilities.getMatrixValues(oMatrix, "BOQ", pVal.Row)
                                        If strproject = "" Then
                                            Exit Sub
                                        End If
                                        oCombobox = oMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                        If oCombobox.Selected.Value = "R" Then
                                            Exit Sub
                                        End If
                                        If strproject <> "" Then
                                            objChoose.ItemUID = pVal.ItemUID
                                            objChoose.SourceFormUID = FormUID
                                            objChoose.SourceLabel = pVal.Row
                                            objChoose.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                            objChoose.choice = "MODULE"
                                            objChoose.prjcode = strproject
                                            objChoose.prjname = strprojectname
                                            oCombobox = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                                            objChoose.stats = oCombobox.Selected.Description
                                            objChoose.boqref = strRef
                                            objChoose.businessprocess = strbusinessprocess
                                            objChoose.busienssactivity = strActvity
                                            objChoose.sourcerowId = pVal.Row
                                            objChoose.BinDescrUID = ""
                                            oApplication.Utilities.LoadForm("frm_BOQ.xml", "frm_BOQ")
                                            objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                            objChoose.databound(objChooseForm)
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "13"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "btnDuplica"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_Duplicate_Row)
                                    Case "10"
                                        oForm.PaneLevel = 1
                                    Case "12"
                                        oMatrix = oForm.Items.Item("12").Specific
                                        If pVal.ColUID = "V_5" Then
                                            oCheckBox = oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific
                                            If oCheckBox.Checked = False Then
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", pVal.Row, "")
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", pVal.Row, "")
                                            End If
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" And pVal.ColUID = "V_2" And pVal.CharPressed = 9 Then
                                    oMatrix = oForm.Items.Item("12").Specific
                                    Dim strHours As String
                                    Dim dblHours, dblRate As Double
                                    strHours = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", pVal.Row)
                                    If strHours <> "" Then
                                        dblHours = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_2", pVal.Row))
                                        If (dblHours > 0) Then
                                            dblHours = dblHours * 8
                                            Dim oTest As SAPbobsCOM.Recordset
                                            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oEditText = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                            oTest.DoQuery("Select isnull(U_Daily_rate,0) from OHEM where empID=" & oEditText.String)
                                            dblRate = oTest.Fields.Item(0).Value
                                            dblRate = dblRate * oApplication.Utilities.getDocumentQuantity(strHours)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, dblHours)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", pVal.Row, dblRate)

                                        End If
                                    End If
                                End If

                                If pVal.ItemUID = "12" And pVal.ColUID = "BOQ" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New ClsBOQ
                                    Dim strproject, strprojectname, strbusinessprocess, strActvity, strRef As String
                                    oMatrix = oForm.Items.Item("12").Specific
                                    oCombobox = oForm.Items.Item("4").Specific
                                    strproject = oCombobox.Selected.Value ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    strprojectname = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                    strbusinessprocess = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                    strActvity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", pVal.Row)
                                    strRef = oApplication.Utilities.getMatrixValues(oMatrix, "BOQ", pVal.Row)
                                    If strproject = "" Then
                                        Exit Sub
                                    End If

                                    oCombobox = oMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                    If oCombobox.Selected.Value = "R" Then
                                        Exit Sub
                                    End If
                                    If strproject <> "" Then
                                        objChoose.ItemUID = pVal.ItemUID
                                        objChoose.SourceFormUID = FormUID
                                        objChoose.SourceLabel = pVal.Row
                                        objChoose.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        objChoose.choice = "MODULE"
                                        objChoose.prjcode = strproject
                                        objChoose.prjname = strprojectname
                                        oCombobox = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                                        objChoose.stats = oCombobox.Selected.Description
                                        objChoose.boqref = strRef
                                        objChoose.businessprocess = strbusinessprocess
                                        objChoose.busienssactivity = strActvity
                                        objChoose.sourcerowId = pVal.Row
                                        objChoose.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("frm_BOQ.xml", "frm_BOQ")
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
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
                                            val = oDataTable.GetValue("U_Z_MODNAME", 0)
                                            oMatrix = oForm.Items.Item("12").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val)
                                        End If
                                        If pVal.ItemUID = "12" And pVal.ColUID = "V_3" Then
                                            val = oDataTable.GetValue("U_Z_ACTNAME", 0)
                                            oMatrix = oForm.Items.Item("12").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", pVal.Row, val)
                                        End If

                                        If pVal.ItemUID = "12" And pVal.ColUID = "V_4" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            val1 = oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0)
                                            oMatrix = oForm.Items.Item("12").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_11", pVal.Row, val1)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", pVal.Row, val)
                                        End If
                                        If pVal.ItemUID = "12" And pVal.ColUID = "V_6" Then
                                            val = oDataTable.GetValue("DocEntry", 0)
                                            val1 = oDataTable.GetValue("DocNum", 0)
                                            oMatrix = oForm.Items.Item("12").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", pVal.Row, val1)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", pVal.Row, val)
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

