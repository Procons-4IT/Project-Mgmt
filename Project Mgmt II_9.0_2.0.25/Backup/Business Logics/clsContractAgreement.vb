Imports System.IO
Public Class clsContractAgreement

    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oColumn As SAPbouiCOM.Column
    Private oCombocolumn As SAPbouiCOM.ComboBoxColumn
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

    Dim strFileName As String
    Dim strSelectedFilepath, strSelectedFolderPath, strSelectedFileName As String
#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_Contract, frm_Contract)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "19"
        AddChooseFromList(oForm)
        oCombobox = oForm.Items.Item("57").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)

        Next
        Try
            oCombobox.ValidValues.Add("C", "Customer")
            oCombobox.ValidValues.Add("S", "Supplier")
        Catch ex As Exception

        End Try

        'oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        ' oForm.EnableMenu(mnu_Duplicate_Row, True)
        oRecset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Select * from NNM1 where ""ObjectCode""='Z_OPAT'")
        oCombobox = oForm.Items.Item("18").Specific
        For intRow As Integer = 0 To oRecset.RecordCount - 1
            oCombobox.ValidValues.Add(oRecset.Fields.Item("Series").Value, oRecset.Fields.Item("seriesname").Value)
            oRecset.MoveNext()
        Next
        oForm.Items.Item("18").DisplayDesc = True
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oRecset.DoQuery("Select * from ONNM where ""ObjectCode""='Z_OPAT'")
        Dim intdef As String = oRecset.Fields.Item("DfltSeries").Value
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PAT1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix = oForm.Items.Item("42").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oForm.Items.Item("29").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        LoadGridValues(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        ' databind(oForm)
        oForm.Freeze(False)
    End Sub

    Public Sub ViewForm(ByVal aCode As String)
        oForm = oApplication.Utilities.LoadForm(xml_Contract, frm_Contract)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "19"
        AddChooseFromList(oForm)
        'oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        ' oForm.EnableMenu(mnu_Duplicate_Row, True)
        oRecset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Select * from NNM1 where ""ObjectCode""='Z_OPAT'")
        oCombobox = oForm.Items.Item("18").Specific
        For intRow As Integer = 0 To oRecset.RecordCount - 1
            oCombobox.ValidValues.Add(oRecset.Fields.Item("Series").Value, oRecset.Fields.Item("seriesname").Value)
            oRecset.MoveNext()
        Next
        oForm.Items.Item("18").DisplayDesc = True
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oRecset.DoQuery("Select * from ONNM where ""ObjectCode""='Z_OPAT'")
        Dim intdef As String = oRecset.Fields.Item("DfltSeries").Value
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PAT1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix = oForm.Items.Item("42").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oForm.Items.Item("29").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        LoadGridValues(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Items.Item("19").Enabled = True
        oApplication.Utilities.setEdittextvalue(oForm, "19", aCode)
        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        ' databind(oForm)
        oForm.Freeze(False)
    End Sub

#Region "ShowFileDialog"

    '*****************************************************************
    'Type               : Procedure
    'Name               : ShowFileDialog
    'Parameter          :
    'Return Value       :
    'Author             : Senthil Kumar B 
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To open a File Browser
    '******************************************************************

    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()

    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog
        Dim strFileName, strMdbFilePath As String
        Dim oEdit As SAPbouiCOM.EditText
        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.SafeFileName
                        strSelectedFilepath = oDialogBox.FileName
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                        strSelectedFileName = strMdbFilePath
                        If strSelectedFolderPath.EndsWith("\") Then
                            strSelectedFolderPath = strSelectedFilepath.Substring(0, strSelectedFolderPath.Length - 1)
                        End If
                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region

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
            Exit Sub

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
            oMatrix = aForm.Items.Item("42").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PAT1")
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
            oMatrix = aForm.Items.Item("42").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PAT1")
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If
            oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
            Try
                If oEditText.Value <> "" Then
                    oMatrix.AddRow()
                    oMatrix.ClearRowData(oMatrix.RowCount)
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

    Private Sub LoadFiles(ByVal aform As SAPbouiCOM.Form)
        oMatrix = aform.Items.Item("42").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(intRow) Then
                Dim strFilename, strFilePath As String
                strFilename = oMatrix.Columns.Item("V_2").Cells.Item(intRow).Specific.value
                strFilePath = oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific.value

                If File.Exists(strFilePath) = False Then
                    Dim oRec As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim strQry = "Select AttachPath From OADP"
                    oRec.DoQuery(strQry)
                    strFilePath = oRec.Fields.Item(0).Value

                    If strFilename = "" Then
                        strFilePath = strFilePath
                    Else
                        strFilePath = strFilePath & strFilename
                    End If
                    If File.Exists(strFilePath) = False Then
                        oApplication.Utilities.Message("File does not exists ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    strFilename = strFilePath
                Else
                    strFilename = strFilePath
                End If

                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No file has been selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    End Sub

    Private Function ValidateDeletion(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Return True
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

    Private Sub Addmode(ByVal aform As SAPbouiCOM.Form)
        oForm = aform
        oForm.Freeze(True)
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from ONNM where ObjectCode='Z_OPAT'")
        Dim intDefSeries As Integer = oRec.Fields.Item("DfltSeries").Value
        oRec.DoQuery("Select series,seriesname,* from NNM1 where ObjectCode='Z_OPAT'and series=" & intDefSeries)
        oCombobox = aform.Items.Item("18").Specific
        oCombobox.Select(intDefSeries.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
        If oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "19")) = 0 Then
            oApplication.Utilities.setEdittextvalue(aform, "19", oRec.Fields.Item("NextNumber").Value)
        End If
        oApplication.Utilities.setEdittextvalue(oForm, "21", "t")
        oApplication.SBO_Application.SendKeys("{TAB}")
        aform.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Freeze(False)
    End Sub

    Private Sub NextNumber(ByVal aform As SAPbouiCOM.Form)
        oForm = aform
        oForm.Freeze(True)
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from ONNM where ObjectCode='Z_OPAT'")
        oCombobox = aform.Items.Item("18").Specific
        Dim intDefSeries As Integer = oCombobox.Selected.Value
        oRec.DoQuery("Select series,seriesname,* from NNM1 where ObjectCode='Z_OPAT'and series=" & intDefSeries)
        oCombobox = aform.Items.Item("18").Specific
        oApplication.Utilities.setEdittextvalue(aform, "19", oRec.Fields.Item("NextNumber").Value)
        oApplication.Utilities.setEdittextvalue(oForm, "21", "t")
        oApplication.SBO_Application.SendKeys("{TAB}")
        oForm.Freeze(False)
    End Sub

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        'If Me.MatrixId = "12" Then
        '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        '    'Else
        '    '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        'End If
        oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_PAT1")
        If intSelectedMatrixrow <= 0 Then
            Exit Sub
        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oMatrix = aForm.Items.Item("42").Specific
        oMatrix.DeleteRow(intSelectedMatrixrow)

        'oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)


        'oMatrix.FlushToDataSource()

        'For count = 1 To oDataSrc_Line.Size - 1
        '    oDataSrc_Line.SetValue("LineId", count - 1, count)
        'Next
        'oMatrix.LoadFromDataSource()
        'If oMatrix.RowCount > 0 Then
        '    oMatrix.DeleteRow(oMatrix.RowCount)
        'End If
        AssignLineNo(aForm)
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
        End If
    End Sub
#End Region
#End Region

#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject, strActivity, strActivity1 As String
        Dim oTemp As SAPbobsCOM.Recordset

        strProject = oApplication.Utilities.getEdittextvalue(aform, "7")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If strProject = "" Then
                oApplication.Utilities.Message("Business Partner Details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Dim dblBudget, dblLineBudget As Double
        Dim strHours As String
        Dim dblHours, dblRate As Double
        dblBudget = 0
        dblLineBudget = 0
        dblRate = 0
        Dim strdays, strdate As String
        Dim dtFrom, dtTo, dtDate As Date
        strdate = oApplication.Utilities.getEdittextvalue(aform, "21")
        If strdate = "" Then
            oApplication.Utilities.Message("Document Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        Else
            dtFrom = oApplication.Utilities.GetDateTimeValue(strdate)
        End If

        strdate = oApplication.Utilities.getEdittextvalue(aform, "23")
        If strdate = "" Then
            oApplication.Utilities.Message("Start Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Items.Item("23").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        Else
            dtFrom = oApplication.Utilities.GetDateTimeValue(strdate)
        End If

        strdate = oApplication.Utilities.getEdittextvalue(aform, "25")
        If strdate = "" Then
            oApplication.Utilities.Message("End Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        Else
            dtTo = oApplication.Utilities.GetDateTimeValue(strdate)
        End If

        If dtFrom > dtTo Then
            oApplication.Utilities.Message("End date should be Greater than or equal to Start date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
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
                Case mnu_Contract
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("19").Enabled = False
                        oForm.Items.Item("18").Enabled = False
                        oForm.Items.Item("7").Enabled = False
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
                        oForm.Items.Item("19").Enabled = False
                        oForm.Items.Item("18").Enabled = True
                        oForm.Items.Item("7").Enabled = True
                        Addmode(oForm)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("19").Enabled = True
                        oForm.Items.Item("18").Enabled = True
                        oForm.Items.Item("7").Enabled = True
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                Dim STRcODE As String
                LoadGridValues(oForm)
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                Dim strProjectCode As String
               LoadGridValues(oForm)
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

        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strQry = "Select AttachPath From OADP"
        oRec.DoQuery(strQry)
        oMatrix = aForm.Items.Item("42").Specific
        Dim SPath As String = oRec.Fields.Item(0).Value
        For intRow As Integer = 1 To oMatrix.RowCount

            If SPath = "" Then
            Else
                Dim DPath As String = ""
                SPath = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                If Not oRec.EoF Then
                    DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                End If
                If Not Directory.Exists(DPath) Then
                    Directory.CreateDirectory(DPath)
                End If
                Dim file = New FileInfo(SPath)
                Dim Filename As String = Path.GetFileName(SPath)
                Dim SavePath As String = Path.Combine(DPath, Filename)
                If System.IO.File.Exists(SavePath) Then
                Else
                    file.CopyTo(Path.Combine(DPath, file.Name), True)
                End If
            End If
        Next
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
            If pVal.FormTypeEx = frm_Contract Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                                'If pVal.ItemUID = "47" And pVal.ColUID = "DocEntry" Then
                                '    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                '    oEditTextColumn = oGrid.Columns.Item("DocEntry")
                                '    oEditTextColumn.LinkedObjectType = oGrid.DataTable.GetValue("ObjType", pVal.Row)
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("42").Specific
                                If pVal.ItemUID = "42" And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "42"
                                    frmSourceMatrix = oMatrix
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "47" And pVal.ColUID = "DocEntry" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oEditTextColumn = oGrid.Columns.Item("DocEntry")
                                    Dim intRow As Integer = oGrid.Rows.GetParent(pVal.Row)
                                    oEditTextColumn.LinkedObjectType = oGrid.DataTable.GetValue("ObjType", intRow)
                                End If

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "18" Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        NextNumber(oForm)
                                    End If
                                End If

                                If pVal.ItemUID = "12" Then
                                    oCombobox = oForm.Items.Item("12").Specific
                                    PopulateContactPersons_Details(oForm, oCombobox.Selected.Value)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "60"
                                        Dim strCntid As String
                                        Dim oobj As New clsBudgetDetails_Contract
                                        strCntid = oApplication.Utilities.getEdittextvalue(oForm, "19")
                                        oobj.LoadForm(strCntid)
                                    Case "13"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "btnDuplica"
                                        '   oApplication.SBO_Application.ActivateMenuItem(mnu_Duplicate_Row)
                                    Case "29"
                                        oForm.PaneLevel = 1
                                    Case "30"
                                        oForm.PaneLevel = 2
                                    Case "49"
                                        oForm.PaneLevel = 4
                                    Case "50"
                                        oForm.PaneLevel = 5
                                    Case "51"
                                        oForm.PaneLevel = 6
                                    Case "52"
                                        oForm.PaneLevel = 7
                                    Case "32"
                                        oForm.PaneLevel = 3
                                    Case "43"
                                        fillopen()
                                        oMatrix = oForm.Items.Item("42").Specific
                                        AddRow(oForm)
                                        Try
                                            oForm.Freeze(True)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, strSelectedFilepath)
                                            Dim strDate As String
                                            Dim dtdate As Date
                                            dtdate = Now.Date
                                            strDate = Now.Date.Today().ToString
                                            Dim oColumn As SAPbouiCOM.Column
                                            oColumn = oMatrix.Columns.Item("V_1")
                                            oColumn.Editable = True
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, strSelectedFileName)
                                            oEditText = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                                            oEditText.String = "t"
                                            oApplication.SBO_Application.SendKeys("{TAB}")
                                            oForm.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oColumn.Editable = False
                                            AssignLineNo(oForm)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            oForm.Freeze(False)
                                        Catch ex As Exception
                                            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oForm.Freeze(False)

                                        End Try
                                    Case "44"
                                        LoadFiles(oForm)
                                    Case "45"
                                        oMatrix = oForm.Items.Item("42").Specific
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                End Select
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
                                        If pVal.ItemUID = "7" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            PopulateContactPersons(oForm, val)
                                            oApplication.Utilities.setEdittextvalue(oForm, "10", oDataTable.GetValue("CardName", 0))
                                            oApplication.Utilities.setEdittextvalue(oForm, "7", val)
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


    Private Sub LoadDocuments(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim strProject, strQuery As String
            Try
                strProject = oApplication.Utilities.getEdittextvalue(aForm, "19")
            Catch ex As Exception
                strProject = ""
            End Try
            Dim oTest, otest1 As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strProject <> "" Then
                otest1.DoQuery("Select * from ""@Z_OPAT"" where ""DocNum""=" & strProject)
                If otest1.RecordCount > 0 Then
                    strProject = otest1.Fields.Item("DocNum").Value
                Else
                    strProject = ""
                End If
            Else
                strProject = "99999"
            End If
            strQuery = strQuery & " select  'Purchase Reqeuest' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription,T0.Quantity 'Quantity'  from PRQ1 T0 inner Join OPRQ T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Purchase Order' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'GRPO' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from PDN1 T0 inner Join OPDN T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Purchase Return' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from RPD1 T0 inner Join ORPD T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Purchase DownPayment Request' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from DPO1 T0 inner Join ODPO T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ") Union all"
            strQuery = strQuery & " select 'Purchase Invoice' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription,T0.Quantity from PCH1 T0 inner Join OPCH T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ") Union all"
            strQuery = strQuery & " select 'Purchase Credit Note' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry', T1.DocDate,T0.LineTotal,T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription,T0.Quantity  from RPC1  T0 inner Join ORPC T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"

            strQuery = strQuery & " select 'Sales Quatation' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from QUT1  T0 inner Join OQUT  T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Sales Order' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription,T0.Quantity  from RDR1 T0 inner Join ORDR T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Delivery' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry 'DocEntry' ,T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from DLN1 T0 inner Join ODLN T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Sales Return' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry  'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from RDN1 T0 inner Join ORDN  T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Sales Invoice' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry   'DocEntry',T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription ,T0.Quantity from RIN1 T0 inner Join OINV T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Salres credit Note' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry   'DocEntry', T1.DocDate,T0.LineTotal,T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription,T0.Quantity  from INV1 T0 inner Join ORIN T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")  Union all"
            strQuery = strQuery & " select 'Sales Downpament' 'TransType',t0.ObjType ,T1.DocNum,T1.DocEntry 'DocEntry' ,T1.DocDate,T0.LineTotal, T0.Project,T0.U_Z_MDNAME,T0.U_Z_ACTNAME,T0.U_Z_BOQREF ,T0.ItemCode, T0.Dscription,T0.Quantity  from DPI1 T0 inner Join ODPI  T1 on T1.DocEntry=T0.DocEntry  where (T0.U_Z_CUSTCNTID='" & strProject & "' or U_Z_CntID=" & strProject & ")"
            oGrid = aForm.Items.Item("47").Specific
            strQuery = "select X.ObjType,x.TransType,X.DocNum,X.DocEntry,X.DocDate,X.Project,X.U_Z_MDNAME,X.U_Z_ACTNAME,X.U_Z_BOQREF,X.ItemCode,X.Dscription,X.Quantity,X.LineTotal from ( " & strQuery & " ) X order by Convert(numeric,X.ObjType)"


            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item("TransType").TitleObject.Caption = "Document Type"
            oGrid.Columns.Item("ObjType").TitleObject.Caption = "Transaction Type"
            oGrid.Columns.Item("ObjType").Visible = False

            oGrid.Columns.Item("DocNum").TitleObject.Caption = "Document Number"
            oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Document Entry"
            oGrid.Columns.Item("DocDate").TitleObject.Caption = "Document Date"
            'oEditTextColumn = oGrid.Columns.Item("TransType")
            'oEditTextColumn.LinkedObjectType = "17"
            'oEditTextColumn = oGrid.Columns.Item("DocEntry")
            'oEditTextColumn.LinkedObjectType = "17"
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("Project").TitleObject.Caption = "Project Code"
            oEditTextColumn = oGrid.Columns.Item("Project")
            oEditTextColumn.LinkedObjectType = "63"
            oGrid.Columns.Item("U_Z_MDNAME").TitleObject.Caption = "Project Phase"
            oGrid.Columns.Item("U_Z_ACTNAME").TitleObject.Caption = "Project Activity"
            oGrid.Columns.Item("U_Z_BOQREF").TitleObject.Caption = "BoQ Reference"
            oGrid.Columns.Item("U_Z_BOQREF").Visible = False
            oGrid.Columns.Item("ItemCode").TitleObject.Caption = "Item Code"
            oEditTextColumn = oGrid.Columns.Item("ItemCode")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("ItemCode").Visible = False
            oGrid.Columns.Item("Dscription").TitleObject.Caption = "Item Name"
            oGrid.Columns.Item("Dscription").Visible = False
            oGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("Quantity").Visible = False
            oGrid.Columns.Item("LineTotal").TitleObject.Caption = "Amount"

            'oEditTextColumn = oGrid.Columns.Item("DocEntry")
            'oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes
            oGrid.CollapseLevel = 2
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)

        End Try
    End Sub
    Private Sub LoadGridValues(ByVal aForm As SAPbouiCOM.Form)
        Dim strQuery, strProject As String
        Try

            aForm.Freeze(True)

            Try
                strProject = oApplication.Utilities.getEdittextvalue(aForm, "19")
            Catch ex As Exception
                strProject = ""
            End Try
            Dim oTest, otest1 As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strProject <> "" Then


                otest1.DoQuery("Select * from ""@Z_OPAT"" where ""DocNum""=" & strProject)
                If otest1.RecordCount > 0 Then
                    strProject = otest1.Fields.Item("DocNum").Value
                Else
                    strProject = ""
                End If
            End If
            If strProject <> "" Then
                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ ,T1.U_Z_EXPTYPE from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='R'  and (T1.U_Z_CNTID='" & strProject & "' or T1.U_Z_CUSTCNTID='" & strProject & "')"
                oGrid = aForm.Items.Item("grdRes").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "R")
                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='I' and (T1.U_Z_CNTID='" & strProject & "' or T1.U_Z_CUSTCNTID='" & strProject & "')"
                oGrid = aForm.Items.Item("grdItem").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "I")
                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_EXPTYPE,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='E' and (T1.U_Z_CNTID='" & strProject & "' or T1.U_Z_CUSTCNTID='" & strProject & "')"
                oGrid = aForm.Items.Item("grdExp").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  (T1.U_Z_CNTID='" & strProject & "' or T1.U_Z_CUSTCNTID='" & strProject & "')"
                oGrid = aForm.Items.Item("grdSum").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                ' FormatGrid(aForm, oGrid, "E")

            Else
                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_CNTID ='x'"
                oGrid = aForm.Items.Item("grdRes").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "R")
                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_CNTID='x'"
                oGrid = aForm.Items.Item("grdItem").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "I")
                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_EXPTYPE,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ ,T1.U_Z_EXPTYPE from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_CNTID='x'"
                oGrid = aForm.Items.Item("grdExp").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

                strQuery = "select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CUSTCNTID,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_CNTID='x'"
                oGrid = aForm.Items.Item("grdSum").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

            End If
            LoadDocuments(aForm)
            oGrid = aForm.Items.Item("grdRes").Specific
            FormatGrid(aForm, oGrid, "R")
            oGrid = aForm.Items.Item("grdItem").Specific
            FormatGrid(aForm, oGrid, "I")
            oGrid = aForm.Items.Item("grdExp").Specific
            FormatGrid(aForm, oGrid, "E")
            aForm.Freeze(False)
            oGrid = aForm.Items.Item("grdSum").Specific
            FormatGrid(aForm, oGrid, "A")
            '   aForm.Items.Item("100").Enabled = False
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub FormatGrid(ByVal aform As SAPbouiCOM.Form, ByVal aGrid As SAPbouiCOM.Grid, ByVal aChoice As String)
        Try
            aform.Freeze(True)
            aGrid.Columns.Item("DocEntry").Visible = False
            aGrid.Columns.Item("U_Z_PRJCODE").TitleObject.Caption = "Project Code"
            aGrid.Columns.Item("U_Z_PRJNAME").TitleObject.Caption = "Project Name"
            aGrid.Columns.Item("U_Z_MODNAME").TitleObject.Caption = "Phase"
            aGrid.Columns.Item("U_Z_ACTNAME").TitleObject.Caption = "Activity "
            aGrid.Columns.Item("U_Z_TYPE").TitleObject.Caption = "Type"
            aGrid.Columns.Item("U_Z_TYPE").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oCombocolumn = aGrid.Columns.Item("U_Z_TYPE")
            oCombocolumn.ValidValues.Add("E", "Expenses")
            oCombocolumn.ValidValues.Add("R", "Resource")
            oCombocolumn.ValidValues.Add("I", "Item")
            oCombocolumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            If aChoice = "A" Then
                aGrid.Columns.Item("U_Z_TYPE").Visible = True
                aGrid.Columns.Item("LineId").Visible = True
            Else
                aGrid.Columns.Item("U_Z_TYPE").Visible = False
                aGrid.Columns.Item("LineId").Visible = False
            End If
            aGrid.Columns.Item("U_Z_EMPID").TitleObject.Caption = "Employee ID"
            aGrid.Columns.Item("U_Z_POSITION").TitleObject.Caption = "Employee Name"
            aGrid.Columns.Item("U_Z_POSITION").Editable = False
            aGrid.Columns.Item("U_Z_FROMDATE").TitleObject.Caption = "Start Date"
            aGrid.Columns.Item("U_Z_TODATE").TitleObject.Caption = "End Date"
            aGrid.Columns.Item("U_Z_QUANTITY").TitleObject.Caption = "Quantity"
            aGrid.Columns.Item("U_Z_MEASURE").TitleObject.Caption = "Measure"
            If aChoice = "R" Then
                aGrid.Columns.Item("U_Z_QUANTITY").Visible = False
                aGrid.Columns.Item("U_Z_MEASURE").Visible = False
            Else
                aGrid.Columns.Item("U_Z_QUANTITY").Visible = True
                aGrid.Columns.Item("U_Z_MEASURE").Visible = True

            End If
            aGrid.Columns.Item("U_Z_DAYS").TitleObject.Caption = "Est.Man Days"
            aGrid.Columns.Item("U_Z_HOURS").TitleObject.Caption = "Est.Man Hours"
            aGrid.Columns.Item("U_Z_HOURS").Editable = False
            aGrid.Columns.Item("U_Z_AMOUNT").TitleObject.Caption = "Estimated Cost"
            aGrid.Columns.Item("U_Z_ORDER").TitleObject.Caption = "Payment Milestone"
            aGrid.Columns.Item("U_Z_ORDER").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            aGrid.Columns.Item("U_Z_ORDENTRY").TitleObject.Caption = "SO Link"
            aGrid.Columns.Item("U_Z_ORDNUM").TitleObject.Caption = "Sales Order Number"
            aGrid.Columns.Item("U_Z_ORDNUM").Editable = False
            aGrid.Columns.Item("U_Z_HOURS").Editable = True
            If aChoice = "I" Then
                aGrid.Columns.Item("U_Z_DAYS").Visible = False
                aGrid.Columns.Item("U_Z_HOURS").Visible = False
                aGrid.Columns.Item("U_Z_ORDER").Visible = False
                aGrid.Columns.Item("U_Z_ORDENTRY").Visible = False
                aGrid.Columns.Item("U_Z_ORDNUM").Visible = False
            Else
                aGrid.Columns.Item("U_Z_DAYS").Visible = False
                aGrid.Columns.Item("U_Z_HOURS").Visible = True
                aGrid.Columns.Item("U_Z_ORDER").Visible = True
                aGrid.Columns.Item("U_Z_ORDENTRY").Visible = True
                aGrid.Columns.Item("U_Z_ORDNUM").Visible = True
            End If
            aGrid.Columns.Item("U_Z_CMPDATE").TitleObject.Caption = "Completion Date"



            aGrid.Columns.Item("U_Z_STATUS").TitleObject.Caption = "Status"
            aGrid.Columns.Item("U_Z_STATUS").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oCombocolumn = aGrid.Columns.Item("U_Z_STATUS")
            Try

                oCombocolumn.ValidValues.Add("I", "InProcess")
                oCombocolumn.ValidValues.Add("P", "Pending")
                oCombocolumn.ValidValues.Add("C", "Completed")
            Catch ex As Exception

            End Try
            oCombocolumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            aGrid.Columns.Item("U_Z_BOQ").TitleObject.Caption = "Bill of Quantities"
            If aChoice = "R" Then
                aGrid.Columns.Item("U_Z_BOQ").Visible = False
            Else
                aGrid.Columns.Item("U_Z_BOQ").Visible = True
            End If
            If aChoice = "E" Then
                aGrid.Columns.Item("U_Z_DAYS").Visible = False
                aGrid.Columns.Item("U_Z_HOURS").Visible = False
                aGrid.Columns.Item("U_Z_ORDER").Visible = False
                aGrid.Columns.Item("U_Z_ORDENTRY").Visible = False
                aGrid.Columns.Item("U_Z_ORDNUM").Visible = False
                aGrid.Columns.Item("U_Z_CMPDATE").Visible = True
                aGrid.Columns.Item("U_Z_QUANTITY").Visible = False
                aGrid.Columns.Item("U_Z_MEASURE").Visible = False
                aGrid.Columns.Item("U_Z_BOQ").Visible = False
            Else
                aGrid.Columns.Item("U_Z_CMPDATE").Visible = True
            End If
            aGrid.Columns.Item("U_Z_EXPTYPE").TitleObject.Caption = "Expense Type"
            ' oEditTextColumn = aGrid.Columns.Item("U_Z_EXPTYPE")
            ' oEditTextColumn.ChooseFromListUID = "CFL_EXP"
            ' oEditTextColumn.ChooseFromListAlias = "U_Z_EXPNAME"
            If aChoice = "E" Or aChoice = "A" Then
                aGrid.Columns.Item("U_Z_EXPTYPE").Visible = True
            Else
                aGrid.Columns.Item("U_Z_EXPTYPE").Visible = False
            End If
            oEditTextColumn = aGrid.Columns.Item("U_Z_EMPID")
            oEditTextColumn.LinkedObjectType = "171"
            oEditTextColumn = aGrid.Columns.Item("U_Z_ORDENTRY")
            oEditTextColumn.LinkedObjectType = "17"
            aGrid.AutoResizeColumns()

            aGrid.Columns.Item("U_Z_CNTID").TitleObject.Caption = "Sub.Contract ID"
            'oEditTextColumn = aGrid.Columns.Item("U_Z_CNTID")
            'oEditTextColumn.LinkedObjectType = "Z_OPAT"
            aGrid.Columns.Item("U_Z_CUSTCNTID").TitleObject.Caption = "Customer Contract ID"

            aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            'T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOUS, T1.U_Z_AMOUNT, T1.U_Z_ORDNUM ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_BOQ  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='E' and T0.U_Z_PRJCODE='" & strProject & "'"

            aform.Freeze(False)
        Catch ex As Exception
            aform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
    Private Sub PopulateContactPersons(ByVal aform As SAPbouiCOM.Form, ByVal aCode As String)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("SELECT cntctCode,Name,T0.[E_MailL], T0.[Tel1] FROM OCPR T0  INNER JOIN OCRD T1 ON T0.CardCode = T1.CardCode WHERE T1.[CardCode] ='" & aCode & "'")
        oCombobox = aform.Items.Item("12").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 0 To oTest.RecordCount - 1
            oCombobox.ValidValues.Add(oTest.Fields.Item(0).Value, oTest.Fields.Item(1).Value)
            oTest.MoveNext()
        Next
        aform.Items.Item("12").DisplayDesc = True
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub

    Private Sub PopulateContactPersons_Details(ByVal aform As SAPbouiCOM.Form, ByVal aCode As String)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("SELECT cntctCode,Name,T0.[E_MailL], T0.[Tel1] FROM OCPR T0  INNER JOIN OCRD T1 ON T0.CardCode = T1.CardCode WHERE T0.[cntctCode] ='" & aCode & "'")
        If oTest.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(aform, "14", oTest.Fields.Item(3).Value)
            oApplication.Utilities.setEdittextvalue(aform, "16", oTest.Fields.Item(2).Value)
        Else
            oApplication.Utilities.setEdittextvalue(aform, "14", "")
            oApplication.Utilities.setEdittextvalue(aform, "16", "")
        End If
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
    End Sub

#End Region
End Class





