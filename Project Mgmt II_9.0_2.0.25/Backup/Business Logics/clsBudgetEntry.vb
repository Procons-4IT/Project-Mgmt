Imports System.IO
Public Class clsBudgetEntry
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombocolumn As SAPbouiCOM.ComboBoxColumn
    Private oColumn As SAPbouiCOM.Column
    Private oCheckBox As SAPbouiCOM.CheckBox
    Private oCheckBoxCOlumn As SAPbouiCOM.CheckBoxColumn
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
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_BudgetEntry) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_BudgetEntry, frm_BudgetEntry)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        ' Exit Sub
        oForm.Freeze(True)
        oForm.EnableMenu("1283", True)
        oForm.DataBrowser.BrowseBy = "4"
        FillProjectCode(oForm)
        AddChooseFromList(oForm)
        AddChooseFromList_Expenses(oForm)
        AddChooseFromList_Item(oForm)
        AddChooseFromList_Resource(oForm)
        AddChooseFromList_Contracts(oForm)
        LoadGridValues(oForm)

        databind(oForm)
        ' Exit Sub
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu(mnu_Duplicate_Row, True)
        oMatrix = oForm.Items.Item("41").Specific
        oForm.DataSources.UserDataSources.Add("Line", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oColumn = oMatrix.Columns.Item("V_-1")
        oColumn.DataBind.SetBound(True, "", "Line")
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

        oForm.PaneLevel = 1

        'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        'For count = 1 To oDataSrc_Line.Size - 1
        '    oDataSrc_Line.SetValue("LineId", count - 1, count)
        'Next

        'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        'For count = 1 To oDataSrc_Line.Size - 1
        '    oDataSrc_Line.SetValue("LineId", count - 1, count)
        'Next

        oForm.Items.Item("10").Visible = True
        oForm.Items.Item("30").Visible = False
        oForm.Items.Item("31").Visible = False
        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        ' oForm.Items.Item("6").Enabled = False

        oForm.PaneLevel = 1

        oForm.Freeze(False)
    End Sub

#Region "Fill Project Code"
    Private Sub FillProjectCode(ByVal aForm As SAPbouiCOM.Form)
        Dim oTempRec As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCombobox = aForm.Items.Item("401").Specific
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


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_DPRJ"
            oCFLCreationParams.UniqueID = "CFL_BASE"
            oCFL = oCFLs.Add(oCFLCreationParams)


            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

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


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_Expances"
            oCFLCreationParams.UniqueID = "CFL_EXP"
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

    Private Sub AddChooseFromList_Conditions(ByVal objForm As SAPbouiCOM.Form, ByVal aRow As Integer)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)



            If objForm.PaneLevel > 3 Then
                Exit Sub
            End If
            Select Case objForm.PaneLevel
                Case 1
                    oGrid = objForm.Items.Item("26").Specific
                    oCFL = oCFLs.Item("CFLR3")
                Case 2
                    oGrid = objForm.Items.Item("27").Specific
                    oCFL = oCFLs.Item("CFLI3")
                Case 3
                    oGrid = objForm.Items.Item("28").Specific
                    oCFL = oCFLs.Item("CFLE3")


            End Select
            Dim strmod As String = oGrid.DataTable.GetValue("U_Z_MODNAME", aRow)

            oCons = oCFL.GetConditions()
            If oCons.Count > 1 Then
                oCon = oCons.Item(1)
                oCon.Alias = "U_Z_MODNAME"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oGrid.DataTable.GetValue("U_Z_MODNAME", aRow)
                oCFL.SetConditions(oCons)
                ' oCon = oCons.Add()
            Else
                oCon = oCons.Add()
                'oCon.Alias = "U_Z_Status"
                oCon.Alias = "U_Z_MODNAME"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oGrid.DataTable.GetValue("U_Z_MODNAME", aRow)
                oCFL.SetConditions(oCons)
                oCon = oCons.Add()
            End If


            ' oCon = oCons.Add
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AddChooseFromList_Conditions_SalesOrder(ByVal objForm As SAPbouiCOM.Form, ByVal aRow As Integer)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            Dim strCardCode As String
            oCombobox = objForm.Items.Item("401").Specific
            Dim strPrj As String
            Try
                strPrj = oCombobox.Selected.Value
            Catch ex As Exception
                strPrj = oApplication.Utilities.getEdittextvalue(objForm, "4")
            End Try

            '   oCombobox = objForm.Items.Item("4").Specific

            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("Select isnull(U_Z_CARDCODE,'') from OPRJ where prjcode='" & strPrj & "'")
            strCardCode = oRec.Fields.Item(0).Value


            If objForm.PaneLevel > 3 Then
                Exit Sub
            End If
            Select Case objForm.PaneLevel
                Case 1
                    oGrid = objForm.Items.Item("26").Specific
                    oCFL = oCFLs.Item("CFLR4")
                Case 2
                    oGrid = objForm.Items.Item("27").Specific
                    oCFL = oCFLs.Item("CFLI4")
                Case 3
                    oGrid = objForm.Items.Item("28").Specific
                    oCFL = oCFLs.Item("CFLE4")
            End Select
            Dim strmod As String = oGrid.DataTable.GetValue("U_Z_MODNAME", aRow)
            '  Exit Sub
            oCons = oCFL.GetConditions()
            If oCons.Count >= 2 Then
                oCon = oCons.Item(1)
                oCon.Alias = "CardCode"
                If strCardCode = "" Then
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                    oCon.CondVal = "X"
                Else
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = strCardCode
                End If
                oCFL.SetConditions(oCons)
                ' oCon = oCons.Add()
            Else

                oCon = oCons.Add()
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                'oCon.Alias = "U_Z_Status"
                oCon.Alias = "CardCode"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = strCardCode
                oCFL.SetConditions(oCons)
                oCon = oCons.Add()
            End If


            ' oCon = oCons.Add
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AddChooseFromList_Resource(ByVal objForm As SAPbouiCOM.Form)
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
            oCFLCreationParams.UniqueID = "CFLR1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "Z_Module"
            oCFLCreationParams.UniqueID = "CFLR2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            'oCFLCreationParams.ObjectType = "Z_Activity"
            'oCFLCreationParams.UniqueID = "CFLR3"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            ' '' Adding Conditions to CFL2
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Z_Status"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "CFLR4"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCon.BracketCloseNum = 1

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "CardCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "x"
            oCon.BracketCloseNum = 2
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "DocStatus"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "O"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()

            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            'oCon = oCons.Add()
            'oCon.Alias = "CardCode"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            'oCon.CondVal = "O"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "171"
            oCFLCreationParams.UniqueID = "CFLR5"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 1

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_Z_IsProject"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AddChooseFromList_Item(ByVal objForm As SAPbouiCOM.Form)
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
            oCFLCreationParams.UniqueID = "CFLI1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "Z_Module"
            oCFLCreationParams.UniqueID = "CFLI2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            'oCFLCreationParams.ObjectType = "Z_Activity"
            'oCFLCreationParams.UniqueID = "CFLI3"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            ' '' Adding Conditions to CFL2
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Z_Status"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "CFLI4"
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
            oCFLCreationParams.UniqueID = "CFLI5"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()

            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 1

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_Z_IsProject"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AddChooseFromList_Expenses(ByVal objForm As SAPbouiCOM.Form)
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
            oCFLCreationParams.UniqueID = "CFLE1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "Z_Module"
            oCFLCreationParams.UniqueID = "CFLE2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            'oCFLCreationParams.ObjectType = "Z_Activity"
            'oCFLCreationParams.UniqueID = "CFLE3"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            ' '' Adding Conditions to CFL2
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Z_Status"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "CFLE4"
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
            oCFLCreationParams.UniqueID = "CFLE5"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCons = oCFL.GetConditions()
            oCon = oCons.Add
            oCon.BracketOpenNum = 2
            oCon.Alias = "Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 1

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_Z_IsProject"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 2
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub AddChooseFromList_Contracts(ByVal objForm As SAPbouiCOM.Form)
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
            oCFLCreationParams.ObjectType = "Z_OPAT"
            oCFLCreationParams.UniqueID = "CFLC1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_TYPE"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_OPAT"
            oCFLCreationParams.UniqueID = "CFLC2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_TYPE"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()




            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_OPAT"
            oCFLCreationParams.UniqueID = "CFLC3"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_TYPE"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_OPAT"
            oCFLCreationParams.UniqueID = "CFLCON"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_TYPE"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oEditText = aForm.Items.Item("4").Specific
            oEditText.ChooseFromListUID = "CFL_PRJ"
            oEditText.ChooseFromListAlias = "prjCode"

            Dim aGrid, aGrid1, aGrid2 As SAPbouiCOM.Grid
            oEditText = aForm.Items.Item("140").Specific
            oEditText.ChooseFromListUID = "CFLCON"
            oEditText.ChooseFromListAlias = "DocNum"

            aGrid = aForm.Items.Item("26").Specific
            aGrid1 = aForm.Items.Item("27").Specific
            aGrid2 = aForm.Items.Item("28").Specific
            oEditText = aForm.Items.Item("31").Specific
            oEditText.ChooseFromListUID = "CFL_BASE"
            oEditText.ChooseFromListAlias = "U_Z_PRJCODE"


            'oMatrix = aForm.Items.Item("12").Specific
            'oColumn = oMatrix.Columns.Item("V_0")
            'oColumn.ChooseFromListUID = "CFL1"
            'oColumn.ChooseFromListAlias = "U_Z_MODNAME"
            oEditTextColumn = aGrid.Columns.Item("U_Z_MODNAME")
            oEditTextColumn.ChooseFromListUID = "CFLR1"
            oEditTextColumn.ChooseFromListAlias = "U_Z_MODNAME"

            oEditTextColumn = aGrid1.Columns.Item("U_Z_MODNAME")
            oEditTextColumn.ChooseFromListUID = "CFLI1"
            oEditTextColumn.ChooseFromListAlias = "U_Z_MODNAME"
            oEditTextColumn = aGrid2.Columns.Item("U_Z_MODNAME")
            oEditTextColumn.ChooseFromListUID = "CFLE1"
            oEditTextColumn.ChooseFromListAlias = "U_Z_MODNAME"


            ''oColumn = oMatrix.Columns.Item("V_1")
            ''oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            ''oColumn = oMatrix.Columns.Item("V_2")
            ''oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            ''oColumn = oMatrix.Columns.Item("V_8")
            ''oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            ''oColumn = oMatrix.Columns.Item("Qty")
            ''oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


            'oMatrix = aForm.Items.Item("12").Specific
            'oColumn = oMatrix.Columns.Item("V_3")
            'oColumn.ChooseFromListUID = "CFL3"
            'oColumn.ChooseFromListAlias = "U_Z_ActName"

            'oEditTextColumn = aGrid.Columns.Item("U_Z_ACTNAME")
            'oEditTextColumn.ChooseFromListUID = "CFLR3"
            'oEditTextColumn.ChooseFromListAlias = "U_Z_ActName"

            'oEditTextColumn = aGrid1.Columns.Item("U_Z_ACTNAME")
            'oEditTextColumn.ChooseFromListUID = "CFLI3"
            'oEditTextColumn.ChooseFromListAlias = "U_Z_ActName"

            'oEditTextColumn = aGrid2.Columns.Item("U_Z_ACTNAME")
            'oEditTextColumn.ChooseFromListUID = "CFLE3"
            'oEditTextColumn.ChooseFromListAlias = "U_Z_ActName"


            'oMatrix = aForm.Items.Item("12").Specific
            'oColumn = oMatrix.Columns.Item("V_6")
            'oColumn.ChooseFromListUID = "CFL4"
            'oColumn.ChooseFromListAlias = "DocEntry"

            oEditTextColumn = aGrid.Columns.Item("U_Z_ORDENTRY")
            oEditTextColumn.ChooseFromListUID = "CFLR4"
            oEditTextColumn.ChooseFromListAlias = "DocEntry"

            oEditTextColumn = aGrid1.Columns.Item("U_Z_ORDENTRY")
            oEditTextColumn.ChooseFromListUID = "CFLI4"
            oEditTextColumn.ChooseFromListAlias = "DocEntry"

            oEditTextColumn = aGrid2.Columns.Item("U_Z_ORDENTRY")
            oEditTextColumn.ChooseFromListUID = "CFLE4"
            oEditTextColumn.ChooseFromListAlias = "DocEntry"

            oEditTextColumn = aGrid.Columns.Item("U_Z_CNTID")
            oEditTextColumn.ChooseFromListUID = "CFLC1"
            oEditTextColumn.ChooseFromListAlias = "DocNum"

            oEditTextColumn = aGrid1.Columns.Item("U_Z_CNTID")
            oEditTextColumn.ChooseFromListUID = "CFLC2"
            oEditTextColumn.ChooseFromListAlias = "DocNum"

            oEditTextColumn = aGrid2.Columns.Item("U_Z_CNTID")
            oEditTextColumn.ChooseFromListUID = "CFLC3"
            oEditTextColumn.ChooseFromListAlias = "DocNum"

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

            'oColumn.DisplayDesc = True
            'oMatrix = aForm.Items.Item("12").Specific
            'oColumn = oMatrix.Columns.Item("V_4")
            'oColumn.ChooseFromListUID = "CFL5"
            'oColumn.ChooseFromListAlias = "empID"


            oEditTextColumn = aGrid.Columns.Item("U_Z_EMPID")
            oEditTextColumn.ChooseFromListUID = "CFLR5"
            oEditTextColumn.ChooseFromListAlias = "empID"
            oEditTextColumn = aGrid1.Columns.Item("U_Z_EMPID")
            oEditTextColumn.ChooseFromListUID = "CFLI5"
            oEditTextColumn.ChooseFromListAlias = "empID"
            oEditTextColumn = aGrid2.Columns.Item("U_Z_EMPID")
            oEditTextColumn.ChooseFromListUID = "CFLE5"
            oEditTextColumn.ChooseFromListAlias = "empID"

            oEditTextColumn = aGrid2.Columns.Item("U_Z_EXPTYPE")
            oEditTextColumn.ChooseFromListUID = "CFL_EXP"
            oEditTextColumn.ChooseFromListAlias = "U_Z_EXPNAME"

            ' oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            ' oMatrix.AutoResizeColumns()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oGrid = aForm.Items.Item("26").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid = aForm.Items.Item("27").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid = aForm.Items.Item("28").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oMatrix = aForm.Items.Item("41").Specific
            For intRow As Integer = 1 To oMatrix.VisualRowCount
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", intRow, intRow)

            Next
            oGrid = aForm.Items.Item("100").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub


#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case 1
                oGrid = aForm.Items.Item("26").Specific
                If oGrid.DataTable.Rows.Count > 0 Then
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", oGrid.DataTable.Rows.Count - 1) <> "" Then
                        oGrid.DataTable.Rows.Add()
                        oGrid.DataTable.SetValue("LineId", oGrid.DataTable.Rows.Count - 1, "9999")
                        oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                        oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                    End If
                Else
                    oGrid.DataTable.Rows.Add()
                    oGrid.DataTable.SetValue("LineId", oGrid.DataTable.Rows.Count - 1, "9999")
                    oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                    oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                End If
                oGrid.Columns.Item("U_Z_MODNAME").Click(oGrid.DataTable.Rows.Count - 1)
                AssignLineNo(aForm)
                Exit Sub
            Case 2
                oGrid = aForm.Items.Item("27").Specific
                If oGrid.DataTable.Rows.Count > 0 Then
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", oGrid.DataTable.Rows.Count - 1) <> "" Then
                        oGrid.DataTable.Rows.Add()
                        oGrid.DataTable.SetValue("LineId", oGrid.DataTable.Rows.Count - 1, "9999")
                        oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                        oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                    End If
                Else
                    oGrid.DataTable.Rows.Add()
                    oGrid.DataTable.SetValue("LineId", oGrid.DataTable.Rows.Count - 1, "9999")
                    oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                    oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                End If
                oGrid.Columns.Item("U_Z_MODNAME").Click(oGrid.DataTable.Rows.Count - 1)
                AssignLineNo(aForm)
                Exit Sub
            Case 3
                oGrid = aForm.Items.Item("28").Specific
                If oGrid.DataTable.Rows.Count > 0 Then
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", oGrid.DataTable.Rows.Count - 1) <> "" Then
                        oGrid.DataTable.Rows.Add()
                        oGrid.DataTable.SetValue("LineId", oGrid.DataTable.Rows.Count - 1, "9999")
                        oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                        oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                    End If
                Else
                    oGrid.DataTable.Rows.Add()
                    oGrid.DataTable.SetValue("LineId", oGrid.DataTable.Rows.Count - 1, "9999")
                    oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                    oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                End If
                oGrid.Columns.Item("U_Z_MODNAME").Click(oGrid.DataTable.Rows.Count - 1)
                AssignLineNo(aForm)
                Exit Sub
            Case "11"
                Exit Sub
                'oMatrix = aForm.Items.Item("12").Specific
                'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
                'Case "2"
                '    oMatrix = aForm.Items.Item("13").Specific
                '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
            Case "5"
                oMatrix = aForm.Items.Item("41").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ3")
                If oMatrix.RowCount <= 0 Then
                    oMatrix.AddRow()
                End If
                oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                Try
                    If oEditText.Value <> "" Then
                        oMatrix.AddRow()
                        Select Case aForm.PaneLevel
                            Case "5"
                                oMatrix.ClearRowData(oMatrix.RowCount)
                                'oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                                'oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, "")
                        End Select
                    End If

                Catch ex As Exception
                    aForm.Freeze(False)
                    oMatrix.AddRow()
                End Try

                oMatrix.FlushToDataSource()
                For count = 1 To oDataSrc_Line.Size
                    oDataSrc_Line.SetValue("LineId", count - 1, count)
                Next
                oMatrix.LoadFromDataSource()
                oMatrix = aForm.Items.Item("41").Specific
                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", intRow, intRow)

                Next

                aForm.Freeze(False)
        End Select
        Try
            aForm.Freeze(True)

            'If oMatrix.RowCount <= 0 Then
            '    oMatrix.AddRow()
            'End If
            'oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
            'Try
            '    If oEditText.Value <> "" Then
            '        oMatrix.AddRow()
            '        Select Case aForm.PaneLevel
            '            Case "1"
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_11", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, "0")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, "0")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, "0")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_9", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_10", oMatrix.RowCount, "")
            '                oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
            '                oCheckBox.Checked = False
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "BOQ", oMatrix.RowCount, "")
            '                oApplication.Utilities.SetMatrixValues(oMatrix, "Qty", oMatrix.RowCount, "0")
            '        End Select
            '    End If
            'Catch ex As Exception
            '    aForm.Freeze(False)
            '    If oMatrix.RowCount <= 0 Then
            '        oMatrix.AddRow()
            '    End If
            'End Try

            'oMatrix.FlushToDataSource()
            'For count = 1 To oDataSrc_Line.Size
            '    oDataSrc_Line.SetValue("LineId", count - 1, count)
            'Next
            'oMatrix.LoadFromDataSource()
            'oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)

        End Try
    End Sub

    Private Sub DuplicateRow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "1"
                oGrid = aForm.Items.Item("26").Specific
            Case "2"
                oGrid = aForm.Items.Item("27").Specific
            Case "3"
                oGrid = aForm.Items.Item("28").Specific

        End Select
        Try
            aForm.Freeze(True)

            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.Rows.IsSelected(intRow) Then
                    Try
                        If oGrid.DataTable.GetValue("U_Z_MODNAME", intRow) <> "" Then
                            oGrid.DataTable.Rows.Add()
                            '                           strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID,"
                            'T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS,
                            '                           T1.U_Z_AMOUNT, T1.U_Z_ORDER, T1.U_Z_ORDENTRY, T1.U_Z_ORDNUM, T1.U_Z_STATUS, T1.U_Z_CMPDATE, T1.U_Z_CMPDATE
                            'from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='R'  and T0.U_Z_PRJCODE='" & strProject & "'"
                            oGrid.DataTable.SetValue("DocEntry", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("DocEntry", intRow))
                            oGrid.DataTable.SetValue("LineId", oGrid.DataTable.Rows.Count - 1, "9999")

                            oGrid.DataTable.SetValue("U_Z_MODNAME", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                            oGrid.DataTable.SetValue("U_Z_ACTNAME", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                            oGrid.DataTable.SetValue("U_Z_TYPE", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_TYPE", intRow))
                            oGrid.DataTable.SetValue("U_Z_EMPID", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                            oGrid.DataTable.SetValue("U_Z_POSITION", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                            oGrid.DataTable.SetValue("U_Z_FROMDATE", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                            oGrid.DataTable.SetValue("U_Z_TODATE", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                            oGrid.DataTable.SetValue("U_Z_QUANTITY", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                            oGrid.DataTable.SetValue("U_Z_MEASURE", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                            oGrid.DataTable.SetValue("U_Z_DAYS", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                            oGrid.DataTable.SetValue("U_Z_HOURS", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                            oGrid.DataTable.SetValue("U_Z_AMOUNT", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                            oGrid.DataTable.SetValue("U_Z_ORDER", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_ORDER", intRow))
                            oGrid.DataTable.SetValue("U_Z_ORDENTRY", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                            oGrid.DataTable.SetValue("U_Z_ORDNUM", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                            Try
                                If oCombocolumn.GetSelectedValue(intRow).Value <> "" Then
                                    oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_STATUS", intRow))
                                Else
                                    oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                                End If
                            Catch ex As Exception
                                oGrid.DataTable.SetValue("U_Z_STATUS", oGrid.DataTable.Rows.Count - 1, "I")
                            End Try

                            oGrid.DataTable.SetValue("U_Z_CMPDATE", oGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                            oGrid.DataTable.SetValue("U_Z_BOQ", oGrid.DataTable.Rows.Count - 1, "")

                        End If
                    Catch ex As Exception
                        aForm.Freeze(False)
                        If oGrid.DataTable.Rows.Count - 1 <= 0 Then
                            oGrid.DataTable.Rows.Add()
                        End If
                    End Try
                End If
            Next
            AssignLineNo(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

    Private Sub DuplicateRow_matrix(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "11"
                Exit Sub
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
    Private Function ValidateDeletion_Attachment(ByVal aForm As SAPbouiCOM.Form) As Boolean
        If aForm.PaneLevel <> 5 Then
            Return False
        End If
        If oApplication.SBO_Application.MessageBox("Deletion of Attachment will be permentaly deleted from project budget. Do you want to continue ?", , "Continue", "Cancel") = 2 Then
            Return False
        End If
        Select Case aForm.PaneLevel
            Case "5"
                oMatrix = aForm.Items.Item("41").Specific

        End Select
        For intRow As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(intRow) Then
                If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    If 1 = 1 Then ' oGrid.DataTable.GetValue("LineId", intSelectedMatrixrow) <> 9999 Then
                        Dim oGeneralService As SAPbobsCOM.GeneralService
                        Dim oGeneralData As SAPbobsCOM.GeneralData
                        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                        Dim oCompanyService As SAPbobsCOM.CompanyService
                        Dim oChild As SAPbobsCOM.GeneralData
                        Dim oChildren As SAPbobsCOM.GeneralDataCollection
                        Dim oTest As SAPbobsCOM.Recordset
                        oCompanyService = oApplication.Company.GetCompanyService
                        oGeneralService = oCompanyService.GetGeneralService("Z_PRJ")
                        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        'oGeneralParams.SetProperty("DocEntry", oGrid.DataTable.GetValue("DocEntry", intSelectedMatrixrow))
                        'oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim st As String = "Select * from [@Z_PRJ3] where DocEntry=" & oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "DocEntry", intRow)) & " and LineId=" & oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "LineID", intRow))
                        oTest.DoQuery(st)

                        ' oTest.DoQuery("Selec [@Z_PRJ1] set LineId=8888 where DocEntry=" & oGrid.DataTable.GetValue("DocEntry", intSelectedMatrixrow) & " and LineId=" & oGrid.DataTable.GetValue("LineId", intSelectedMatrixrow))
                        If oTest.RecordCount > 0 Then

                            'Dim intDocEntry As Integer = oApplication.Utilities.getMatrixValues(oMatrix, "DocEntry", intRow)
                            '' AddUDT(aForm, intDocEntry)
                            'oGeneralParams.SetProperty("DocEntry", intDocEntry)
                            'oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                            'oChildren = oGeneralData.Child("Z_PRJ3")
                            'Dim intlIne As Integer = oApplication.Utilities.getMatrixValues(oMatrix, "LineID", intRow)
                            'oChildren.Remove(intlIne - 1)
                            'oGeneralService.Update(oGeneralData)
                            ''  oMatrix.DeleteRow(intRow)
                            ''  LoadGridValues(aForm)
                            st = "Delete  from [@Z_PRJ3] where DocEntry=" & oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "DocEntry", intRow)) & " and LineId=" & oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "LineID", intRow))
                            oTest.DoQuery(st)
                            oMatrix.DeleteRow(intRow)
                            oMatrix = aForm.Items.Item("41").Specific
                            For intRow1 As Integer = 1 To oMatrix.VisualRowCount
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", intRow1, intRow1)

                            Next
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Return True
                        Else
                            oMatrix.DeleteRow(intRow)
                            oMatrix = aForm.Items.Item("41").Specific
                            For intRow1 As Integer = 1 To oMatrix.VisualRowCount
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", intRow1, intRow1)

                            Next
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Return True
                        End If

                    Else
                        oMatrix.DeleteRow(intRow)
                    End If
                Else
                    oMatrix.DeleteRow(intRow)
                End If
            End If
        Next

        Return True
    End Function
    Private Function ValidateDeletion(ByVal aForm As SAPbouiCOM.Form) As Boolean
        If aForm.PaneLevel > 3 Then
            Return False
        End If
        If oApplication.SBO_Application.MessageBox("Deletion of Details will be permentaly deleted from project budget. Do you want to continue ?", , "Continue", "Cancel") = 2 Then
            Return False
        End If
        Select Case aForm.PaneLevel
            Case "1"
                oGrid = aForm.Items.Item("26").Specific
            Case "2"
                oGrid = aForm.Items.Item("27").Specific
            Case "3"
                oGrid = aForm.Items.Item("28").Specific
        End Select
        If intSelectedMatrixrow < 0 Then
            Return True
        End If
        Dim strPrjCode, strActivity, strProcess, strMessage As String
        Dim otemp As SAPbobsCOM.Recordset
        strMessage = ""
        If oGrid.DataTable.GetValue("U_Z_MODNAME", intSelectedMatrixrow) <> "" Then

            oCombobox = aForm.Items.Item("4").Specific
            strPrjCode = oCombobox.Selected.Value
            strProcess = oGrid.DataTable.GetValue("U_Z_MODNAME", intSelectedMatrixrow)
            strActivity = oGrid.DataTable.GetValue("U_Z_ACTNAME", intSelectedMatrixrow)
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("select * from [@Z_OEXP] T0 Inner Join [@Z_EXP1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strPrjCode & "'")
            strMessage = "Project Code=" & strPrjCode
            otemp.DoQuery("select * from [@Z_OTIM] T0 Inner Join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strPrjCode & "' and U_Z_PRCNAME='" & strProcess.Replace("'", "''") & "' and U_Z_ACTNAME='" & strActivity.Replace("'", "''") & "'")
            If otemp.RecordCount > 0 Then
                strMessage = "Project Code : " & strPrjCode & " , Phase : " & strProcess & " , Activity : " & strActivity
                oApplication.Utilities.Message("Time Sheet already entered for this " & strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            otemp.DoQuery("select * from [@Z_OEXP] T0 Inner Join [@Z_EXP1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strPrjCode & "'")
            If otemp.RecordCount > 0 Then
                strMessage = "Project Code : " & strPrjCode & " , Phase : " & strProcess & " , Activity : " & strActivity
                oApplication.Utilities.Message("Expenses  already entered for this " & strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Dim oTest As SAPbobsCOM.Recordset
            If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                If oGrid.DataTable.GetValue("LineId", intSelectedMatrixrow) <> 9999 Then
                    Dim oGeneralService As SAPbobsCOM.GeneralService
                    Dim oGeneralData As SAPbobsCOM.GeneralData
                    Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                    Dim oCompanyService As SAPbobsCOM.CompanyService
                    Dim oChild As SAPbobsCOM.GeneralData
                    Dim oChildren As SAPbobsCOM.GeneralDataCollection
                    oCompanyService = oApplication.Company.GetCompanyService
                    oGeneralService = oCompanyService.GetGeneralService("Z_PRJ")
                    oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    'oGeneralParams.SetProperty("DocEntry", oGrid.DataTable.GetValue("DocEntry", intSelectedMatrixrow))
                    'oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim st As String = "Select * from [@Z_PRJ1] where DocEntry=" & oGrid.DataTable.GetValue("DocEntry", intSelectedMatrixrow) & " and LineId=" & oGrid.DataTable.GetValue("LineId", intSelectedMatrixrow)
                    oTest.DoQuery(st)

                    ' oTest.DoQuery("Selec [@Z_PRJ1] set LineId=8888 where DocEntry=" & oGrid.DataTable.GetValue("DocEntry", intSelectedMatrixrow) & " and LineId=" & oGrid.DataTable.GetValue("LineId", intSelectedMatrixrow))
                    If oTest.RecordCount > 0 Then

                        Dim intDocEntry As Integer = oTest.Fields.Item("DocEntry").Value
                        AddUDT(aForm, intDocEntry)
                        oGeneralParams.SetProperty("DocEntry", intDocEntry)
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                        oChildren = oGeneralData.Child("Z_PRJ1")
                        Dim intlIne As Integer = oTest.Fields.Item("LineId").Value
                        oChildren.Remove(intlIne - 1)
                        oGeneralService.Update(oGeneralData)
                        ' oGrid.DataTable.Rows.Remove(intSelectedMatrixrow)
                        ' AddUDT(aForm, intDocEntry)
                        LoadGridValues(aForm)
                        aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        'oApplication.SBO_Application.ActivateMenuItem(mnu_NEXT)
                        'oApplication.SBO_Application.ActivateMenuItem(mnu_PREVIOUS)
                        'aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        Return True
                    End If
                    oGrid.DataTable.Rows.Remove(intSelectedMatrixrow)
                Else
                    oGrid.DataTable.Rows.Remove(intSelectedMatrixrow)
                End If
            Else
                oGrid.DataTable.Rows.Remove(intSelectedMatrixrow)

            End If
        End If
        Return True
        Return True
    End Function

    Private Function ValidateDeletion_Matrix(ByVal aForm As SAPbouiCOM.Form) As Boolean
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
    Private Sub RefereshDeleteRow_Matrix(ByVal aForm As SAPbouiCOM.Form)
        If Me.MatrixId = "12" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
            'Else
            'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
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
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        If Me.MatrixId = "12" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
            'Else
            'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
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
        Try
            aform.Freeze(True)

            oCombobox = aform.Items.Item("401").Specific

            Dim dtFrom, dtTo, dtDate As Date
            Dim strdate As String

            strProject = oCombobox.Selected.Value
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If strProject = "" Then
                    oApplication.Utilities.Message("Project code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Freeze(False)
                    Return False
                End If
                oTemp.DoQuery("Select * from [@Z_HPRJ] where U_Z_PRJCODE='" & strProject & "'")
                If oTemp.RecordCount > 0 Then
                    oApplication.Utilities.Message("Project Budget already exists for this Project : " & strProject, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Freeze(False)
                    Return False
                End If
            End If

            Dim dblBudget, dblLineBudget As Double
            Dim strHours As String
            Dim dblHours, dblRate As Double
            dblBudget = 0
            dblLineBudget = 0
            dblRate = 0
            'Resource Matrix
            oCombobox = aform.Items.Item("17").Specific
            If oCombobox.Selected.Value <> "X" Then
                aform.Freeze(False)
                Return True
            End If

            strdate = oApplication.Utilities.getEdittextvalue(aform, "19")
            If strdate = "" Then
                oApplication.Utilities.Message("Project Budget From Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aform.Freeze(False)
                aform.Items.Item("19").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            Else
                dtFrom = oApplication.Utilities.GetDateTimeValue(strdate)
            End If
            strdate = oApplication.Utilities.getEdittextvalue(aform, "21")
            If strdate = "" Then
                oApplication.Utilities.Message("Project Budget End Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aform.Freeze(False)
                aform.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            Else
                dtTo = oApplication.Utilities.GetDateTimeValue(strdate)
            End If

            If dtFrom > dtTo Then
                oApplication.Utilities.Message("Project Budget End date should be Greater than or equal to From date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aform.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                aform.Freeze(False)
                Return False
            End If
            ' Return True
            strProject = oCombobox.Selected.Value
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If strProject = "" Then
                    oApplication.Utilities.Message("Project code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Freeze(False)
                    Return False
                End If
                oTemp.DoQuery("Select * from [@Z_HPRJ] where U_Z_PRJCODE='" & strProject & "'")
                If oTemp.RecordCount > 0 Then
                    oApplication.Utilities.Message("Project code already exists...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Freeze(False)
                    Return False
                End If
            End If


            'Resource Matrix

            oGrid = aform.Items.Item("26").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                strCode = oGrid.DataTable.GetValue("U_Z_MODNAME", intRow)
                If strCode <> "" Then
                    strActivity = oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow)
                    If strActivity = "" Then
                        oApplication.Utilities.Message("Activity detail is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aform.Freeze(False)
                        Return False
                    End If
                    strHours = oGrid.DataTable.GetValue("U_Z_DAYS", intRow)
                    If strHours <> "" Then
                        dblHours = oApplication.Utilities.getDocumentQuantity(strHours)
                        If (dblHours > 0) Then
                            dblHours = dblHours * 8
                            Dim oTest As SAPbobsCOM.Recordset
                            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            ' oEditTextColumn = oGrid.Columns.Item("U_Z_EMPID").TitleObject

                            If oGrid.DataTable.GetValue("U_Z_EMPID", intRow) <> "" Then
                                oTest.DoQuery("Select isnull(U_Daily_rate,0) from OHEM where empID=" & oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                                dblRate = oTest.Fields.Item(0).Value
                            Else
                                dblRate = 1
                            End If
                            dblRate = dblRate * oApplication.Utilities.getDocumentQuantity(strHours)
                            oGrid.DataTable.SetValue("U_Z_HOURS", intRow, dblHours)
                            'oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", intRow, dblRate)
                        End If
                    End If

                    strcode1 = oGrid.DataTable.GetValue("U_Z_HOURS", intRow)
                    ' oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific
                    If 1 = 1 Then 'oCombobox.Selected.Value = "R" Then
                        If strcode1 = "" Then
                            oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'oMatrix.Columns.Item("V_1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If
                        If CInt(strcode1) <= 0 Then
                            oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If
                    End If
                    'oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific
                    'If oCombobox.Selected.Value = "I" Then
                    '    If oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "Qty", intRow)) <= 0 Then
                    '        oApplication.Utilities.Message("Required Quantity should be greater than or equal to Zero. Line number:" & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        oMatrix.Columns.Item("Qty").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        Return False
                    '    End If
                    '    If oApplication.Utilities.getMatrixValues(oMatrix, "BOQ", intRow) = "" Then
                    '        '  oApplication.Utilities.Message("BOQ Not entered for the line number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '        '  Return False
                    '    End If
                    'End If

                    strdate = oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow)
                    oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                    Try
                        If oCombocolumn.GetSelectedValue(intRow).Value = "C" Then
                            If strdate = "" Then
                                oApplication.Utilities.Message("Completion date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oGrid.Columns.Item("U_Z_CMPDATE").Click(intRow)
                                aform.Freeze(False)
                                Return False
                            End If
                        End If
                    Catch ex As Exception

                    End Try

                    strdate = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                    If strdate = "" Then
                        oApplication.Utilities.Message("From date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oGrid.Columns.Item("U_Z_FROMDATE").Click(intRow)
                        aform.Freeze(False)
                        Return False
                    Else
                        dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                        If dtDate >= dtFrom And dtDate <= dtTo Then
                        Else
                            oApplication.Utilities.Message("From date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_FROMDATE").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If
                    End If

                    strdate = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                    If strdate = "" Then
                        oApplication.Utilities.Message("End date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oGrid.Columns.Item("U_Z_TODATE").Click(intRow)
                        aform.Freeze(False)
                        Return False
                    Else
                        dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                        If dtDate >= dtFrom And dtDate <= dtTo Then
                        Else
                            oApplication.Utilities.Message("End date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_TODATE").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If

                    End If
                Else
                    oGrid.DataTable.Rows.Remove(intRow)
                End If
            Next

            'Item Matrix
            oGrid = aform.Items.Item("27").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                strCode = oGrid.DataTable.GetValue("U_Z_MODNAME", intRow)
                If strCode <> "" Then
                    strActivity = oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow)
                    If strActivity = "" Then
                        oApplication.Utilities.Message("Activity detail is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aform.Freeze(False)
                        Return False
                    End If
                    strHours = oGrid.DataTable.GetValue("U_Z_DAYS", intRow)
                    If strHours <> "" Then
                        dblHours = oApplication.Utilities.getDocumentQuantity(strHours)
                        If (dblHours > 0) Then
                            dblHours = dblHours * 8
                            Dim oTest As SAPbobsCOM.Recordset
                            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            ' oEditTextColumn = oGrid.Columns.Item("U_Z_EMPID").TitleObject
                            If oGrid.DataTable.GetValue("U_Z_EMPID", intRow) <> "" Then
                                oTest.DoQuery("Select isnull(U_Daily_rate,0) from OHEM where empID=" & oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                                dblRate = oTest.Fields.Item(0).Value
                            Else
                                dblRate = 1
                            End If
                            dblRate = dblRate * oApplication.Utilities.getDocumentQuantity(strHours)
                            oGrid.DataTable.SetValue("U_Z_HOURS", intRow, dblHours)
                            'oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", intRow, dblRate)
                        End If
                    End If

                    strcode1 = oGrid.DataTable.GetValue("U_Z_HOURS", intRow)
                    ' oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific
                    'If 1 = 1 Then 'oCombobox.Selected.Value = "R" Then
                    '    If strcode1 = "" Then
                    '        oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        'oMatrix.Columns.Item("V_1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                    '        Return False
                    '    End If
                    '    If CInt(strcode1) <= 0 Then
                    '        oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                    '        Return False
                    '    End If
                    'End If
                    ' oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific
                    If 1 = 1 Then 'oCombobox.Selected.Value = "I" Then
                        If oApplication.Utilities.getDocumentQuantity(oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow)) <= 0 Then
                            oApplication.Utilities.Message("Required Quantity should be greater than or equal to Zero. Line number:" & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_QUANTITY").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If
                        'If oApplication.Utilities.getMatrixValues(oMatrix, "BOQ", intRow) = "" Then
                        '    '  oApplication.Utilities.Message("BOQ Not entered for the line number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        '    '  Return False
                        'End If
                    End If

                    strdate = oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow)
                    oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                    Try


                        If oCombocolumn.GetSelectedValue(intRow).Value = "C" Then
                            If strdate = "" Then
                                oApplication.Utilities.Message("Completion date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oGrid.Columns.Item("U_Z_CMPDATE").Click(intRow)
                                aform.Freeze(False)
                                Return False
                            End If
                        End If

                    Catch ex As Exception

                    End Try
                    strdate = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                    If strdate = "" Then
                        oApplication.Utilities.Message("From date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oGrid.Columns.Item("U_Z_FROMDATE").Click(intRow)
                        aform.Freeze(False)
                        Return False
                    Else
                        dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                        If dtDate >= dtFrom And dtDate <= dtTo Then
                        Else
                            oApplication.Utilities.Message("From date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_FROMDATE").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If
                    End If

                    strdate = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                    If strdate = "" Then
                        oApplication.Utilities.Message("End date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oGrid.Columns.Item("U_Z_TODATE").Click(intRow)
                        aform.Freeze(False)
                        Return False
                    Else
                        dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                        If dtDate >= dtFrom And dtDate <= dtTo Then
                        Else
                            oApplication.Utilities.Message("End date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_TODATE").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If

                    End If
                Else
                    oGrid.DataTable.Rows.Remove(intRow)

                End If
            Next


            'Item Matrix
            oGrid = aform.Items.Item("28").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                strCode = oGrid.DataTable.GetValue("U_Z_MODNAME", intRow)
                If strCode <> "" Then
                    strActivity = oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow)
                    If strActivity = "" Then
                        oApplication.Utilities.Message("Activity detail is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aform.Freeze(False)
                        Return False
                    End If
                    strHours = oGrid.DataTable.GetValue("U_Z_DAYS", intRow)
                    If strHours <> "" Then
                        dblHours = oApplication.Utilities.getDocumentQuantity(strHours)
                        If (dblHours > 0) Then
                            dblHours = dblHours * 8
                            Dim oTest As SAPbobsCOM.Recordset
                            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            ' oEditTextColumn = oGrid.Columns.Item("U_Z_EMPID").TitleObject
                            If oGrid.DataTable.GetValue("U_Z_EMPID", intRow) <> "" Then
                                oTest.DoQuery("Select isnull(U_Daily_rate,0) from OHEM where empID=" & oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                                dblRate = oTest.Fields.Item(0).Value
                            Else
                                dblRate = 1
                            End If
                            dblRate = dblRate * oApplication.Utilities.getDocumentQuantity(strHours)
                            oGrid.DataTable.SetValue("U_Z_HOURS", intRow, dblHours)
                            'oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", intRow, dblRate)
                        End If
                    End If

                    strcode1 = oGrid.DataTable.GetValue("U_Z_HOURS", intRow)
                    ' oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific
                    'If 1 = 1 Then 'oCombobox.Selected.Value = "R" Then
                    '    If strcode1 = "" Then
                    '        oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        'oMatrix.Columns.Item("V_1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                    '        Return False
                    '    End If
                    '    If CInt(strcode1) <= 0 Then
                    '        oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                    '        Return False
                    '    End If
                    'End If
                    ' oCombobox = oMatrix.Columns.Item("Type").Cells.Item(intRow).Specific


                    strdate = oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow)
                    oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                    Try
                        If oCombocolumn.GetSelectedValue(intRow).Value = "C" Then
                            If strdate = "" Then
                                oApplication.Utilities.Message("Completion date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oForm.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oGrid.Columns.Item("U_Z_CMPDATE").Click(intRow)
                                aform.Freeze(False)
                                Return False
                            End If
                        End If
                    Catch ex As Exception

                    End Try

                    strdate = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                    If strdate = "" Then
                        oApplication.Utilities.Message("From date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oGrid.Columns.Item("U_Z_FROMDATE").Click(intRow)
                        aform.Freeze(False)
                        Return False
                    Else
                        dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                        If dtDate >= dtFrom And dtDate <= dtTo Then
                        Else
                            oApplication.Utilities.Message("From date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_FROMDATE").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If
                    End If

                    strdate = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                    If strdate = "" Then
                        oApplication.Utilities.Message("End date is missing. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oGrid.Columns.Item("U_Z_TODATE").Click(intRow)
                        aform.Freeze(False)
                        Return False
                    Else
                        dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                        If dtDate >= dtFrom And dtDate <= dtTo Then
                        Else
                            oApplication.Utilities.Message("End date should be between project budget From and End date. Line Number  : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oGrid.Columns.Item("U_Z_TODATE").Click(intRow)
                            aform.Freeze(False)
                            Return False
                        End If

                    End If
                Else
                    oGrid.DataTable.Rows.Remove(intRow)

                End If
            Next





            'oMatrix = aform.Items.Item("12").Specific
            'If oMatrix.RowCount <= 0 Then
            '    oApplication.Utilities.Message("Process details are missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'dblBudget = CDbl(oApplication.Utilities.getEdittextvalue(aform, "8"))
            'If dblBudget <> dblLineBudget Then
            '    If oApplication.SBO_Application.MessageBox("Total man days does not match with Line man days. Do you want to save this document ? ", , "Continue", "Cancel") = 2 Then
            '        Return False
            '    Else


            '    End If
            'End If
            aform.Freeze(False)
            Return True

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False

        End Try
    End Function
#End Region

#Region "LoadGridValues"
    Private Sub LoadGridValues(ByVal aForm As SAPbouiCOM.Form)
        Dim strQuery, strProject As String
        Try

            aForm.Freeze(True)
            oCombobox = aForm.Items.Item("401").Specific
            Try
                strProject = oApplication.Utilities.getEdittextvalue(aForm, "4") ' oCombobox.Selected.Value
                oCombobox.Select(strProject, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception
                strProject = ""
            End Try

            If strProject <> "" Then
                Dim oTest, otest1 As SAPbobsCOM.Recordset
                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTest.DoQuery("Select * from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry and T0.U_Z_PRJCODE='" & strProject & "'")
                For intRow As Integer = 0 To oTest.RecordCount - 1
                    otest1.DoQuery("Update [@Z_PRJ1] set LineId=" & intRow + 1 & "where DocEntry=" & oTest.Fields.Item("DocEntry").Value & " and LineId=" & oTest.Fields.Item("LineId").Value)
                    oTest.MoveNext()
                Next
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ ,T1.U_Z_EXPTYPE from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='R'  and T0.U_Z_PRJCODE='" & strProject & "'"

                oGrid = aForm.Items.Item("26").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "R")

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='I' and T0.U_Z_PRJCODE='" & strProject & "'"
                oGrid = aForm.Items.Item("27").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "I")

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_EXPTYPE,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='E' and T0.U_Z_PRJCODE='" & strProject & "'"
                oGrid = aForm.Items.Item("28").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T0.U_Z_PRJCODE='" & strProject & "'"
                oGrid = aForm.Items.Item("100").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                ' FormatGrid(aForm, oGrid, "E")

            Else
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("26").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "R")
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("27").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "I")
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_EXPTYPE,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ ,T1.U_Z_EXPTYPE from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("28").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_CNTID,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("100").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

            End If
            databind(aForm)
            oGrid = aForm.Items.Item("26").Specific
            FormatGrid(aForm, oGrid, "R")
            oGrid = aForm.Items.Item("27").Specific
            FormatGrid(aForm, oGrid, "I")
            oGrid = aForm.Items.Item("28").Specific
            FormatGrid(aForm, oGrid, "E")
            aForm.Freeze(False)
            oGrid = aForm.Items.Item("100").Specific
            FormatGrid(aForm, oGrid, "A")
            aForm.Items.Item("100").Enabled = False
            AssignLineNo(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub PopulateBudgetfromTemplate(ByVal aForm As SAPbouiCOM.Form, ByVal aDocEntry As String)
        Dim strQuery, strProject As String
        Try

            aForm.Freeze(True)
            oCombobox = aForm.Items.Item("401").Specific
            Try
                strProject = aDocEntry
            Catch ex As Exception
                strProject = ""
            End Try

            If oApplication.SBO_Application.MessageBox("Do you want to copy the budget details from selected Template : " & strProject & "?", , "Continue", "Cancel") = 2 Then
                Exit Sub
            End If

            If strProject <> "" Then
                Dim oTest, otest1 As SAPbobsCOM.Recordset
                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oTest.DoQuery("Select * from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry and T0.U_Z_PRJCODE='" & strProject & "'")
                'For intRow As Integer = 0 To oTest.RecordCount - 1
                '    otest1.DoQuery("Update [@Z_PRJ1] set LineId=" & intRow + 1 & "where DocEntry=" & oTest.Fields.Item("DocEntry").Value & " and LineId=" & oTest.Fields.Item("LineId").Value)
                '    oTest.MoveNext()
                'Next
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ ,T1.U_Z_EXPTYPE from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='R'  and T0.U_Z_PRJCODE='" & strProject & "'"

                oGrid = aForm.Items.Item("26").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "R")

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='I' and T0.U_Z_PRJCODE='" & strProject & "'"
                oGrid = aForm.Items.Item("27").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "I")

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_EXPTYPE,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ  from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='E' and T0.U_Z_PRJCODE='" & strProject & "'"
                oGrid = aForm.Items.Item("28").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where T0.U_Z_PRJCODE='" & strProject & "'"
                oGrid = aForm.Items.Item("100").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                ' FormatGrid(aForm, oGrid, "E")

            Else
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("26").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "R")
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("27").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)
                'FormatGrid(aForm, oGrid, "I")
                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_EXPTYPE,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ ,T1.U_Z_EXPTYPE from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("28").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

                strQuery = "select T1.DocEntry,T1.LineId,T1.U_Z_MODNAME ,T1.U_Z_ACTNAME,T1.U_Z_TYPE,T1.U_Z_EMPID, T1.U_Z_POSITION, T1.U_Z_FROMDATE, T1.U_Z_TODATE, T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOURS, T1.U_Z_AMOUNT, T1.U_Z_ORDER ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_CMPDATE,T1.U_Z_BOQ,T1.U_Z_EXPTYPE  from [@Z_DHPRJ] T0 inner Join [@Z_DPRJ1] T1 on T1.DocEntry=T0.DocEntry where 1=2"
                oGrid = aForm.Items.Item("100").Specific
                oGrid.DataTable.ExecuteQuery(strQuery)

            End If
            databind(aForm)
            oGrid = aForm.Items.Item("26").Specific
            FormatGrid(aForm, oGrid, "R")
            oGrid = aForm.Items.Item("27").Specific
            FormatGrid(aForm, oGrid, "I")
            oGrid = aForm.Items.Item("28").Specific
            FormatGrid(aForm, oGrid, "E")
            aForm.Freeze(False)
            oGrid = aForm.Items.Item("100").Specific
            FormatGrid(aForm, oGrid, "A")
            aForm.Items.Item("100").Enabled = False
            AssignLineNo(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
    Private Sub FormatGrid(ByVal aform As SAPbouiCOM.Form, ByVal aGrid As SAPbouiCOM.Grid, ByVal aChoice As String)
        Try
            aform.Freeze(True)
            aGrid.Columns.Item("DocEntry").Visible = False
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
                aGrid.Columns.Item("LineId").Visible = False
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
            oEditTextColumn = aGrid.Columns.Item("U_Z_CNTID")
            oEditTextColumn.LinkedObjectType = "Z_OPAT"

            aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            'T1.U_Z_QUANTITY, T1.U_Z_MEASURE, T1.U_Z_DAYS, T1.U_Z_HOUS, T1.U_Z_AMOUNT, T1.U_Z_ORDNUM ,T1.U_Z_ORDENTRY,T1.U_Z_ORDNUM ,T1.U_Z_STATUS,T1.U_Z_BOQ  from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where T1.U_Z_TYPE='E' and T0.U_Z_PRJCODE='" & strProject & "'"

            aform.Freeze(False)
        Catch ex As Exception
            aform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
#End Region

    Private Function CheckProject(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strProject, strQuery As String
        Dim Orec As SAPbobsCOM.Recordset
        'oCombobox = aForm.Items.Item("4").Specific
        'strProject = oCombobox.Selected.Value
        strProject = oApplication.Utilities.getEdittextvalue(aForm, "4")
        strQuery = "Select * from [@Z_TIM1] where U_Z_PrjCode='" & strProject & "'"
        Orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Orec.DoQuery(strQuery)
        If Orec.RecordCount > 0 Then
            oApplication.Utilities.Message("Time sheet entry already exists for this project. you can not remove ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        strQuery = "Select * from [@Z_EXP1] where U_Z_PrjCode='" & strProject & "'"
        Orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Orec.DoQuery(strQuery)
        If Orec.RecordCount > 0 Then
            oApplication.Utilities.Message("Expenses entry already exists for this project. you can not remove ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        Return True


    End Function

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_BudgetEntry
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("6").Enabled = False
                        oForm.Items.Item("35").Enabled = False
                        oForm.Items.Item("37").Enabled = False
                        oForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        LoadGridValues(oForm)
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
                        ' RefereshDeleteRow(oForm)
                    Else
                        If oForm.PaneLevel = 5 Then
                            If ValidateDeletion_Attachment(oForm) = True Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else

                            If ValidateDeletion(oForm) = False Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                                BubbleEvent = False
                            End If
                        End If
                    End If
                Case "1283"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then

                        If oApplication.SBO_Application.MessageBox("Do you want to remove the project budget details ?", , "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If CheckProject(oForm) = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            oCombobox = oForm.Items.Item("401").Specific
                            ProjectDetailstoSAP(oCombobox.Selected.Value, "Delete")
                        End If
                    Else
                        ' oCombobox = oForm.Items.Item("4").Specific
                        ' ProjectDetailstoSAP(oCombobox.Selected.Value, "Delete")
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = False
                        'oForm.Items.Item("8").Enabled = True
                        oForm.Items.Item("30").Visible = True
                        oForm.Items.Item("31").Visible = True
                        oForm.Items.Item("35").Enabled = False
                        oForm.Items.Item("37").Enabled = False
                        oForm.Items.Item("8")
                        ' Addmode(oForm)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = True
                        oForm.Items.Item("8").Enabled = True
                        oForm.Items.Item("30").Visible = False
                        oForm.Items.Item("31").Visible = False
                        oForm.Items.Item("35").Enabled = True
                        oForm.Items.Item("37").Enabled = False
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


    Private Function AddUDT(ByVal aform As SAPbouiCOM.Form, ByVal aDocEntry As Integer) As Boolean
        Try
            aform.Freeze(True)

            Dim strDocEntry, strLineId, firstName, LastName, strBPCode, strBPName As String
            Dim oRec As SAPbobsCOM.Recordset
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren, ochildern1 As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            oCompanyService = oApplication.Company.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("Z_PRJ")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombobox = aform.Items.Item("401").Specific

            Dim strCode As String = oCombobox.Selected.Value
            Dim blnExits As Boolean = False
            oRec.DoQuery("SElect * from [@Z_HPRJ] where U_Z_PRJCODE='" & strCode & "'")
            If oRec.RecordCount > 0 Then
                aDocEntry = oRec.Fields.Item("DocEntry").Value
                blnExits = True
            Else
                oRec.DoQuery("select * from ONNM where Objectcode='Z_PRJ'")
                aDocEntry = oRec.Fields.Item("AutoKey").Value
            End If
            Dim strPrjName, strstatus, strBudget, strExpe, strFromdate, strTodate, strApproval As String
            strPrjName = oApplication.Utilities.getEdittextvalue(aform, "6")
            oCombobox = aform.Items.Item("33").Specific
            Try
                strApproval = oCombobox.Selected.Value
            Catch ex As Exception
                strApproval = "N"
            End Try
            oCombobox = aform.Items.Item("17").Specific
            Try
                strstatus = oCombobox.Selected.Value
            Catch ex As Exception
                strstatus = "E"
            End Try
            Dim strCardCode1, strCardName As String

            strBudget = oApplication.Utilities.getEdittextvalue(aform, "8")
            strExpe = oApplication.Utilities.getEdittextvalue(aform, "23")
            strFromdate = oApplication.Utilities.getEdittextvalue(aform, "19")
            strTodate = oApplication.Utilities.getEdittextvalue(aform, "21")
            strCardCode1 = oApplication.Utilities.getEdittextvalue(aform, "35")
            strCardName = oApplication.Utilities.getEdittextvalue(aform, "37")

            Dim dtFrom, dtTo As Date
            dtFrom = oApplication.Utilities.GetDateTimeValue(strFromdate)
            dtTo = oApplication.Utilities.GetDateTimeValue(strTodate)
            Dim dblBudget, dblExp As Double
            dblBudget = oApplication.Utilities.getDocumentQuantity(strBudget)
            dblExp = oApplication.Utilities.getDocumentQuantity(strExpe)
            oRec.DoQuery("Select * from OPRJ where prjCode='" & strCode & "'")
            oApplication.Utilities.setEdittextvalue(aform, "35", oRec.Fields.Item("U_Z_CARDCODE").Value)
            oApplication.Utilities.setEdittextvalue(aform, "37", oRec.Fields.Item("U_Z_CARDNAME").Value)

            If blnExits = False Then
                oGeneralData.SetProperty("U_Z_PRJCODE", strCode)
                oGeneralData.SetProperty("U_Z_PRJNAME", strPrjName)
                oGeneralData.SetProperty("U_Z_BUDGET", dblBudget)
                oGeneralData.SetProperty("U_Z_TOTALEXPENSE", dblExp)
                oGeneralData.SetProperty("U_Z_FROMDATE", dtFrom)
                oGeneralData.SetProperty("U_Z_TODATE", dtTo)
                oGeneralData.SetProperty("U_Z_STATUS", strstatus)
                oGeneralData.SetProperty("U_Z_APPROVAL", strApproval)
                oGeneralData.SetProperty("U_Z_CARDCODE", oApplication.Utilities.getEdittextvalue(aform, "35"))
                oGeneralData.SetProperty("U_Z_CARDNAME", oApplication.Utilities.getEdittextvalue(aform, "37"))
                oGeneralData.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                oChildren = oGeneralData.Child("Z_PRJ1")
                oGrid = aform.Items.Item("26").Specific
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", intRow) <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                        oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                        oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                        oChild.SetProperty("U_Z_TYPE", "R")
                        oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                        oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                        Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                        If st <> "" Then
                            oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                        End If
                        st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                        If st <> "" Then
                            oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                        End If

                        oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                        oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                        oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                        oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                        oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                        oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                        oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                        If oCheckBoxCOlumn.IsChecked(intRow) Then
                            oChild.SetProperty("U_Z_ORDER", "Y")
                        Else
                            oChild.SetProperty("U_Z_ORDER", "N")
                        End If

                        oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                        oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                        oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                        Try
                            oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                        Catch ex As Exception
                            oChild.SetProperty("U_Z_STATUS", "I")
                        End Try

                        Try
                            oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                        Catch ex As Exception

                        End Try
                        oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                        oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                    End If
                Next

                oGrid = aform.Items.Item("27").Specific
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", intRow) <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                        oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                        oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                        oChild.SetProperty("U_Z_TYPE", "I")
                        oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                        oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                        Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                        If st <> "" Then
                            oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                        End If
                        st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                        If st <> "" Then
                            oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                        End If

                        oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                        oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                        oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                        oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                        oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                        oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                        If oCheckBoxCOlumn.IsChecked(intRow) Then
                            oChild.SetProperty("U_Z_ORDER", "Y")
                        Else
                            oChild.SetProperty("U_Z_ORDER", "N")
                        End If

                        oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                        oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                        oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                        Try
                            oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                        Catch ex As Exception
                            oChild.SetProperty("U_Z_STATUS", "I")
                        End Try

                        Try
                            oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                        Catch ex As Exception

                        End Try
                        oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                        oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                        oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                    End If
                Next

                oGrid = aform.Items.Item("28").Specific
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", intRow) <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                        oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                        oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                        oChild.SetProperty("U_Z_TYPE", "E")
                        oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                        oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                        Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                        If st <> "" Then
                            oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                        End If
                        st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                        If st <> "" Then
                            oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                        End If

                        oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                        oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                        oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                        oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                        oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                        oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                        If oCheckBoxCOlumn.IsChecked(intRow) Then
                            oChild.SetProperty("U_Z_ORDER", "Y")
                        Else
                            oChild.SetProperty("U_Z_ORDER", "N")
                        End If

                        oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                        oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                        oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                        Try
                            oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                        Catch ex As Exception
                            oChild.SetProperty("U_Z_STATUS", "I")
                        End Try

                        Try
                            oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                        Catch ex As Exception

                        End Try
                        oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                        oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                        oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                    End If
                Next
                ochildern1 = oGeneralData.Child("Z_PRJ3")
                oMatrix = aform.Items.Item("41").Specific
                Dim oDel As SAPbobsCOM.Recordset
                oDel = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oDel.DoQuery("Delete from ""@Z_PRJ3"" where ""DocEntry""= " & aDocEntry)

                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow) <> "" Then
                        oChild = ochildern1.Add()
                        oChild.SetProperty("U_Z_FILEPATH", oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow))
                        oChild.SetProperty("U_Z_FILENAME", oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow))
                    End If
                Next



                oGeneralService.Add(oGeneralData)
            Else
                Dim oCheckRs As SAPbobsCOM.Recordset
                oCheckRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oGeneralParams.SetProperty("DocEntry", aDocEntry)
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                oGeneralData.SetProperty("U_Z_PRJCODE", strCode)
                oGeneralData.SetProperty("U_Z_PRJNAME", strPrjName)
                oGeneralData.SetProperty("U_Z_BUDGET", dblBudget)
                oGeneralData.SetProperty("U_Z_TOTALEXPENSE", dblExp)
                oGeneralData.SetProperty("U_Z_FROMDATE", dtFrom)
                oGeneralData.SetProperty("U_Z_TODATE", dtTo)
                oGeneralData.SetProperty("U_Z_STATUS", strstatus)
                oGeneralData.SetProperty("U_Z_APPROVAL", strApproval)
                oGeneralData.SetProperty("U_Z_CARDCODE", oApplication.Utilities.getEdittextvalue(aform, "35"))
                oGeneralData.SetProperty("U_Z_CARDNAME", oApplication.Utilities.getEdittextvalue(aform, "37"))
                oGeneralData.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                oChildren = oGeneralData.Child("Z_PRJ1")
                oGrid = aform.Items.Item("26").Specific
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", intRow) <> "" Then
                        strLineId = oGrid.DataTable.GetValue("LineId", intRow)
                        '  oCheckRs.DoQuery("Select * from [@Z_PRJ1] where DocEntry=" & aDocEntry & " and LineId=" & strLineId & " U_Z_TYPE='R'")
                        If strLineId = 0 Then
                            strLineId = 9999
                        End If
                        If strLineId = 9999 Then
                            oChild = oChildren.Add()
                            oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                            oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                            oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                            oChild.SetProperty("U_Z_TYPE", "R")
                            oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                            oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                            Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                            End If
                            st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                            End If

                            oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                            oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                            oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                            oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                            oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                            oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                            oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                            If oCheckBoxCOlumn.IsChecked(intRow) Then
                                oChild.SetProperty("U_Z_ORDER", "Y")
                            Else
                                oChild.SetProperty("U_Z_ORDER", "N")
                            End If
                            oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                            oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                            Try
                                oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                            Catch ex As Exception
                                oChild.SetProperty("U_Z_STATUS", "I")
                            End Try
                            Try
                                oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                            Catch ex As Exception
                            End Try
                            oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                            oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                        Else
                            oChild = oChildren.Item(CInt(strLineId) - 1)
                            oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                            oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                            oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                            oChild.SetProperty("U_Z_TYPE", "R")
                            oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                            oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                            Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                            End If
                            st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                            End If

                            oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                            oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                            oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                            oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                            oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                            oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                            oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                            If oCheckBoxCOlumn.IsChecked(intRow) Then
                                oChild.SetProperty("U_Z_ORDER", "Y")
                            Else
                                oChild.SetProperty("U_Z_ORDER", "N")
                            End If

                            oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                            oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")

                            Try
                                oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                            Catch ex As Exception
                                oChild.SetProperty("U_Z_STATUS", "I")
                            End Try
                            Try
                                oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                            Catch ex As Exception
                            End Try
                            oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                            oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                        End If
                    End If
                Next
                oGrid = aform.Items.Item("27").Specific
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", intRow) <> "" Then
                        strLineId = oGrid.DataTable.GetValue("LineId", intRow)
                        If strLineId = 0 Then
                            strLineId = 9999
                        End If
                        If strLineId = 9999 Then
                            oChild = oChildren.Add()
                            oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                            oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                            oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                            oChild.SetProperty("U_Z_TYPE", "I")
                            oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                            oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                            Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                            End If
                            st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                            End If

                            oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                            oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                            oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                            oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                            oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                            oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                            oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                            If oCheckBoxCOlumn.IsChecked(intRow) Then
                                oChild.SetProperty("U_Z_ORDER", "Y")
                            Else
                                oChild.SetProperty("U_Z_ORDER", "N")
                            End If

                            oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                            oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                            Try
                                oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                            Catch ex As Exception
                                oChild.SetProperty("U_Z_STATUS", "I")
                            End Try

                            Try
                                oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                            Catch ex As Exception

                            End Try
                            oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                            oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                        Else
                            oChild = oChildren.Item(CInt(strLineId) - 1)
                            oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                            oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                            oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                            oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                            oChild.SetProperty("U_Z_TYPE", "I")
                            oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                            oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                            Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                            End If
                            st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                            End If

                            oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                            oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                            oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                            oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                            oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                            oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                            If oCheckBoxCOlumn.IsChecked(intRow) Then
                                oChild.SetProperty("U_Z_ORDER", "Y")
                            Else
                                oChild.SetProperty("U_Z_ORDER", "N")
                            End If

                            oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                            oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")

                            Try
                                oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                            Catch ex As Exception
                                oChild.SetProperty("U_Z_STATUS", "I")
                            End Try
                            Try
                                oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                            Catch ex As Exception

                            End Try
                            oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                            oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                        End If
                    End If
                Next
                oGrid = aform.Items.Item("28").Specific
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("U_Z_MODNAME", intRow) <> "" Then
                        strLineId = oGrid.DataTable.GetValue("LineId", intRow)
                        If strLineId = 0 Then
                            strLineId = 9999
                        End If
                        If strLineId = 9999 Then
                            oChild = oChildren.Add()
                            oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                            oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                            oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                            oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                            oChild.SetProperty("U_Z_TYPE", "E")
                            oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                            oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                            Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                            End If
                            st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                            End If

                            oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                            oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                            oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                            oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                            oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                            oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                            If oCheckBoxCOlumn.IsChecked(intRow) Then
                                oChild.SetProperty("U_Z_ORDER", "Y")
                            Else
                                oChild.SetProperty("U_Z_ORDER", "N")
                            End If

                            oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                            oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                            Try
                                oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                            Catch ex As Exception
                                oChild.SetProperty("U_Z_STATUS", "I")
                            End Try

                            Try
                                oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                            Catch ex As Exception

                            End Try
                            oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                            oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                        Else
                            oChild = oChildren.Item(CInt(strLineId) - 1)
                            oChild.SetProperty("U_Z_CNTID", oGrid.DataTable.GetValue("U_Z_CNTID", intRow))
                            oChild.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEdittextvalue(aform, "140"))
                            oChild.SetProperty("U_Z_MODNAME", oGrid.DataTable.GetValue("U_Z_MODNAME", intRow))
                            oChild.SetProperty("U_Z_ACTNAME", oGrid.DataTable.GetValue("U_Z_ACTNAME", intRow))
                            oChild.SetProperty("U_Z_TYPE", "E")
                            oChild.SetProperty("U_Z_EMPID", oGrid.DataTable.GetValue("U_Z_EMPID", intRow))
                            oChild.SetProperty("U_Z_POSITION", oGrid.DataTable.GetValue("U_Z_POSITION", intRow))
                            Dim st As String = oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_FROMDATE", oGrid.DataTable.GetValue("U_Z_FROMDATE", intRow))
                            End If
                            st = oGrid.DataTable.GetValue("U_Z_TODATE", intRow)
                            If st <> "" Then
                                oChild.SetProperty("U_Z_TODATE", oGrid.DataTable.GetValue("U_Z_TODATE", intRow))
                            End If

                            oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_QUANTITY", intRow))
                            oChild.SetProperty("U_Z_MEASURE", oGrid.DataTable.GetValue("U_Z_MEASURE", intRow))
                            oChild.SetProperty("U_Z_DAYS", oGrid.DataTable.GetValue("U_Z_DAYS", intRow))
                            oChild.SetProperty("U_Z_HOURS", oGrid.DataTable.GetValue("U_Z_HOURS", intRow))
                            oChild.SetProperty("U_Z_AMOUNT", oGrid.DataTable.GetValue("U_Z_AMOUNT", intRow))
                            oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                            If oCheckBoxCOlumn.IsChecked(intRow) Then
                                oChild.SetProperty("U_Z_ORDER", "Y")
                            Else
                                oChild.SetProperty("U_Z_ORDER", "N")
                            End If

                            oChild.SetProperty("U_Z_ORDENTRY", oGrid.DataTable.GetValue("U_Z_ORDENTRY", intRow))
                            oChild.SetProperty("U_Z_ORDNUM", oGrid.DataTable.GetValue("U_Z_ORDNUM", intRow))
                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")

                            Try
                                oChild.SetProperty("U_Z_STATUS", oCombocolumn.GetSelectedValue(intRow).Value)
                            Catch ex As Exception
                                oChild.SetProperty("U_Z_STATUS", "I")
                            End Try
                            Try
                                oChild.SetProperty("U_Z_CMPDATE", oGrid.DataTable.GetValue("U_Z_CMPDATE", intRow))
                            Catch ex As Exception

                            End Try
                            oChild.SetProperty("U_Z_EXPTYPE", oGrid.DataTable.GetValue("U_Z_EXPTYPE", intRow))
                            oChild.SetProperty("U_Z_BOQ", oGrid.DataTable.GetValue("U_Z_BOQ", intRow))
                        End If
                    End If
                Next

                ochildern1 = oGeneralData.Child("Z_PRJ3")
                oMatrix = aform.Items.Item("41").Specific
                Dim oDel As SAPbobsCOM.Recordset
                oDel = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oDel.DoQuery("Delete from ""@Z_PRJ3"" where ""DocEntry""= " & aDocEntry)

                For intRow As Integer = 1 To oMatrix.VisualRowCount
                    If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow) <> "" Then
                        strLineId = oApplication.Utilities.getMatrixValues(oMatrix, "LineID", intRow)
                        Try
                            oChild = ochildern1.Item(CInt(strLineId) - 1)
                        Catch ex As Exception
                            oChild = ochildern1.Add()
                        End Try
                        oChild.SetProperty("U_Z_FILEPATH", oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow))
                        oChild.SetProperty("U_Z_FILENAME", oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow))
                    End If
                Next
                oGeneralService.Update(oGeneralData)
                CopyAttachment(aform)
            End If
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                LoadGridValues(aform)
                aform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            Else
                LoadGridValues(aform)
                aform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            End If
            aform.Freeze(False)
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End Try
    End Function

#Region "Copy attachment to Attachment Folder"
    Private Sub CopyAttachment(ByVal aForm As SAPbouiCOM.Form)
        oMatrix = aForm.Items.Item("41").Specific
        For i As Integer = 1 To oMatrix.VisualRowCount
            '  Dim odr As DataRow = oDT.Rows(i)
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select AttachPath From OADP"
            oRec.DoQuery(strQry)
            Dim SPath As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", i)
            If SPath = "" Then
            Else
                Dim DPath As String = ""
                If Not oRec.EoF Then
                    DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                End If
                If Not Directory.Exists(DPath) Then
                    Directory.CreateDirectory(DPath)
                End If
                Dim file = New FileInfo(SPath)
                Dim Filename As String = Path.GetFileName(SPath)
                Dim SavePath As String = Path.Combine(DPath, Filename)
                If System.IO.File.Exists(SPath) Then


                    If System.IO.File.Exists(SavePath) Then
                    Else
                        file.CopyTo(Path.Combine(DPath, file.Name), True)
                    End If
                End If
            End If
        Next
    End Sub
    Private Sub LoadAttachment(ByVal aForm As SAPbouiCOM.Form)
        oMatrix = aForm.Items.Item("41").Specific
        For i As Integer = 1 To oMatrix.VisualRowCount
            '  Dim odr As DataRow = oDT.Rows(i)
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select AttachPath From OADP"
            oRec.DoQuery(strQry)
            If oMatrix.IsRowSelected(i) Then
                Dim strFilename As String
                strFilename = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific.value
                Dim SPath As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", i)
                Dim DPath As String = ""
                If Not oRec.EoF Then
                    DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                End If
                If Not Directory.Exists(DPath) Then
                    Directory.CreateDirectory(DPath)
                End If
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                Dim Filename As String = Path.GetFileName(SPath)
                Dim SavePath As String = Path.Combine(DPath, Filename)
                If System.IO.File.Exists(strFilename) Then
                    strFilename = strFilename
                Else
                    strFilename = SavePath
                End If
                x.FileName = strFilename
                If System.IO.File.Exists(strFilename) Then
                    System.Diagnostics.Process.Start(x)
                Else
                    oApplication.Utilities.Message("File does not exits", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If

                x = Nothing
                Exit Sub
            End If
        Next
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                Dim STRcODE As String
                Dim oTest As SAPbobsCOM.Recordset
                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTest.DoQuery("select * from ONNM where Objectcode='Z_PRJ'")
                STRcODE = oTest.Fields.Item("AutoKey").Value
                Dim intDocEntry = CInt(STRcODE)
                intDocEntry = intDocEntry - 1
                'oApplication.Company.GetNewObjectC
                Dim strProjectCode As String
                oGrid = oForm.Items.Item("26").Specific
                Dim strDocEntry, strLineId, firstName, LastName As String
                Dim oRec As SAPbobsCOM.Recordset
                Dim oChild As SAPbobsCOM.GeneralData
                Dim oChildren As SAPbobsCOM.GeneralDataCollection
                Dim oGeneralService As SAPbobsCOM.GeneralService
                Dim oGeneralData As SAPbobsCOM.GeneralData
                Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                Dim oCompanyService As SAPbobsCOM.CompanyService
                oCompanyService = oApplication.Company.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("Z_PRJ")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)

                'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                LoadGridValues(oForm)
                oForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oForm.Items.Item("30").Visible = False
                oForm.Items.Item("31").Visible = False
                oForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                'End If

            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                Dim strProjectCode As String
                oGrid = oForm.Items.Item("26").Specific
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strDocEntry, strProject As String
        Dim otest As SAPbobsCOM.Recordset
        strProject = oApplication.Utilities.getEdittextvalue(aForm, "4") ' oCombobox.Selected.Value
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select isnull(U_Z_CARDCODE,'') from OPRJ where prjCode='" & strProject & "'")
        strProject = otest.Fields.Item(0).Value
        If strProject <> "" Then
            otest.DoQuery("Select * from OCRD where CardCode='" & strProject & "'")
            oApplication.Utilities.setEdittextvalue(aForm, "35", otest.Fields.Item("CardCode").Value)
            oApplication.Utilities.setEdittextvalue(aForm, "37", otest.Fields.Item("CardName").Value)
        Else
            oApplication.Utilities.setEdittextvalue(aForm, "35", " ")
            oApplication.Utilities.setEdittextvalue(aForm, "37", " ")
        End If
        If validation(aForm) = False Then
            Return False
        End If
        oCombobox = aForm.Items.Item("401").Specific
        strProject = oApplication.Utilities.getEdittextvalue(aForm, "4") ' oCombobox.Selected.Value
        '   Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select isnull(U_Z_CARDCODE,'') from OPRJ where prjCode='" & strProject & "'")
        strProject = otest.Fields.Item(0).Value
        If strProject <> "" Then
            otest.DoQuery("Select * from OCRD where CardCode='" & strProject & "'")
            oApplication.Utilities.setEdittextvalue(aForm, "35", otest.Fields.Item("CardCode").Value)
            oApplication.Utilities.setEdittextvalue(aForm, "37", otest.Fields.Item("CardName").Value)
        Else
            oApplication.Utilities.setEdittextvalue(aForm, "35", " ")
            oApplication.Utilities.setEdittextvalue(aForm, "37", " ")
        End If
        AssignLineNo(aForm)
        Return True
    End Function
    Private Sub PopulateProjectDetails(ByVal aform As SAPbouiCOM.Form)
        Dim strProject As String
        oForm = aform
        oForm.Freeze(True)
        oCombobox = oForm.Items.Item("401").Specific
        strProject = oApplication.Utilities.getEdittextvalue(aform, "4")
        oCombobox.Select(strProject, SAPbouiCOM.BoSearchKey.psk_ByValue)
        oApplication.Utilities.setEdittextvalue(oForm, "6", oCombobox.Selected.Description)
        strProject = oCombobox.Selected.Value
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select isnull(U_Z_CARDCODE,'') from OPRJ where prjCode='" & strProject & "'")
        strProject = otest.Fields.Item(0).Value
        If strProject <> "" Then
            otest.DoQuery("Select * from OCRD where CardCode='" & strProject & "'")
            oApplication.Utilities.setEdittextvalue(oForm, "35", otest.Fields.Item("CardCode").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "37", otest.Fields.Item("CardName").Value)
        Else
            oApplication.Utilities.setEdittextvalue(oForm, "35", " ")
            oApplication.Utilities.setEdittextvalue(oForm, "37", " ")
        End If

        ' oCombobox = oForm.Items.Item("4").Specific
        oApplication.Utilities.setEdittextvalue(oForm, "6", oCombobox.Selected.Description)
        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery("Select * from [@Z_HPRJ] where U_Z_PRJCODE='" & oCombobox.Selected.Value & "'")
            If otest.RecordCount > 0 Then
                oForm.Freeze(True)
                oApplication.Utilities.Message("Budget Details already defined for selected project", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oApplication.Utilities.setEdittextvalue(oForm, "19", otest.Fields.Item("U_Z_FROMDATE").Value)
                oApplication.Utilities.setEdittextvalue(oForm, "21", otest.Fields.Item("U_Z_TODATE").Value)
                oCombobox = oForm.Items.Item("17").Specific
                oCombobox.Select(otest.Fields.Item("U_Z_STATUS").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oForm.Items.Item("33").Specific
                oCombobox.Select(otest.Fields.Item("U_Z_APPROVAL").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                oApplication.Utilities.setEdittextvalue(oForm, "8", otest.Fields.Item("U_Z_BUDGET").Value)
                oApplication.Utilities.setEdittextvalue(oForm, "23", otest.Fields.Item("U_Z_TOTALEXPENSE").Value)
                oForm.Freeze(False)
                LoadGridValues(oForm)
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            Else
                LoadGridValues(oForm)
            End If
            ' LoadGridValues(oForm)
        End If
        AssignLineNo(aform)
        oForm.Freeze(False)
    End Sub

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
#Region "FileOpen"
    Private Sub FileOpen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()
    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog

        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.SafeFileName
                        strFilepath = oDialogBox.FileName

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

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BudgetEntry Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_ORDENTRY" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                                    If oCheckBoxCOlumn.IsChecked(pVal.Row) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    AddChooseFromList_Conditions_SalesOrder(oForm, pVal.Row)
                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_CNTID" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oGrid.DataTable.GetValue("U_Z_EMPID", pVal.Row) <> "" Then
                                        oApplication.Utilities.Message("You can select either Employee ID / Contract ID", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_EMPID" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oGrid.DataTable.GetValue("U_Z_CNTID", pVal.Row) <> "" Then
                                        oApplication.Utilities.Message("You can select either Employee ID / Contract ID", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_CNTID" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oGrid.DataTable.GetValue("U_Z_EMPID", pVal.Row) <> "" Then
                                        oApplication.Utilities.Message("You can select either Employee ID / Contract ID", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_EMPID" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oGrid.DataTable.GetValue("U_Z_CNTID", pVal.Row) <> "" Then
                                        oApplication.Utilities.Message("You can select either Employee ID / Contract ID", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_ORDENTRY" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                                    If oCheckBoxCOlumn.IsChecked(pVal.Row) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    AddChooseFromList_Conditions_SalesOrder(oForm, pVal.Row)
                                End If
                                If pVal.ItemUID = "12" And pVal.ColUID = "BOQ" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_BOQ" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oCombobox = oForm.Items.Item("17").Specific
                                If oCombobox.Selected.Value = "C" Or oCombobox.Selected.Value = "H" Then
                                    If (pVal.ItemUID = "12" Or pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") Then
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

                                'If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_BOQ" Then
                                '    BubbleEvent = False
                                '    Exit Sub
                                'End If

                                'If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_BOQ" Then
                                '    BubbleEvent = False
                                '    Exit Sub
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    If AddUDT(oForm, 0) Then

                                    End If
                                    BubbleEvent = False
                                    Exit Sub

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                ' oMatrix = oForm.Items.Item("12").Specific
                                If pVal.ItemUID = "12" And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "12"
                                    frmSourceMatrix = oMatrix
                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "12"
                                    frmSourceGrid = oGrid
                                    'If pVal.ColUID = "U_Z_ACTNAME" Then
                                    '    AddChooseFromList_Conditions(oForm, pVal.Row)

                                    'End If
                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_CNTID" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oGrid.DataTable.GetValue("U_Z_EMPID", pVal.Row) <> "" Then
                                        oApplication.Utilities.Message("You can select either Employee ID / Contract ID", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_EMPID" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oGrid.DataTable.GetValue("U_Z_CNTID", pVal.Row) <> "" Then
                                        oApplication.Utilities.Message("You can select either Employee ID / Contract ID", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "12" And pVal.ColUID = "V_6" Then
                                    oCheckBox = oMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific
                                    If oCheckBox.Checked = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_ORDENTRY" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oCheckBoxCOlumn = oGrid.Columns.Item("U_Z_ORDER")
                                    If oCheckBoxCOlumn.IsChecked(pVal.Row) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ColUID = "U_Z_CNTID" Then
                                    Dim obj As New clsContractAgreement
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    obj.ViewForm(oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row))
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "401" Then
                                    Dim strProject As String
                                    strProject = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    oApplication.Utilities.setEdittextvalue(oForm, "6", oCombobox.Selected.Description)
                                    strProject = oCombobox.Selected.Value
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Select isnull(U_Z_CARDCODE,'') from OPRJ where prjCode='" & strProject & "'")
                                    strProject = otest.Fields.Item(0).Value
                                    If strProject <> "" Then
                                        otest.DoQuery("Select * from OCRD where CardCode='" & strProject & "'")
                                        oApplication.Utilities.setEdittextvalue(oForm, "35", otest.Fields.Item("CardCode").Value)
                                        oApplication.Utilities.setEdittextvalue(oForm, "37", otest.Fields.Item("CardName").Value)
                                    Else
                                        oApplication.Utilities.setEdittextvalue(oForm, "35", " ")
                                        oApplication.Utilities.setEdittextvalue(oForm, "37", " ")
                                    End If
                                    oApplication.Utilities.setEdittextvalue(oForm, "6", oApplication.Utilities.getEdittextvalue(oForm, "4"))
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        otest.DoQuery("Select * from [@Z_HPRJ] where U_Z_PRJCODE='" & oCombobox.Selected.Value & "'")
                                        If otest.RecordCount > 0 Then
                                            oForm.Freeze(True)
                                            oApplication.Utilities.Message("Budget Details already defined for selected project", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oApplication.Utilities.setEdittextvalue(oForm, "19", otest.Fields.Item("U_Z_FROMDATE").Value)
                                            oApplication.Utilities.setEdittextvalue(oForm, "21", otest.Fields.Item("U_Z_TODATE").Value)
                                            oCombobox = oForm.Items.Item("17").Specific
                                            oCombobox.Select(otest.Fields.Item("U_Z_STATUS").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            oCombobox = oForm.Items.Item("33").Specific
                                            oCombobox.Select(otest.Fields.Item("U_Z_APPROVAL").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                            oApplication.Utilities.setEdittextvalue(oForm, "8", otest.Fields.Item("U_Z_BUDGET").Value)
                                            oApplication.Utilities.setEdittextvalue(oForm, "23", otest.Fields.Item("U_Z_TOTALEXPENSE").Value)
                                            oForm.Freeze(False)
                                            LoadGridValues(oForm)
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        Else
                                            LoadGridValues(oForm)
                                        End If
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" And pVal.ColUID <> "V_-1" Then
                                    If 1 = 1 Then
                                        Dim objChooseForm As SAPbouiCOM.Form
                                        Dim objChoose As New ClsBOQ
                                        Dim strproject, strprojectname, strbusinessprocess, strActvity, strRef As String
                                        oMatrix = oForm.Items.Item("12").Specific
                                        oCombobox = oForm.Items.Item("401").Specific
                                        strproject = oCombobox.Selected.Value ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                        strprojectname = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                        strbusinessprocess = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                        strActvity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", pVal.Row)
                                        strRef = oApplication.Utilities.getMatrixValues(oMatrix, "BOQ", pVal.Row)
                                        If strproject = "" Then
                                            Exit Sub
                                        End If
                                        oCombobox = oMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                        If oCombobox.Selected.Value <> "I" Then
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

                                '   If (pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID <> "V_-1" Then
                                If (pVal.ItemUID = "27") And pVal.ColUID <> "V_-1" Then
                                    If 1 = 1 Then
                                        Dim objChooseForm As SAPbouiCOM.Form
                                        Dim objChoose As New ClsBOQ
                                        Dim strproject, strprojectname, strbusinessprocess, strActvity, strRef As String
                                        oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                        oCombobox = oForm.Items.Item("401").Specific
                                        strproject = oCombobox.Selected.Value ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                        strprojectname = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                        strbusinessprocess = oGrid.DataTable.GetValue("U_Z_MODNAME", pVal.Row)
                                        strActvity = oGrid.DataTable.GetValue("U_Z_ACTNAME", pVal.Row)
                                        strRef = oGrid.DataTable.GetValue("U_Z_BOQ", pVal.Row)
                                        If strproject = "" Then
                                            Exit Sub
                                        End If

                                        'oCombobox = oMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                        'If oCombobox.Selected.Value = "R" Then
                                        '    Exit Sub
                                        'End If
                                        If strproject <> "" Then
                                            objChoose.ItemUID = pVal.ItemUID
                                            objChoose.SourceFormUID = FormUID
                                            objChoose.SourceLabel = pVal.Row
                                            objChoose.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                            objChoose.choice = "MODULE"
                                            objChoose.prjcode = strproject
                                            objChoose.prjname = strprojectname
                                            oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                                            Try
                                                objChoose.stats = oCombocolumn.Selected.Description
                                            Catch ex As Exception
                                                objChoose.stats = "In Process"
                                            End Try

                                            'oCombocolumn = oMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific
                                            'objChoose.stats = oCombobox.Selected.Description

                                            objChoose.boqref = strRef
                                            objChoose.businessprocess = strbusinessprocess
                                            objChoose.busienssactivity = strActvity
                                            objChoose.sourcerowId = pVal.Row
                                            objChoose.BinDescrUID = "Grid"
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
                                    Case "24"
                                        oForm.PaneLevel = 2
                                    Case "25"
                                        oForm.PaneLevel = 3
                                    Case "29"
                                        oForm.PaneLevel = 11
                                    Case "40"
                                        oForm.PaneLevel = 5
                                    Case "43"

                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "42"
                                        oMatrix = oForm.Items.Item("41").Specific
                                        ' Dim strPath As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)

                                        FileOpen()
                                        AddRow(oForm)
                                        If strFilepath = "" Then
                                            '  oApplication.Utilities.Message("Please Select a File", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            ' BubbleEvent = False
                                        Else
                                            '    oMatrix.AddRow()
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, strFilepath)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, strMdbFilePath)
                                            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            oMatrix = oForm.Items.Item("41").Specific
                                            For intRow As Integer = 1 To oMatrix.VisualRowCount
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_-1", intRow, intRow)

                                            Next

                                        End If
                                    Case "44"
                                        LoadAttachment(oForm)
                                    Case "1"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            LoadGridValues(oForm)
                                        End If
                                        'Case "12"
                                        '    oMatrix = oForm.Items.Item("12").Specific
                                        '    If pVal.ColUID = "V_5" Then
                                        '        oCheckBox = oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific
                                        '        If oCheckBox.Checked = False Then
                                        '            oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", pVal.Row, "")
                                        '            oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", pVal.Row, "")
                                        '        End If
                                        '    End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_DAYS" And pVal.CharPressed = 9 Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strHours As String
                                    Dim dblHours, dblRate As Double
                                    strHours = oGrid.DataTable.GetValue("U_Z_DAYS", pVal.Row)
                                    If strHours <> "" Then
                                        dblHours = oApplication.Utilities.getDocumentQuantity(strHours)
                                        If (dblHours > 0) Then
                                            dblHours = dblHours * 8
                                            Dim oTest As SAPbobsCOM.Recordset
                                            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            'oEditText = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                            If oGrid.DataTable.GetValue("U_Z_EMPID", pVal.Row) <> "" Then
                                                oTest.DoQuery("Select isnull(U_Daily_rate,0) from OHEM where empID=" & oGrid.DataTable.GetValue("U_Z_EMPID", pVal.Row))
                                                dblRate = oTest.Fields.Item(0).Value
                                                dblRate = dblRate * oApplication.Utilities.getDocumentQuantity(strHours)
                                                oGrid.DataTable.SetValue("U_Z_AMOUNT", pVal.Row, dblRate)
                                                oGrid.DataTable.SetValue("U_Z_HOURS", pVal.Row, dblHours)
                                            End If
                                            oGrid.Columns.Item("U_Z_AMOUNT").Click(pVal.Row, False, 0)
                                        End If
                                    End If
                                End If

                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_HOURS" And pVal.CharPressed = 9 Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strHours As String
                                    Dim dblHours, dblRate, dblDays As Double
                                    strHours = oGrid.DataTable.GetValue("U_Z_HOURS", pVal.Row)
                                    If strHours <> "" Then
                                        dblHours = oApplication.Utilities.getDocumentQuantity(strHours)
                                        If (dblHours > 0) Then
                                            dblHours = dblHours '* 8
                                            dblDays = dblHours / 8
                                            Dim oTest As SAPbobsCOM.Recordset
                                            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            If oGrid.DataTable.GetValue("U_Z_EMPID", pVal.Row) <> "" Then
                                                oTest.DoQuery("Select isnull(U_HR_rate,0) from OHEM where empID=" & oGrid.DataTable.GetValue("U_Z_EMPID", pVal.Row))
                                                dblRate = oTest.Fields.Item(0).Value
                                                dblRate = dblRate * oApplication.Utilities.getDocumentQuantity(strHours)
                                                oGrid.DataTable.SetValue("U_Z_DAYS", pVal.Row, dblDays)
                                                ' oGrid.DataTable.SetValue("U_Z_HOURS", pVal.Row, dblHours)
                                                oGrid.DataTable.SetValue("U_Z_AMOUNT", pVal.Row, dblRate)

                                            End If
                                            oGrid.Columns.Item("U_Z_AMOUNT").Click(pVal.Row, False, 0)
                                        End If
                                    End If
                                End If



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


                                If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_ACTNAME" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsCFLActivity
                                    Dim strproject, strprojectname, strbusinessprocess, strActvity, strRef As String
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    strbusinessprocess = oGrid.DataTable.GetValue("U_Z_MODNAME", pVal.Row)
                                    strActvity = oGrid.DataTable.GetValue("U_Z_ACTNAME", pVal.Row)
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If strActvity <> "" Then
                                        otest.DoQuery("Select * from ""@Z_ACTIVITY""  where U_Z_ACTNAME='" & strActvity & "' and  (isnull(U_Z_MODNAME,'0')='0' or U_Z_MODNAME='" & strbusinessprocess & "')")
                                        If otest.RecordCount > 0 And strbusinessprocess <> "" Then
                                            Exit Sub
                                        Else
                                            oGrid.DataTable.SetValue("U_Z_ACTNAME", pVal.Row, "")
                                        End If
                                    End If
                                    If strbusinessprocess <> "" Then
                                        oGrid.DataTable.SetValue("U_Z_ACTNAME", pVal.Row, "")
                                        objChoose.ItemUID = pVal.ItemUID
                                        objChoose.SourceFormUID = FormUID
                                        objChoose.SourceLabel = pVal.Row
                                        objChoose.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        objChoose.choice = "ACTIVITY"
                                        objChoose.ItemCode = ""
                                        objChoose.sourcerowId = pVal.Row
                                        objChoose.Documentchoice = strbusinessprocess
                                        oApplication.Utilities.LoadForm("frm_ACT1.xml", "frm_ACT")
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
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


                                If (pVal.ItemUID = "27") And pVal.ColUID = "U_Z_BOQ" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New ClsBOQ
                                    Dim strproject, strprojectname, strbusinessprocess, strActvity, strRef As String
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oCombobox = oForm.Items.Item("4").Specific
                                    strproject = oCombobox.Selected.Value ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    strprojectname = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                    strbusinessprocess = oGrid.DataTable.GetValue("U_Z_MODNAME", pVal.Row)
                                    strActvity = oGrid.DataTable.GetValue("U_Z_ACTNAME", pVal.Row)
                                    strRef = oGrid.DataTable.GetValue("U_Z_BOQ", pVal.Row)
                                    If strproject = "" Then
                                        Exit Sub
                                    End If
                                    'oCombobox = oMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                                    'If oCombobox.Selected.Value = "R" Then
                                    '    Exit Sub
                                    'End If
                                    If strproject <> "" Then
                                        objChoose.ItemUID = pVal.ItemUID
                                        objChoose.SourceFormUID = FormUID
                                        objChoose.SourceLabel = pVal.Row
                                        objChoose.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        objChoose.choice = "MODULE"
                                        objChoose.prjcode = strproject
                                        objChoose.prjname = strprojectname
                                        oCombocolumn = oGrid.Columns.Item("U_Z_STATUS")
                                        Try
                                            objChoose.stats = oCombocolumn.Selected.Description
                                        Catch ex As Exception
                                            objChoose.stats = "In Process"
                                        End Try
                                        objChoose.boqref = strRef
                                        objChoose.businessprocess = strbusinessprocess
                                        objChoose.busienssactivity = strActvity
                                        objChoose.sourcerowId = pVal.Row
                                        objChoose.BinDescrUID = "Grid"
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
                                        'If pVal.ItemUID = "12" And pVal.ColUID = "V_0" Then
                                        '    val = oDataTable.GetValue("U_Z_MODNAME", 0)
                                        '    oMatrix = oForm.Items.Item("12").Specific
                                        '    oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val)
                                        'End If
                                        If pVal.ItemUID = "31" Then
                                            val = oDataTable.GetValue("U_Z_PRJCODE", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception

                                            End Try
                                            PopulateBudgetfromTemplate(oForm, val)
                                        End If
                                        If pVal.ItemUID = "4" Then
                                            val = oDataTable.GetValue("PrjCode", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "4", val)
                                            Catch ex As Exception

                                            End Try
                                            PopulateProjectDetails(oForm)
                                        End If
                                        If pVal.ItemUID = "28" And pVal.ColUID = "U_Z_EXPTYPE" Then
                                            val = oDataTable.GetValue("U_Z_EXPNAME", 0)
                                            oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        End If

                                        If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_MODNAME" Then
                                            val = oDataTable.GetValue("U_Z_MODNAME", 0)
                                            oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        End If
                                        'If pVal.ItemUID = "12" And pVal.ColUID = "V_3" Then
                                        '    val = oDataTable.GetValue("U_Z_ACTNAME", 0)
                                        '    oMatrix = oForm.Items.Item("12").Specific
                                        '    oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", pVal.Row, val)
                                        'End If

                                        If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_ACTNAME" Then
                                            val = oDataTable.GetValue("U_Z_ACTNAME", 0)
                                            oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        End If

                                        'If pVal.ItemUID = "12" And pVal.ColUID = "V_4" Then
                                        '    val = oDataTable.GetValue("empID", 0)
                                        '    val1 = oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0)
                                        '    oMatrix = oForm.Items.Item("12").Specific
                                        '    oApplication.Utilities.SetMatrixValues(oMatrix, "V_11", pVal.Row, val1)
                                        '    oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", pVal.Row, val)
                                        'End If
                                        If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_EMPID" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            val1 = oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0)
                                            oGrid = oForm.Items.Item(pVal.ItemUID).Specific

                                            oGrid.DataTable.SetValue("U_Z_POSITION", pVal.Row, val1)
                                            oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        End If

                                        If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_CNTID" Then
                                            val = oDataTable.GetValue("DocNum", 0)
                                            val1 = oDataTable.GetValue("U_Z_CARDNAME", 0)
                                            oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            oGrid.DataTable.SetValue("U_Z_POSITION", pVal.Row, val1)
                                            oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        End If

                                        'If pVal.ItemUID = "12" And pVal.ColUID = "V_6" Then
                                        '    val = oDataTable.GetValue("DocEntry", 0)
                                        '    val1 = oDataTable.GetValue("DocNum", 0)
                                        '    oMatrix = oForm.Items.Item("12").Specific
                                        '    oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", pVal.Row, val1)
                                        '    oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", pVal.Row, val)
                                        'End If

                                        If (pVal.ItemUID = "26" Or pVal.ItemUID = "27" Or pVal.ItemUID = "28") And pVal.ColUID = "U_Z_ORDENTRY" Then
                                            val = oDataTable.GetValue("DocEntry", 0)
                                            val1 = oDataTable.GetValue("DocNum", 0)
                                            oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            oGrid.DataTable.SetValue("U_Z_ORDNUM", pVal.Row, val)
                                            oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        End If

                                        If pVal.ItemUID = "140" Then
                                            val = oDataTable.GetValue("DocNum", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
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


Public Class WindowWrapper
    Implements System.Windows.Forms.IWin32Window
    Private _hwnd As IntPtr

    Public Sub New(ByVal handle As IntPtr)
        _hwnd = handle
    End Sub

    Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property

End Class



