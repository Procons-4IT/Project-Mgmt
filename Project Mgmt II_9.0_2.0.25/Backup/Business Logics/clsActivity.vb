Public Class clsActivity
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private objMatrix, oMatrix As SAPbouiCOM.Matrix
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

#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Activity) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Activity, frm_Activity)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        If oForm.TypeEx = frm_Activity Then
            'oForm.EnableMenu("1283", False)
            '  oForm.EnableMenu(mnu_Delete, True)
        End If
        oForm.Settings.Enabled = True
        AddChooseFromList(oForm)
        oForm.Freeze(True)
        BindData(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        oForm.Freeze(False)
    End Sub

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

            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Z_Status"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()

            'oCFLCreationParams.ObjectType = "Z_Module"
            'oCFLCreationParams.UniqueID = "CFL2"
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
#Region "DataBind"

    Public Sub BindData(ByVal objform As SAPbouiCOM.Form)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oDBDataSrc As SAPbouiCOM.DBDataSource
        Try
            oMatrix = objform.Items.Item("3").Specific
            oDBDataSrc = objform.DataSources.DBDataSources.Add("@Z_ACTIVITY")
            Try
                oDBDataSrc.Query()
            Catch ex As Exception

            End Try


            Dim oColum As SAPbouiCOM.Column
            oColum = oMatrix.Columns.Item("V_1")
            oColum.ValidValues.Add("Y", "Yes")
            oColum.ValidValues.Add("N", "No")
            oColum.DisplayDesc = True

            oMatrix = oForm.Items.Item("3").Specific
            oColum = oMatrix.Columns.Item("V_2")
            oColum.ChooseFromListUID = "CFL1"
            oColum.ChooseFromListAlias = "U_Z_MODNAME"

            oMatrix.LoadFromDataSource()
            If oMatrix.RowCount >= 1 Then
                If oMatrix.Columns.Item(1).Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                    oDBDataSrc.Clear()
                    oMatrix.AddRow()
                    oMatrix.Columns.Item(0).Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount
                    oCombobox = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                    oCombobox.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    oMatrix.Columns.Item(1).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            ElseIf oMatrix.RowCount = 0 Then
                oMatrix.AddRow()
                oMatrix.Columns.Item(0).Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount
                oCombobox = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                oCombobox.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                oMatrix.Columns.Item(1).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#Region "Enable Matrix After Update"
    '***************************************************************************
    'Type               : Procedure
    'Name               : EnblMatrixAfterUpdate
    'Parameter          : Application,Company,Form
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Enable the Matrix after update button is pressed.
    '***************************************************************************
    Private Sub EnblMatrixAfterUpdate(ByVal objApplication As SAPbouiCOM.Application, ByVal ocompany As SAPbobsCOM.Company, ByVal oForm As SAPbouiCOM.Form)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oDBDSource As SAPbouiCOM.DBDataSource
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim lnErrCode As Long
        Dim strErrMsg As String
        Dim i As Integer
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oForm.Freeze(True)
            If oMatrix.Columns.Item(1).Cells.Item(oMatrix.RowCount).Specific.value = "" Then
                oUserTable = ocompany.UserTables.Item("Z_ACTIVITY")
                oDBDSource = oForm.DataSources.DBDataSources.Item("@Z_ACTIVITY")
                oMatrix.DeleteRow(oMatrix.RowCount)
                oMatrix.FlushToDataSource()
                For i = 0 To oDBDSource.Size - 1
                    If oUserTable.GetByKey(oDBDSource.GetValue("Code", i)) Then
                        oUserTable.Name = oDBDSource.GetValue("Name", i)
                        oUserTable.Code = oDBDSource.GetValue("Code", i)
                        oUserTable.UserFields.Fields.Item("U_Z_Actname").Value = oDBDSource.GetValue("U_Z_Actname", i)
                        oUserTable.UserFields.Fields.Item("U_Z_ModName").Value = oDBDSource.GetValue("U_Z_ModName", i)
                        oUserTable.UserFields.Fields.Item("U_Z_Status").Value = oDBDSource.GetValue("U_Z_Status", i)
                        If oUserTable.Update <> 0 Then
                            MsgBox(ocompany.GetLastErrorDescription)
                        End If
                    End If
                Next
                oDBDSource.Query()
                oMatrix.Columns.Item(1).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If
            oForm.Freeze(False)
            Exit Sub
        Catch ex As Exception
            ocompany.GetLastError(lnErrCode, strErrMsg)
            If strErrMsg <> "" Then
                objApplication.MessageBox(strErrMsg)
            Else
                objApplication.MessageBox(ex.Message)
            End If
        End Try
    End Sub
#End Region

#Region "Insert Code and Doc Entry"
    '******************************************************************
    'Type               : Procedure
    'Name               : InsertCodeAndDocEntry
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Inserting code and docEntry values.
    '******************************************************************
    Public Sub InsertCodeAndDocEntry(ByVal aForm As SAPbouiCOM.Form)
        Dim oDBDSource As SAPbouiCOM.DBDataSource
        Dim strValue As String = "1"
        Try
            objForm = aForm
            aForm.Freeze(True)
            oDBDSource = objForm.DataSources.DBDataSources.Item("@Z_ACTIVITY")
            objMatrix = objForm.Items.Item("3").Specific
            objMatrix.FlushToDataSource()
            If objMatrix.RowCount = 1 Then
                oDBDSource.SetValue("Code", 0, strValue.PadLeft(8, "0"))
                oDBDSource.SetValue("DocEntry", 0, strValue.PadLeft(8, "0"))
            Else
                oDBDSource.SetValue("Code", objMatrix.RowCount - 1, oDBDSource.GetValue("DocEntry", objMatrix.RowCount - 1).PadLeft(8, "0"))
                oDBDSource.SetValue("DocEntry", objMatrix.RowCount - 1, oDBDSource.GetValue("DocEntry", objMatrix.RowCount - 1).PadLeft(8, "0"))
            End If
            objMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region





#End Region

#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Activity
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
            If pVal.FormTypeEx = frm_Activity Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_2") And pVal.CharPressed <> 9 Then
                                    objMatrix = oForm.Items.Item("3").Specific
                                    Dim strValue, strPhase As String
                                    strValue = oApplication.Utilities.getMatrixValues(objMatrix, "V_0", pVal.Row)
                                    strPhase = oApplication.Utilities.getMatrixValues(objMatrix, "V_2", pVal.Row)
                                    If oApplication.Utilities.getMatrixValues(objMatrix, "V_2", pVal.Row) <> "" Then

                                        If oApplication.Utilities.validaterecordexists_Activity(strValue, "U_Z_ACTNAME", "Z_PRJ1", strPhase) = True Then
                                            oApplication.Utilities.Message("Activity already mapped to Projects. You can not change the Activity Name", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        If oApplication.Utilities.validaterecordexists_Activity(strValue, "U_Z_ACTNAME", "Z_DPRJ1", strPhase) = True Then
                                            oApplication.Utilities.Message("Activity already mapped to Projects Budget Template. You can not change the Activity Name", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_2") And pVal.CharPressed <> 9 Then
                                    objMatrix = oForm.Items.Item("3").Specific
                                    Dim strValue, strPhase As String
                                    strValue = oApplication.Utilities.getMatrixValues(objMatrix, "V_0", pVal.Row)
                                    strPhase = oApplication.Utilities.getMatrixValues(objMatrix, "V_2", pVal.Row)
                                    'If oApplication.Utilities.validaterecordexists(strValue, "U_Z_ACTNAME", "Z_PRJ2") = True Then
                                    If strValue <> "" And strPhase <> "" Then
                                        'If oApplication.Utilities.getMatrixValues(objMatrix, "V_2", pVal.Row) <> "" Then
                                        '    If oApplication.Utilities.validaterecordexists_Activity(strValue, "U_Z_ACTNAME", "Z_PRJ1", strPhase) = True Then
                                        '        oApplication.Utilities.Message("Activity already mapped to Projects. You can not change the Activity Name", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        '        BubbleEvent = False
                                        '        Exit Sub
                                        '    End If

                                        '    If oApplication.Utilities.validaterecordexists_Activity(strValue, "U_Z_ACTNAME", "Z_DPRJ1", strPhase) = True Then
                                        '        oApplication.Utilities.Message("Activity already mapped to Projects Budget Template. You can not change the Activity Name", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        '        BubbleEvent = False
                                        '        Exit Sub
                                        '    End If
                                        'End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    oForm.Freeze(True)
                                    InsertCodeAndDocEntry(oForm)
                                    EnblMatrixAfterUpdate(oApplication.SBO_Application, oApplication.Company, oForm)
                                    oForm.Freeze(False)
                                End If
                        End Select
                    Case False
                        If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.ItemUID = "1")) Then
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            oForm.Freeze(True)
                            objForm = oForm
                            objMatrix = objForm.Items.Item("3").Specific
                            objMatrix.AddRow()
                            objMatrix.Columns.Item(0).Cells.Item(objMatrix.RowCount).Specific.value = objMatrix.RowCount
                            objMatrix.Columns.Item(1).Cells.Item(objMatrix.RowCount).Specific.value = ""
                            objMatrix.Columns.Item("V_2").Cells.Item(objMatrix.RowCount).Specific.value = ""
                            oCombobox = objMatrix.Columns.Item("V_1").Cells.Item(objMatrix.RowCount).Specific
                            oCombobox.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            objMatrix.Columns.Item(1).Cells.Item(objMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            oForm.Freeze(False)
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed = "9" And pVal.ItemUID = "3" And pVal.ColUID = "V_0" Then
                            Dim strVal As String
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            objForm = oForm
                            objMatrix = objForm.Items.Item("3").Specific
                            'For i As Integer = 1 To objMatrix.RowCount - 1
                            '    If i <> pVal.Row Then
                            '        If objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value = objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value Then
                            '            oApplication.Utilities.Message("Activity already exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '            objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            '            BubbleEvent = False
                            '            Exit Sub
                            '        End If
                            '    End If
                            'Next
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
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
                                    If pVal.ItemUID = "3" And pVal.ColUID = "V_2" Then
                                        val = oDataTable.GetValue("U_Z_MODNAME", 0)
                                        oMatrix = oForm.Items.Item("3").Specific
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", pVal.Row, val)
                                    End If

                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                oForm.Freeze(False)
                            End Try
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

#End Region
End Class
