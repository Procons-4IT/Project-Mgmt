Public Class ClsBOQ
    Inherits clsBase

#Region "Declarations"
    Public Shared ItemUID As String
    Public Shared SourceFormUID As String
    Public Shared SourceLabel As Integer
    Public Shared CFLChoice As String
    Public Shared ItemCode As String
    Public Shared choice As String
    Public Shared sourceColumID As String
    Public Shared sourcerowId As Integer
    Public Shared BinDescrUID As String
    Public Shared Documentchoice As String
    Public Shared prjcode As String
    Public Shared prjname As String
    Public Shared businessprocess As String
    Public Shared busienssactivity As String
    Public Shared boqref As String
    Public Shared stats As String

    Private oDbDatasource As SAPbouiCOM.DBDataSource
    Private Ouserdatasource As SAPbouiCOM.UserDataSource
    Private oConditions As SAPbouiCOM.Conditions
    Private ocondition As SAPbouiCOM.Condition
    Private intRowId As Integer
    Private strRowNum As Integer
    Private i As Integer
    Private oedit As SAPbouiCOM.EditText
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboBox As SAPbouiCOM.ComboBoxColumn
    '   Private oForm As SAPbouiCOM.Form
    Private objSoureceForm As SAPbouiCOM.Form
    Private objform As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Grid
    Private osourcegrid As SAPbouiCOM.Matrix
    Private Const SEPRATOR As String = "~~~"
    Private SelectedRow As Integer
    Private sSearchColumn As String
    Private oItem As SAPbouiCOM.Item
    Public stritemid As SAPbouiCOM.Item
    Private intformmode As SAPbouiCOM.BoFormMode
    Private objGrid As SAPbouiCOM.Grid
    Private objSourcematrix As SAPbouiCOM.Matrix
    Private dtTemp As SAPbouiCOM.DataTable
    Private objStatic As SAPbouiCOM.StaticText
    Private inttable As Integer = 0
    Public strformid As String
    Public strStaticValue As String
    Public strSQL As String
    Private strSelectedItem1 As String
    Private strSelectedItem2 As String
    Private strSelectedItem3 As String
    Private strSelectedItem4 As String
    Private oRecSet As SAPbobsCOM.Recordset
    '   Private objSBOAPI As ClsSBO
    '   Dim objTransfer As clsTransfer
#End Region

#Region "New"
    '*****************************************************************
    'Type               : Constructor
    'Name               : New
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create object for classes.
    '******************************************************************
    Public Sub New()
        '   objSBOAPI = New ClsSBO
        MyBase.New()
    End Sub
#End Region

#Region "Bind Data"
    '******************************************************************
    'Type               : Procedure
    'Name               : BindData
    'Parameter          : Form
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Binding the fields.
    '******************************************************************
    Public Sub databound(ByVal objform As SAPbouiCOM.Form)
        Try
            Dim strSQL As String = ""
            Dim stCode As String
            Dim ObjSegRecSet As SAPbobsCOM.Recordset
            ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objform.Freeze(True)
            objform.DataSources.DataTables.Add("dtLevel3")
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 250)
            oedit = objform.Items.Item("etFind").Specific
            oedit.DataBind.SetBound(True, "", "dbFind")
            objGrid = objform.Items.Item("mtchoose").Specific
            dtTemp = objform.DataSources.DataTables.Item("dtLevel3")
            If choice = "MODULE" Then
                objform.Title = "Project BoQ Details"
                stCode = oApplication.Utilities.getMaxCode("@Z_PRJ2", "U_Z_BOQRef")

                If boqref <> "" Then
                    strSQL = "Select * from [@Z_PRJ2] where U_Z_BOQRef='" & boqref & "'"
                Else
                    boqref = stCode
                    strSQL = "Select * from [@Z_PRJ2] where U_Z_BOQRef='" & boqref & "'"
                End If
                dtTemp.ExecuteQuery(strSQL)
                objGrid.DataTable = dtTemp
            End If
            oForm = objform
            oApplication.Utilities.setEdittextvalue(oForm, "etFind", prjcode)
            oApplication.Utilities.setEdittextvalue(oForm, "7", prjname)
            oApplication.Utilities.setEdittextvalue(oForm, "9", businessprocess)
            oApplication.Utilities.setEdittextvalue(oForm, "11", busienssactivity)
            oApplication.Utilities.setEdittextvalue(oForm, "15", boqref)
            Dim otest As SAPbobsCOM.Recordset
            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery("Select  * from ""@Z_PRJ1"" where ""U_Z_BOQ""='" & boqref & "'")

            oApplication.Utilities.setEdittextvalue(oForm, "17", stats)
            If otest.Fields.Item("U_Z_CNTID").Value = "" Then
                oApplication.Utilities.setEdittextvalue(oForm, "25", "")
                oApplication.Utilities.setEdittextvalue(oForm, "27", "")
                oApplication.Utilities.setEdittextvalue(oForm, "29", "")
            Else
                oApplication.Utilities.setEdittextvalue(oForm, "25", otest.Fields.Item("U_Z_CNTID").Value)
                oApplication.Utilities.setEdittextvalue(oForm, "27", otest.Fields.Item("U_Z_POSITION").Value)
                oApplication.Utilities.setEdittextvalue(oForm, "29", otest.Fields.Item("U_Z_CUSTCNTID").Value)
            End If
            AddChooseFromList(objform)
            oForm.Items.Item("mtchoose").Enabled = True
            FormatGrid(objGrid)
            objGrid.AutoResizeColumns()
            objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            If objGrid.Rows.Count > 0 Then
                objGrid.Rows.SelectedRows.Add(0)
            End If
            objform.Freeze(False)
            objform.Update()
            
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub

    Public Sub PopulateBoQDetails(ByVal objform As SAPbouiCOM.Form)
        Try
            Dim strSQL As String = ""
            Dim stCode As String
            Dim ObjSegRecSet As SAPbobsCOM.Recordset
            ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objform.Freeze(True)
            ' objform.DataSources.DataTables.Add("dtLevel3")
            objGrid = objform.Items.Item("mtchoose").Specific
            dtTemp = objform.DataSources.DataTables.Item("dtLevel3")
            If choice = "MODULE" Then
                objform.Title = "Project BoQ Details"
                stCode = oApplication.Utilities.getMaxCode("@Z_PRJ2", "U_Z_BOQRef")
                If boqref <> "" Then
                    strSQL = "Select * from [@Z_PRJ2] where U_Z_BOQRef='" & boqref & "'"
                Else
                    boqref = stCode
                    strSQL = "Select * from [@Z_PRJ2] where U_Z_BOQRef='" & boqref & "'"
                End If
                dtTemp.ExecuteQuery(strSQL)
                objGrid.DataTable = dtTemp
            End If
            oForm = objform
            oApplication.Utilities.setEdittextvalue(oForm, "etFind", prjcode)
            oApplication.Utilities.setEdittextvalue(oForm, "7", prjname)
            oApplication.Utilities.setEdittextvalue(oForm, "9", businessprocess)
            oApplication.Utilities.setEdittextvalue(oForm, "11", busienssactivity)
            oApplication.Utilities.setEdittextvalue(oForm, "15", boqref)
            oApplication.Utilities.setEdittextvalue(oForm, "17", stats)

            '  AddChooseFromList(objform)
            oForm.Items.Item("mtchoose").Enabled = True
            FormatGrid(objGrid)
            objGrid.AutoResizeColumns()
            objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            If objGrid.Rows.Count > 0 Then
                objGrid.Rows.SelectedRows.Add(0)
            End If
            objform.Freeze(False)
            objform.Update()

        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub
#End Region


    Private Function ValidateDeletion(ByVal aForm As SAPbouiCOM.Form, ByVal aItemCode As String) As Boolean
        If intSelectedMatrixrow <= 0 Then
            Return True
        End If
        ' oMatrix = frmSourceMatrix
        Dim strPrjCode, strActivity, strProcess, strMessage As String
        Dim otemp As SAPbobsCOM.Recordset
        strMessage = ""


        'oApplication.Utilities.setEdittextvalue(oForm, "etFind", prjcode)
        'oApplication.Utilities.setEdittextvalue(oForm, "7", prjname)
        'oApplication.Utilities.setEdittextvalue(oForm, "9", businessprocess)
        'oApplication.Utilities.setEdittextvalue(oForm, "11", busienssactivity)
        'oApplication.Utilities.setEdittextvalue(oForm, "15", boqref)
        'oApplication.Utilities.setEdittextvalue(oForm, "17", stats)
        If 1 = 1 Then 'oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intSelectedMatrixrow) <> "" Then
            'oComboBox = aForm.Items.Item("4").Specific
            'strPrjCode = oComboBox.Selected.Value
            'strProcess = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intSelectedMatrixrow)
            'strActivity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intSelectedMatrixrow)
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select isnull(sum(X.Quantity),0) 'Qty' from ("

            strSQL = strSQL & " select itemcode,Quantity from POR1 where Project='" & prjcode & "' and U_Z_MDName='" & businessprocess & "' and U_Z_ActName='" & busienssactivity & "' and ItemCode='" & aItemCode & "'"
            strSQL = strSQL & " union all "
            strSQL = strSQL & " select itemcode,Quantity from PDN1 where Project='" & prjcode & "' and U_Z_MDName='" & businessprocess & "' and U_Z_ActName='" & busienssactivity & "' and ItemCode='" & aItemCode & "'"
            strSQL = strSQL & " union all"
            strSQL = strSQL & " select itemcode,Quantity from PCH1 where Project='" & prjcode & "' and U_Z_MDName='" & businessprocess & "' and U_Z_ActName='" & busienssactivity & "' and ItemCode='" & aItemCode & "'"
            strSQL = strSQL & " ) as x group by x.ItemCode"
            'otemp.DoQuery("select * from [@Z_OEXP] T0 Inner Join [@Z_EXP1] T1 on T1.U_Z_REFCODE=T0.Code where U_Z_PRJCODE='" & strPrjCode & "'")
            'strMessage = "Project Code=" & strPrjCode
            otemp.DoQuery(strSQL)
            If otemp.Fields.Item(0).Value > 0 Then
                strMessage = "Project Code : " & strPrjCode & " , Phase : " & strProcess & " , Activity : " & strActivity & " ItemCode : " & aItemCode
                oApplication.Utilities.Message(" Purchase Order / GRPO / AP Invoice already entered for this " & strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Return True

    End Function

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
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL_6"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            'oCon = oCons.Add()

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "CFL_7"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "PrchseItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        objGrid = aForm.Items.Item("mtchoose").Specific
        If objGrid.DataTable.GetValue("U_Z_ITEMCODE", objGrid.DataTable.Rows.Count - 1) <> "" Then
            objGrid.DataTable.Rows.Add()
        End If
    End Sub

    Private Sub DeleteRos(ByVal aForm As SAPbouiCOM.Form)
        objGrid = aForm.Items.Item("mtchoose").Specific
        For intRow As Integer = 0 To objGrid.DataTable.Rows.Count - 1
            If objGrid.Rows.IsSelected(intRow) Then
                Dim oTest As SAPbobsCOM.Recordset
                If ValidateDeletion(aForm, objGrid.DataTable.GetValue("U_Z_ITEMCODE", intRow)) = False Then
                    Exit Sub
                End If


                Dim dblqty As Double
                Try
                    dblqty = objGrid.DataTable.GetValue("U_Z_DOCENTRY", intRow)

                Catch ex As Exception
                    dblqty = 0
                End Try

                If dblqty > 0 Then
                    oApplication.Utilities.Message("Purchase Request already created for this row. You can not remove.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTest.DoQuery("Update [@Z_PRJ2] set name= name +'_XD' where Code='" & objGrid.DataTable.GetValue("Code", intRow) & "'")
                objGrid.DataTable.Rows.Remove(intRow)
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No rows selected", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    End Sub
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid)
        aGrid.Columns.Item("Code").Visible = False
        aGrid.Columns.Item("Name").Visible = False
        aGrid.Columns.Item("U_Z_PRJCODE").Visible = False
        aGrid.Columns.Item("U_Z_PRJNAME").Visible = False
        aGrid.Columns.Item("U_Z_MODNAME").Visible = False
        aGrid.Columns.Item("U_Z_ACTNAME").Visible = False
        aGrid.Columns.Item("U_Z_BOQREF").Visible = False
        aGrid.Columns.Item("U_Z_STATUS").Visible = False
        aGrid.Columns.Item("U_Z_ITEMCODE").TitleObject.Caption = "Item Code"

        oEditTextColumn = aGrid.Columns.Item("U_Z_ITEMCODE")
        oEditTextColumn.ChooseFromListUID = "CFL_7"
        oEditTextColumn.ChooseFromListAlias = "ItemCode"
        oEditTextColumn.LinkedObjectType = "4"
        aGrid.Columns.Item("U_Z_ITEMNAME").TitleObject.Caption = "Item Description"
        aGrid.Columns.Item("U_Z_ITEMNAME").Editable = False
        aGrid.Columns.Item("U_Z_REQQTY").TitleObject.Caption = "Required.Quantity"

        oEditTextColumn = aGrid.Columns.Item("U_Z_REQQTY")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        aGrid.Columns.Item("U_Z_REQDATE").TitleObject.Caption = "Required.Date"
        aGrid.Columns.Item("U_Z_UOM").TitleObject.Caption = "UOM"
        aGrid.Columns.Item("U_Z_UOM").Editable = False
        aGrid.Columns.Item("U_Z_ESTCOST").TitleObject.Caption = "Estimated Cost"
        oEditTextColumn = aGrid.Columns.Item("U_Z_ESTCOST")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        aGrid.Columns.Item("U_Z_VENDOR").TitleObject.Caption = "Vendor Code"
        oEditTextColumn = aGrid.Columns.Item("U_Z_VENDOR")
        oEditTextColumn.ChooseFromListUID = "CFL_6"
        oEditTextColumn.ChooseFromListAlias = "CardCode"
        oEditTextColumn.LinkedObjectType = 2
        aGrid.Columns.Item("U_Z_VENDORNAME").TitleObject.Caption = "Vendor Name"
        aGrid.Columns.Item("U_Z_VENDORNAME").Editable = False
        aGrid.Columns.Item("U_Z_COMMENTS").TitleObject.Caption = "Comments"
        aGrid.Columns.Item("U_Z_DOCENTRY").TitleObject.Caption = "Purchase Request DocEntry"
        aGrid.Columns.Item("U_Z_DOCNUM").TitleObject.Caption = "Purchase Request Number"
        aGrid.Columns.Item("U_Z_DOCENTRY").Editable = False
        aGrid.Columns.Item("U_Z_DOCNUM").Editable = False
        aGrid.Columns.Item("U_Z_PR").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        aGrid.Columns.Item("U_Z_PR").TitleObject.Caption = "Purchase Request"

        aGrid.Columns.Item("U_Z_CNTID").TitleObject.Caption = "Contract ID"
        aGrid.Columns.Item("U_Z_POSITION").TitleObject.Caption = "Contractor Name"
        aGrid.Columns.Item("U_Z_CNTID").Visible = False
        aGrid.Columns.Item("U_Z_POSITION").Visible = False
        oEditTextColumn = aGrid.Columns.Item("U_Z_DOCENTRY")
        oEditTextColumn.LinkedObjectType = "1470000113"

        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
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
        oBPGrid = aform.Items.Item("mtchoose").Specific

        ousertable = oApplication.Company.UserTables.Item("Z_PRJ2")
        ocheckbox = oBPGrid.Columns.Item("U_Z_PR")
        'strEmpID = oApplication.Utilities.getEdittextvalue(aform, "6")
        ' strEmpName = oApplication.Utilities.getEdittextvalue(aform, "8")
        For intRow As Integer = 0 To oBPGrid.DataTable.Rows.Count - 1
            If oBPGrid.DataTable.GetValue("U_Z_ITEMCODE", intRow) <> "" Then
                strCode = oBPGrid.DataTable.GetValue("Code", intRow)
                If strCode = "" Then
                    strCode = oApplication.Utilities.getMaxCode("@Z_PRJ2", "Code")
                    ousertable.Code = strCode
                    ousertable.Name = strCode
                    ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = oApplication.Utilities.getEdittextvalue(aform, "etFind")
                    ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                    ousertable.UserFields.Fields.Item("U_Z_ModName").Value = oApplication.Utilities.getEdittextvalue(aform, "9")

                    ousertable.UserFields.Fields.Item("U_Z_ActName").Value = oApplication.Utilities.getEdittextvalue(aform, "11")
                    ousertable.UserFields.Fields.Item("U_Z_BOQRef").Value = oApplication.Utilities.getEdittextvalue(aform, "15")
                    ousertable.UserFields.Fields.Item("U_Z_Status").Value = oApplication.Utilities.getEdittextvalue(aform, "17")
                    ousertable.UserFields.Fields.Item("U_Z_ItemCode").Value = oBPGrid.DataTable.GetValue("U_Z_ITEMCODE", intRow)
                    ousertable.UserFields.Fields.Item("U_Z_ItemName").Value = oBPGrid.DataTable.GetValue("U_Z_ITEMNAME", intRow)
                    ousertable.UserFields.Fields.Item("U_Z_UOM").Value = oBPGrid.DataTable.GetValue("U_Z_UOM", intRow)
                    ousertable.UserFields.Fields.Item("U_Z_ReqQty").Value = oBPGrid.DataTable.GetValue("U_Z_REQQTY", intRow)

                    ousertable.UserFields.Fields.Item("U_Z_EstCost").Value = oBPGrid.DataTable.GetValue("U_Z_ESTCOST", intRow)
                    dtDate = oBPGrid.DataTable.GetValue("U_Z_REQDATE", intRow)
                    If ocheckbox.IsChecked(intRow) Then
                        ousertable.UserFields.Fields.Item("U_Z_PR").Value = "Y"
                    Else
                        ousertable.UserFields.Fields.Item("U_Z_PR").Value = "N"

                    End If
                    ousertable.UserFields.Fields.Item("U_Z_DOCENTRY").Value = oBPGrid.DataTable.GetValue("U_Z_DOCENTRY", intRow)
                    ousertable.UserFields.Fields.Item("U_Z_DOCNUM").Value = oBPGrid.DataTable.GetValue("U_Z_DOCNUM", intRow)

                    '   MsgBox(Year(dtDate))
                    If Year(dtDate) <> 1 Then
                        ousertable.UserFields.Fields.Item("U_Z_ReqDate").Value = oBPGrid.DataTable.GetValue("U_Z_REQDATE", intRow)
                    End If
                    ousertable.UserFields.Fields.Item("U_Z_Vendor").Value = oBPGrid.DataTable.GetValue("U_Z_VENDOR", intRow)
                    ousertable.UserFields.Fields.Item("U_Z_VendorName").Value = oBPGrid.DataTable.GetValue("U_Z_VENDORNAME", intRow)
                    ousertable.UserFields.Fields.Item("U_Z_Comments").Value = oBPGrid.DataTable.GetValue("U_Z_COMMENTS", intRow)
                    ousertable.UserFields.Fields.Item("U_Z_CNTID").Value = oApplication.Utilities.getEdittextvalue(aform, "25")
                    ousertable.UserFields.Fields.Item("U_Z_POSITION").Value = oApplication.Utilities.getEdittextvalue(aform, "27")
                    ousertable.UserFields.Fields.Item("U_Z_CUSTCNTID").Value = oApplication.Utilities.getEdittextvalue(aform, "29")

                    If ousertable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Else
                    If ousertable.GetByKey(strCode) Then
                        ' strCode = oApplication.Utilities.getMaxCode("@Z_OLEV", "Code")
                        ousertable.Code = strCode
                        ousertable.Name = strCode
                        ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = oApplication.Utilities.getEdittextvalue(aform, "etFind")
                        ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                        ousertable.UserFields.Fields.Item("U_Z_ModName").Value = oApplication.Utilities.getEdittextvalue(aform, "9")

                        ousertable.UserFields.Fields.Item("U_Z_ActName").Value = oApplication.Utilities.getEdittextvalue(aform, "11")
                        ousertable.UserFields.Fields.Item("U_Z_BOQRef").Value = oApplication.Utilities.getEdittextvalue(aform, "15")
                        ousertable.UserFields.Fields.Item("U_Z_Status").Value = oApplication.Utilities.getEdittextvalue(aform, "17")
                        ousertable.UserFields.Fields.Item("U_Z_ItemCode").Value = oBPGrid.DataTable.GetValue("U_Z_ITEMCODE", intRow)
                        ousertable.UserFields.Fields.Item("U_Z_ItemName").Value = oBPGrid.DataTable.GetValue("U_Z_ITEMNAME", intRow)
                        ousertable.UserFields.Fields.Item("U_Z_UOM").Value = oBPGrid.DataTable.GetValue("U_Z_UOM", intRow)
                        ousertable.UserFields.Fields.Item("U_Z_ReqQty").Value = oBPGrid.DataTable.GetValue("U_Z_REQQTY", intRow)

                        ousertable.UserFields.Fields.Item("U_Z_EstCost").Value = oBPGrid.DataTable.GetValue("U_Z_ESTCOST", intRow)
                        dtDate = oBPGrid.DataTable.GetValue("U_Z_REQDATE", intRow)
                        If Year(dtDate) <> 1 Then
                            ousertable.UserFields.Fields.Item("U_Z_ReqDate").Value = oBPGrid.DataTable.GetValue("U_Z_REQDATE", intRow)
                        End If
                        ousertable.UserFields.Fields.Item("U_Z_Vendor").Value = oBPGrid.DataTable.GetValue("U_Z_VENDOR", intRow)
                        ousertable.UserFields.Fields.Item("U_Z_VendorName").Value = oBPGrid.DataTable.GetValue("U_Z_VENDORNAME", intRow)
                        ousertable.UserFields.Fields.Item("U_Z_Comments").Value = oBPGrid.DataTable.GetValue("U_Z_COMMENTS", intRow)
                        If ocheckbox.IsChecked(intRow) Then
                            ousertable.UserFields.Fields.Item("U_Z_PR").Value = "Y"
                        Else
                            ousertable.UserFields.Fields.Item("U_Z_PR").Value = "N"
                        End If
                        ousertable.UserFields.Fields.Item("U_Z_DOCENTRY").Value = oBPGrid.DataTable.GetValue("U_Z_DOCENTRY", intRow)
                        ousertable.UserFields.Fields.Item("U_Z_DOCNUM").Value = oBPGrid.DataTable.GetValue("U_Z_DOCNUM", intRow)
                        ousertable.UserFields.Fields.Item("U_Z_CNTID").Value = oApplication.Utilities.getEdittextvalue(aform, "25")
                        ousertable.UserFields.Fields.Item("U_Z_POSITION").Value = oApplication.Utilities.getEdittextvalue(aform, "27")
                        ousertable.UserFields.Fields.Item("U_Z_CUSTCNTID").Value = oApplication.Utilities.getEdittextvalue(aform, "29")
                        If ousertable.Update <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                End If
            End If
        Next
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec.DoQuery("Delete from [@Z_PRJ2] where name like '%_XD'")
        oApplication.Utilities.Message("Operation completed successfuly", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#Region "Update On hand Qty"
    Private Sub FillOnhandqty(ByVal strItemcode As String, ByVal strwhs As String, ByVal aGrid As SAPbouiCOM.Grid)
        Dim oTemprec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strBin, strSql As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strBin = aGrid.DataTable.GetValue(0, intRow)
            strSql = "Select isnull(Sum(U_InQty)-sum(U_OutQty),0) from [@DABT_BTRN] where U_Itemcode='" & strItemcode & "' and U_FrmWhs='" & strwhs & "' and U_BinCode='" & strBin & "'"
            oTemprec.DoQuery(strSql)
            Dim dblOnhand As Double
            dblOnhand = oTemprec.Fields.Item(0).Value

            aGrid.DataTable.SetValue(2, intRow, dblOnhand.ToString)
        Next
    End Sub
#End Region

#Region "Get Form"
    '******************************************************************
    'Type               : Function
    'Name               : GetForm
    'Parameter          : FormUID
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Get The Form
    '******************************************************************
    Public Function GetForm(ByVal FormUID As String) As SAPbouiCOM.Form
        Return oApplication.SBO_Application.Forms.Item(FormUID)
    End Function
#End Region

#Region "FormDataEvent"


#End Region

#Region "Class Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

    End Sub
#End Region

#Region "Validation"
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            If aGrid.DataTable.GetValue("U_Z_ITEMCODE", intRow) <> "" Then
                If aGrid.DataTable.GetValue("U_Z_REQQTY", intRow) <= 0 Then
                    oApplication.Utilities.Message("Required quantity should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aGrid.Columns.Item("U_Z_REQQTY").Click(intRow, False, 1)
                    Return False
                End If
                If aGrid.DataTable.GetValue("U_Z_ESTCOST", intRow) <= 0 Then
                    oApplication.Utilities.Message("Esimated Cost should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aGrid.Columns.Item("U_Z_ESTCOST").Click(intRow, False, 1)
                    Return False
                End If

            End If
        Next
        Return True
    End Function
#End Region

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = "frm_BOQ" Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "mtchoose" Then
                                    objGrid = oForm.Items.Item("mtchoose").Specific
                                    Try
                                        If objGrid.DataTable.GetValue("U_Z_DOCNUM", pVal.Row) <> 0 Then
                                            oApplication.Utilities.Message("Purase Request already created for this entry . You can not modify any data", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If pVal.ItemUID = "mtchoose" Then
                                    objGrid = oForm.Items.Item("mtchoose").Specific
                                    Try
                                        If objGrid.DataTable.GetValue("U_Z_DOCNUM", pVal.Row) <> 0 Then
                                            oApplication.Utilities.Message("Purase Request already created for this entry . You can not modify any data", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If pVal.ItemUID = "mtchoose" Then
                                    objGrid = oForm.Items.Item("mtchoose").Specific
                                    Try
                                        If objGrid.DataTable.GetValue("U_Z_DOCNUM", pVal.Row) <> 0 Then
                                            oApplication.Utilities.Message("Purase Request already created for this entry . You can not modify any data", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                Dim oTempRec As SAPbobsCOM.Recordset
                                If pVal.ItemUID = "2" Then
                                    oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTempRec.DoQuery("Update  [@Z_PRJ2] set name= code where name like '%_XD' and U_Z_BOQRef='" & oApplication.Utilities.getEdittextvalue(oForm, "17") & "'")
                                End If
                                If pVal.ItemUID = "btnChoose" Then
                                    objGrid = oForm.Items.Item("mtchoose").Specific
                                    If validation(objGrid) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "mtchoose" Then
                                    objGrid = oForm.Items.Item("mtchoose").Specific
                                    Try
                                        If objGrid.DataTable.GetValue("U_Z_DOCNUM", pVal.Row) <> 0 Then
                                            oApplication.Utilities.Message("Purase Request already created for this entry . You can not modify any data", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception

                                    End Try

                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '  oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "12"
                                        AddRow(oForm)
                                    Case "13"
                                        DeleteRos(oForm)
                                    Case "18"
                                        If AddToUDT_Table(oForm) Then
                                           
                                            Dim oTemp As SAPbobsCOM.Recordset
                                            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            If oApplication.SBO_Application.MessageBox("Do you want to create Purchase Request Documents?", , "Yes", "No") = 2 Then
                                                Exit Sub
                                            End If
                                            Dim oSourceform As SAPbouiCOM.Form
                                            Dim oMatrix As SAPbouiCOM.Grid
                                            oSourceform = oApplication.SBO_Application.Forms.Item(SourceFormUID)
                                            oMatrix = oSourceform.Items.Item(ItemUID).Specific
                                            '  oApplication.Utilities.SetMatrixValues(oMatrix, "BOQ", sourcerowId, oApplication.Utilities.getEdittextvalue(oForm, "15"))
                                            oMatrix.DataTable.SetValue("U_Z_BOQ", sourcerowId, oApplication.Utilities.getEdittextvalue(oForm, "15"))
                                            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oTemp.DoQuery("Select sum(U_Z_REQQTY) 'Qty',sum(U_Z_ESTCOST) 'Cost' from [@Z_PRJ2] where U_Z_BOQRef='" & oApplication.Utilities.getEdittextvalue(oForm, "15") & "'")
                                            oMatrix.DataTable.SetValue("U_Z_QUANTITY", sourcerowId, oTemp.Fields.Item(0).Value)
                                            oMatrix.DataTable.SetValue("U_Z_AMOUNT", sourcerowId, oTemp.Fields.Item(1).Value)
                                            oApplication.Utilities.createPurchaseRequest(oApplication.Utilities.getEdittextvalue(oForm, "15"))
                                            oForm.Close()
                                            '    PopulateBoQDetails(oForm)
                                        End If

                                    Case "btnChoose"
                                        If oApplication.SBO_Application.MessageBox("Do you want to Confirm the changes?", , "Yes", "No") = 2 Then
                                            Exit Sub
                                        End If
                                        If AddToUDT_Table(oForm) Then
                                            If BinDescrUID = "" Then
                                                Dim oSourceform As SAPbouiCOM.Form
                                                Dim oMatrix As SAPbouiCOM.Matrix
                                                oSourceform = oApplication.SBO_Application.Forms.Item(SourceFormUID)
                                                oMatrix = oSourceform.Items.Item(ItemUID).Specific
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "BOQ", sourcerowId, oApplication.Utilities.getEdittextvalue(oForm, "15"))
                                                Dim oTemp As SAPbobsCOM.Recordset
                                                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oTemp.DoQuery("Select sum(U_Z_REQQTY) 'Qty',sum(U_Z_ESTCOST) 'Cost' from [@Z_PRJ2] where U_Z_BOQRef='" & oApplication.Utilities.getEdittextvalue(oForm, "15") & "'")
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "Qty", sourcerowId, oTemp.Fields.Item(0).Value)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", sourcerowId, oTemp.Fields.Item(1).Value)
                                                oForm.Close()
                                            Else
                                                Dim oSourceform As SAPbouiCOM.Form
                                                Dim oMatrix As SAPbouiCOM.Grid
                                                oSourceform = oApplication.SBO_Application.Forms.Item(SourceFormUID)
                                                oMatrix = oSourceform.Items.Item(ItemUID).Specific
                                                '  oApplication.Utilities.SetMatrixValues(oMatrix, "BOQ", sourcerowId, oApplication.Utilities.getEdittextvalue(oForm, "15"))
                                                oMatrix.DataTable.SetValue("U_Z_BOQ", sourcerowId, oApplication.Utilities.getEdittextvalue(oForm, "15"))
                                                Dim oTemp As SAPbobsCOM.Recordset
                                                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oTemp.DoQuery("Select sum(U_Z_REQQTY) 'Qty',sum(U_Z_ESTCOST) 'Cost' from [@Z_PRJ2] where U_Z_BOQRef='" & oApplication.Utilities.getEdittextvalue(oForm, "15") & "'")
                                                oMatrix.DataTable.SetValue("U_Z_QUANTITY", sourcerowId, oTemp.Fields.Item(0).Value)
                                                oMatrix.DataTable.SetValue("U_Z_AMOUNT", sourcerowId, oTemp.Fields.Item(1).Value)
                                                oForm.Close()
                                            End If

                                        End If
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

                                        If pVal.ItemUID = "mtchoose" And pVal.ColUID = "U_Z_ITEMCODE" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            val1 = oDataTable.GetValue("ItemName", 0)
                                            objGrid = oForm.Items.Item("mtchoose").Specific
                                            objGrid.DataTable.SetValue("U_Z_ITEMCODE", pVal.Row, val)
                                            objGrid.DataTable.SetValue("U_Z_ITEMNAME", pVal.Row, val1)
                                            objGrid.DataTable.SetValue("U_Z_UOM", pVal.Row, oDataTable.GetValue("InvntryUom", 0))

                                        End If

                                        If pVal.ItemUID = "mtchoose" And pVal.ColUID = "U_Z_VENDOR" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            objGrid = oForm.Items.Item("mtchoose").Specific
                                            objGrid.DataTable.SetValue("U_Z_VENDOR", pVal.Row, val)
                                            objGrid.DataTable.SetValue("U_Z_VENDORNAME", pVal.Row, val1)

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

End Class
