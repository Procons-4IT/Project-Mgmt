Public Class clsBOMCopy
    Inherits clsBase

#Region "Declarations"
    Public Shared ItemUID As String
    Public Shared SourceFormUID As String
    Public Shared SourceLabel As Integer
    Public Shared CFLChoice As String
    Public Shared ItemCode As String
    Public Shared sourceItemCode As String
    Public Shared choice As String
    Public Shared sourceColumID As String
    Public Shared sourcerowId As Integer
    Public Shared BinDescrUID As String
    Public Shared Documentchoice As String
    Public Shared ContractID As String

    Private oDbDatasource As SAPbouiCOM.DBDataSource
    Private Ouserdatasource As SAPbouiCOM.UserDataSource
    Private oConditions As SAPbouiCOM.Conditions
    Private ocondition As SAPbouiCOM.Condition
    Private intRowId As Integer
    Private strRowNum As Integer
    Private i As Integer
    Private oedit As SAPbouiCOM.EditTextColumn
    '   Private oForm As SAPbouiCOM.Form
    Private objSoureceForm As SAPbouiCOM.Form
    Private objform As SAPbouiCOM.Form
    Private oMatrix, oGrid, oGrid1 As SAPbouiCOM.Grid
    Private osourcegrid As SAPbouiCOM.Matrix
    Private oChecbox As SAPbouiCOM.CheckBoxColumn
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
    Private strContractID, strContractName, strCustCntID As String
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
            Dim ObjSegRecSet As SAPbobsCOM.Recordset
            ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oForm = objform
            oForm.Freeze(True)

            oGrid = oForm.Items.Item("6").Specific
            oGrid1 = oForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery("Select ItemCode,ItemName,' ' 'Select' from OITM where TreeType<>'N'")
            ' oGrid1.DataTable.ExecuteQuery("SELECT T1.[Father], T1.[ChildNum], T1.[Code],T2.[ItemName] FROM OITT T0  INNER JOIN ITT1 T1 ON T0.Code = T1.Father INNER JOIN OITM T2 ON T1.Code = T2.ItemCode  where 1=2 ORDER BY T1.[Father]")
            Dim st As String = "SELECT T1.[Father] 'ParentItem', T1.[Code] 'ItemCode',T2.[ItemName], T2.cardcode 'CardCode',T3.CardName,T2.InvntryUom 'UoM'  ,T1.Quantity 'Quantity',' ' 'Select' FROM OITT T0  INNER JOIN ITT1 T1 ON T0.Code = T1.Father Inner JOIN OITM T2 ON T1.Code = T2.ItemCode left outer join OCRD T3 on T3.CardCode=T2.CardCode "
            oGrid1.DataTable.ExecuteQuery(st & "  where 1=2 ORDER BY T1.[Father]")
            formatGrid(oGrid, "BOM", oForm)
            formatGrid(oGrid1, "Item", oForm)
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub
#End Region

    Private Sub databind(ByVal aform As SAPbouiCOM.Form)
        If aform.PaneLevel = 2 Then
            oGrid = aform.Items.Item("6").Specific
            oGrid1 = aform.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery("Select ItemCode,ItemName,' ' 'Select' from OITM where TreeType<>'N'")
            Dim st As String = "SELECT T1.[Father] 'ParentItem', T1.[Code] 'ItemCode',T2.[ItemName], T2.cardcode 'CardCode',T3.CardName,T2.InvntryUom 'UoM' ,T1.Quantity , 'Quantity',' ' 'Select'  FROM OITT T0  INNER JOIN ITT1 T1 ON T0.Code = T1.Father Inner JOIN OITM T2 ON T1.Code = T2.ItemCode left outer join OCRD T3 on T3.CardCode=T2.CardCode "
            oGrid1.DataTable.ExecuteQuery(st & " where 1=2 ORDER BY T1.[Father]")
            formatGrid(oGrid, "BOM", aform)
            formatGrid(oGrid1, "Item", aform)
        ElseIf aform.PaneLevel = 3 Then
            oGrid = aform.Items.Item("6").Specific
            Dim strParentItem As String = "'x'"
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oChecbox = oGrid.Columns.Item("Select")
                If oChecbox.IsChecked(intRow) Then
                    strParentItem = strParentItem & ",'" & oGrid.DataTable.GetValue("ItemCode", intRow) & "'"
                End If
            Next
            oGrid1 = aform.Items.Item("7").Specific
            Dim st As String = "SELECT T1.[Father] 'ParentItem', T1.[Code] 'ItemCode',T2.[ItemName], T2.cardcode 'CardCode',T3.CardName,T2.InvntryUom 'UoM' ,T1.Quantity  'Quantity',' ' 'Select'  FROM OITT T0  INNER JOIN ITT1 T1 ON T0.Code = T1.Father Inner JOIN OITM T2 ON T1.Code = T2.ItemCode left outer join OCRD T3 on T3.CardCode=T2.CardCode "
            st = st & "  where T1.Father in (" & strParentItem & ") ORDER BY T1.[Father]"
            oGrid1.DataTable.ExecuteQuery(st)
            formatGrid(oGrid1, "Item", aform)
        End If
    End Sub
    Private Sub AssignLineNo1(ByVal aform As SAPbouiCOM.Form, ByVal agrid As SAPbouiCOM.Grid)
        Try
            aform.Freeze(True)
            For count As Integer = 0 To agrid.DataTable.Rows.Count - 1
                agrid.RowHeaders.SetText(count, count + 1)
            Next
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
    Private Sub formatGrid(ByVal aGrid As SAPbouiCOM.Grid, ByVal aChoice As String, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        Select Case aChoice
            Case "BOM"
                aGrid.Columns.Item("ItemCode").TitleObject.Caption = "ItemCode"
                aGrid.Columns.Item("ItemCode").Editable = False
                oedit = aGrid.Columns.Item("ItemCode")
                oedit.LinkedObjectType = "4"
                aGrid.Columns.Item("ItemName").TitleObject.Caption = "Description"
                aGrid.Columns.Item("ItemName").Editable = False
                aGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                aGrid.Columns.Item("Select").TitleObject.Caption = "Select"
                '  aGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
            Case "Item"
                aGrid.Columns.Item("ParentItem").TitleObject.Caption = "Parent ItemCode"
                aGrid.Columns.Item("ParentItem").Editable = False
                oedit = aGrid.Columns.Item("ParentItem")
                oedit.LinkedObjectType = "4"
                aGrid.Columns.Item("CardCode").TitleObject.Caption = "Supplier Code"
                aGrid.Columns.Item("CardCode").Editable = False
                oedit = aGrid.Columns.Item("CardCode")
                oedit.LinkedObjectType = "2"
                aGrid.Columns.Item("CardName").TitleObject.Caption = "Supplier Name"
                aGrid.Columns.Item("CardName").Editable = False
                aGrid.Columns.Item("ItemName").TitleObject.Caption = "Description"
                aGrid.Columns.Item("ItemName").Editable = False
                aGrid.Columns.Item("ItemCode").TitleObject.Caption = "ItemCode"
                aGrid.Columns.Item("ItemCode").Editable = False
                oedit = aGrid.Columns.Item("ItemCode")
                oedit.LinkedObjectType = "4"
                aGrid.Columns.Item("ItemName").TitleObject.Caption = "Description"
                aGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                aGrid.Columns.Item("Select").TitleObject.Caption = "Select"
                aGrid.Columns.Item("UoM").TitleObject.Caption = "InvntryUom"
                aGrid.Columns.Item("UoM").Editable = False
                aGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
        End Select
        AssignLineNo(aform, aGrid)
        aform.Freeze(False)
    End Sub

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

#Region "getBOQReference"
    Private Function getBOQReference(ByVal aItemCode As String, ByVal aProject As String, ByVal aProcess As String, ByVal aActivity As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select isnull(U_Z_BOQREF,'') from [@Z_PRJ2] where U_Z_ItemCode='" & aItemCode & "' and  U_Z_PRJCODE='" & aProject.Replace("'", "''") & "' and U_Z_MODNAME='" & aProcess.Replace("'", "''") & "' and U_Z_ACTNAME='" & aActivity.Replace("'", "''") & "'")
        Return oTest.Fields.Item(0).Value
    End Function
#End Region

    Private Function CopyBOM(ByVal aform As SAPbouiCOM.Form) As Boolean
        frmSourceForm = oApplication.SBO_Application.Forms.Item(SourceFormUID)

        Try
            frmSourceForm.Freeze(True)

            frmSourceGrid = frmSourceForm.Items.Item(ItemUID).Specific
            oGrid = aform.Items.Item("7").Specific
            For intLoop As Integer = 0 To oGrid.Rows.Count - 1
                oChecbox = oGrid.Columns.Item("Select")
                If oChecbox.IsChecked(intLoop) = True Then
                    If frmSourceGrid.DataTable.GetValue("U_Z_ITEMCODE", frmSourceGrid.DataTable.Rows.Count - 1) <> "" Then
                        frmSourceGrid.DataTable.Rows.Add()
                    End If
                    frmSourceGrid.DataTable.SetValue("U_Z_ITEMCODE", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("ItemCode", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_ITEMNAME", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("ItemName", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_UOM", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("UoM", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_VENDOR", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("CardCode", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_VENDORNAME", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("CardName", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_REQQTY", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("Quantity", intLoop))
                    Dim dblPrice As Double = oApplication.Utilities.getPriceforSupplier(oGrid.DataTable.GetValue("CardCode", intLoop), oGrid.DataTable.GetValue("ItemCode", intLoop))
                    Dim dblQty As Double = oGrid.DataTable.GetValue("Quantity", intLoop)

                    Try
                        frmSourceGrid.DataTable.SetValue("U_Z_UnitPrice", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice)
                    Catch ex As Exception
                        frmSourceGrid.DataTable.SetValue("U_Z_UNITPRICE", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice)
                    End Try

                    Try
                        frmSourceGrid.DataTable.SetValue("U_Z_ESTCOST", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice * dblQty)
                    Catch ex As Exception
                        frmSourceGrid.DataTable.SetValue("U_Z_ESTCOST", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice * dblQty)
                    End Try
                End If
            Next
            AssignLineNo1(frmSourceForm, frmSourceGrid)
            frmSourceForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            frmSourceForm.Freeze(False)
            Return False
        End Try
    End Function
    Private Sub AssignLineNo(ByVal aform As SAPbouiCOM.Form, ByVal agrid As SAPbouiCOM.Grid)
        Try
            aform.Freeze(True)
            For count As Integer = 0 To agrid.DataTable.Rows.Count - 1
                agrid.RowHeaders.SetText(count, count + 1)
            Next
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub SelectAll(ByVal aform As SAPbouiCOM.Form, ByVal achoice As Boolean)
        aform.Freeze(True)
        oGrid = aform.Items.Item("7").Specific
        oChecbox = oGrid.Columns.Item("Select")
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oChecbox.Check(intRow, achoice)
        Next
        aform.Freeze(False)
    End Sub
#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.BeforeAction
            Case True

            Case False
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        If pVal.ItemUID = "8" Then
                            SelectAll(oForm, True)
                        End If
                        If pVal.ItemUID = "9" Then
                            SelectAll(oForm, False)
                        End If
                        If pVal.ItemUID = "3" Then
                            oForm.PaneLevel = oForm.PaneLevel + 1
                            databind(oForm)
                        End If
                        If pVal.ItemUID = "4" Then
                            oForm.PaneLevel = oForm.PaneLevel - 1
                            databind(oForm)
                        End If
                        If pVal.ItemUID = "5" Then
                            If CopyBOM(oForm) Then
                                oForm.Close()
                            End If
                        End If
                End Select

        End Select
    End Sub
#End Region

End Class
