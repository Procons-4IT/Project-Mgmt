Public Class clsMyTasks
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oCheckbox As SAPbouiCOM.CheckBox
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private oTemp As SAPbobsCOM.Recordset
    Private InvBaseDocNo, strname As String
    Private InvForConsumedItems As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_MyTasks, frm_MyTasks)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Dim st As String = oApplication.Company.UserName
        st = "SELECT empID, isnull(firstName,'') + ' '+ isnull(middleName,'') + ' '+isnull(lastName,'')  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & st & "'"
        Dim oTes As SAPbobsCOM.Recordset
        oTes = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTes.DoQuery(st)
        oApplication.Utilities.setEdittextvalue(oForm, "9", oTes.Fields.Item(0).Value)
        oApplication.Utilities.setEdittextvalue(oForm, "17", oTes.Fields.Item(1).Value)
        oForm.Items.Item("100").TextStyle = SAPbouiCOM.BoTextStyle.ts_UNDERLINE
        oCombobox = oForm.Items.Item("25").Specific
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("T1.U_Z_PRJCODE", "Project Code")
        oCombobox.ValidValues.Add("T0.U_Z_STATUS", "Project Status")

        oForm.Items.Item("25").DisplayDesc = True

        oCombobox = oForm.Items.Item("22").Specific
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("T1.U_Z_PRJCODE", "Project Code")
        oCombobox.ValidValues.Add("T0.U_Z_FROMDATE", "Start Date")
        oCombobox.ValidValues.Add("T0.U_Z_TODATE", "End Date")
        oForm.Items.Item("25").DisplayDesc = True

        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        Databind(oForm)
        'AddChooseFromList(oForm)
    End Sub

    Private Sub FillValueCombo(ByVal aForm As SAPbouiCOM.Form)
        Dim strFieldchoice As String
        oCombobox = aForm.Items.Item("25").Specific
        strFieldchoice = oCombobox.Selected.Value
        If strFieldchoice = "" Then
            oCombobox = aForm.Items.Item("20").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception

                End Try
            Next
            oCombobox.ValidValues.Add("", "")
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        ElseIf strFieldchoice = "T1.U_Z_PRJCODE" Then
            oCombobox = aForm.Items.Item("20").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception

                End Try

            Next
            oCombobox.ValidValues.Add("", "")
            Dim oTes As SAPbobsCOM.Recordset
            oTes = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTes.DoQuery("Select * from ""@Z_HPRJ"" order by DocEntry")
            For intLoop As Integer = 0 To oTes.RecordCount - 1
                oCombobox.ValidValues.Add(oTes.Fields.Item("U_Z_PRJCODE").Value, oTes.Fields.Item("U_Z_PRJNAME").Value)
                oTes.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Else
            oCombobox = aForm.Items.Item("20").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception

                End Try
            Next
            oCombobox.ValidValues.Add("", "")
            oCombobox.ValidValues.Add("I", "In Process")
            oCombobox.ValidValues.Add("P", "Pending")
            oCombobox.ValidValues.Add("C", "Completed")
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End If
        aForm.Items.Item("20").DisplayDesc = True
        oCombobox = aForm.Items.Item("20").Specific
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


    End Sub
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            'dtTemp.ExecuteQuery("Select * from [@Z_CBS1] where U_Z_RefCode='" & ItemCode & "' order by Code")
            ' Dim st As String = oApplication.Company.UserName
            'st = "SELECT empID  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & st & "'"
            'st = "SELECT T1.[U_Z_PRJCODE] 'Project Code', T1.[U_Z_PRJNAME] 'Project Name', T0.[U_Z_MODNAME] 'Phase', T0.[U_Z_ACTNAME] 'Activity', T0.[U_Z_TYPE] 'Type', T0.[U_Z_EMPID] 'Employee ID',  T0.[U_Z_FROMDATE] 'Start Date', T0.[U_Z_TODATE] 'End Date', T0.[U_Z_DAYS] 'Estimated Days',T0.[U_Z_STATUS] 'Status' FROM [dbo].[@Z_PRJ1]  T0  inner Join  [dbo].[@Z_HPRJ]  T1 on T1.DocEntry=T0.DocEntry where  U_Z_EMPID =(" & st & ")"
            'st = st & " Order by T1.U_Z_PRJCODE,T0.U_Z_MODNAME,T0.U_Z_ACTNAME,T0.U_Z_FROMDATE"

            Dim st As String = oApplication.Company.UserName
            st = "SELECT empID  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & st & "'"
            st = "SELECT T1.[U_Z_PRJCODE] 'Project Code', T1.[U_Z_PRJNAME] 'Project Name', T0.[U_Z_MODNAME] 'Phase', T0.[U_Z_ACTNAME] 'Activity', T0.[U_Z_TYPE] 'Type', T0.[U_Z_EMPID] 'Employee ID',  T0.[U_Z_FROMDATE] 'Start Date', T0.[U_Z_TODATE] 'End Date', T0.[U_Z_HOURS]/8 'Estimated Days',T0.[U_Z_STATUS] 'Status' FROM [dbo].[@Z_PRJ1]  T0  inner Join  [dbo].[@Z_HPRJ]  T1 on T1.DocEntry=T0.DocEntry where T0.U_Z_Type='R' and   T0.U_Z_EMPID =(" & st & ")"
            st = st & " Order by T1.U_Z_PRJCODE,T0.U_Z_MODNAME,T0.U_Z_ACTNAME,T0.U_Z_FROMDATE"

            dtTemp.ExecuteQuery(st)
            oGrid.DataTable = dtTemp
            oGrid.AutoResizeColumns()
            oGrid.CollapseLevel = 2
            oGrid.Columns.Item("Employee ID").Visible = False
            Dim oComboColumn As SAPbouiCOM.ComboBoxColumn
            oGrid.Columns.Item("Type").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("Type")
            oComboColumn.ValidValues.Add("I", "Material")
            oComboColumn.ValidValues.Add("R", "Resource")
            oComboColumn.ValidValues.Add("E", "Expenses")
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("Status")
            oComboColumn.ValidValues.Add("I", "In Process")
            oComboColumn.ValidValues.Add("P", "Pending")
            oComboColumn.ValidValues.Add("C", "Completed")
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            aform.Items.Item("5").Enabled = False


            '' oGrid.Columns.Item("Sequence").TitleObject.Sortable = True
            ' oGrid.Columns.Item("ItemCode").TitleObject.Sortable = True
            ' oEditTextColumn = oGrid.Columns.Item("Project Code")
            ' oEditTextColumn.LinkedObjectType = "4"
            'oEditTextColumn = oGrid.Columns.Item("Warehouse")
            'oEditTextColumn.LinkedObjectType = "64"
            'Dim otest As SAPbobsCOM.Recordset
            'otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ''st = "SELECT T0.[Code] 'ItemCode',Count(*) FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "' group by T0.Code "

            ''otest.DoQuery(st)
            'oCombobox = aform.Items.Item("10").Specific
            'oCombobox.ValidValues.Add("", "")
            '' oCombobox.ValidValues.Add("ItemCode", "T0.Code")
            'oCombobox.ValidValues.Add("Phase", "T0.U_Z_MODNAME")
            'oCombobox.ValidValues.Add("Activity", "T0.U_Z_ACTNAME")
            'oCombobox.ValidValues.Add("StartDate", "T0.U_Z_STARTDATE")
            'oCombobox.ValidValues.Add("EndDate", "T0.U_Z_EndDate")
            'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
            ''st = "SELECT T0.[U_Z_Phase] 'ItemCode',Count(*) FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "' group by T0.U_Z_Phase"
            ''otest.DoQuery(st)
            'oCombobox = aform.Items.Item("12").Specific
            'oCombobox.ValidValues.Add("", "")
            '' oCombobox.ValidValues.Add("ItemCode", "T0.Code")
            'oCombobox.ValidValues.Add("Phase", "T0.U_Z_MODNAME")
            'oCombobox.ValidValues.Add("Activity", "T0.U_Z_ACTNAME")
            'oCombobox.ValidValues.Add("StartDate", "T0.U_Z_STARTDATE")
            'oCombobox.ValidValues.Add("EndDate", "T0.U_Z_EndDate")
            'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
            ''st = "SELECT T0.[U_Z_Seq] 'ItemCode',Count(*) FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "' group by T0.U_Z_Seq"
            ''otest.DoQuery(st)
            'oCombobox = aform.Items.Item("14").Specific
            'oCombobox.ValidValues.Add("", "")
            '' oCombobox.ValidValues.Add("ItemCode", "T0.Code")
            'oCombobox.ValidValues.Add("Phase", "T0.U_Z_MODNAME")
            'oCombobox.ValidValues.Add("Activity", "T0.U_Z_ACTNAME")
            'oCombobox.ValidValues.Add("StartDate", "T0.U_Z_STARTDATE")
            'oCombobox.ValidValues.Add("EndDate", "T0.U_Z_EndDate")
            'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            'oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
            'oForm.Items.Item("10").DisplayDesc = False
            'oForm.Items.Item("12").DisplayDesc = False
            'oForm.Items.Item("14").DisplayDesc = False
            ' assignLineNo(aform)
            aform.Items.Item("5").Enabled = False
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub Databind_Filter(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strField, strChoice, strSortBy As String
            oCombobox = aform.Items.Item("25").Specific
            strField = oCombobox.Selected.Value
            oCombobox = aform.Items.Item("20").Specific
            Try
                strChoice = oCombobox.Selected.Value
            Catch ex As Exception
                strChoice = ""
            End Try

            If strField = "" Then
                strField = " 1=1"
            Else
                If strChoice = "" Then
                    strField = " 1=1"
                Else
                    strField = strField & "='" & strChoice & "'"
                End If
            End If
            oCombobox = aform.Items.Item("22").Specific
            Try
                strSortBy = oCombobox.Selected.Value
            Catch ex As Exception
                strSortBy = ""

            End Try
            'If strSortBy = "" Then
            '    strSortBy = "  Order by T1.U_Z_PRJCODE,T1.[U_Z_PRJNAME],T0.U_Z_MODNAME,T0.U_Z_ACTNAME,T0.U_Z_FROMDATE"
            'Else
            '    strSortBy = "  Order by   " & strSortBy & ",T1.U_Z_PRJCODE,T1.[U_Z_PRJNAME],T0.U_Z_MODNAME,T0.U_Z_ACTNAME"
            'End If
            strSortBy = "  Order by T1.U_Z_PRJCODE,T1.[U_Z_PRJNAME],T0.U_Z_MODNAME,T0.U_Z_ACTNAME,T0.U_Z_FROMDATE"
            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            'dtTemp.ExecuteQuery("Select * from [@Z_CBS1] where U_Z_RefCode='" & ItemCode & "' order by Code")
            Dim st As String = oApplication.Company.UserName
            st = "SELECT empID  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & st & "'"
            st = "SELECT T1.[U_Z_PRJCODE] 'Project Code', T1.[U_Z_PRJNAME] 'Project Name', T0.[U_Z_MODNAME] 'Phase', T0.[U_Z_ACTNAME] 'Activity', T0.[U_Z_TYPE] 'Type', T0.[U_Z_EMPID] 'Employee ID',  T0.[U_Z_FROMDATE] 'Start Date', T0.[U_Z_TODATE] 'End Date', T0.[U_Z_HOURS]/8 'Estimated Days',T0.[U_Z_STATUS] 'Status' FROM [dbo].[@Z_PRJ1]  T0  inner Join  [dbo].[@Z_HPRJ]  T1 on T1.DocEntry=T0.DocEntry where  T0.U_Z_Type='R' and   T0.U_Z_EMPID =(" & st & ")"
            st = st & " and " & strField & strSortBy
            dtTemp.ExecuteQuery(st)
            oGrid.DataTable = dtTemp
            oGrid.AutoResizeColumns()
            oGrid.CollapseLevel = 1
            oGrid.Columns.Item("Employee ID").Visible = False
            Dim oComboColumn As SAPbouiCOM.ComboBoxColumn
            oGrid.Columns.Item("Type").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("Type")
            oComboColumn.ValidValues.Add("I", "Material")
            oComboColumn.ValidValues.Add("R", "Resource")
            oComboColumn.ValidValues.Add("E", "Expenses")
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("Status")
            oComboColumn.ValidValues.Add("I", "In Process")
            oComboColumn.ValidValues.Add("P", "Pending")
            oComboColumn.ValidValues.Add("C", "Completed")
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            aform.Items.Item("5").Enabled = False



            aform.Items.Item("5").Enabled = False
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub Databind_Sort(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strField, strChoice, strSortBy, strFields As String
            oCombobox = aform.Items.Item("25").Specific
            strField = oCombobox.Selected.Value
            oCombobox = aform.Items.Item("20").Specific
            Try
                strChoice = oCombobox.Selected.Value
            Catch ex As Exception
                strChoice = ""
            End Try
            If strField = "" Then
                strField = " 1=1"
            Else
                If strChoice = "" Then
                    strField = " 1=1"
                Else
                    strField = strField & "='" & strChoice & "'"
                End If
            End If
            oCombobox = aform.Items.Item("22").Specific
            Try
                strSortBy = oCombobox.Selected.Value
            Catch ex As Exception
                strSortBy = ""
            End Try
            If strSortBy = "" Then
                strSortBy = "  Order by T1.U_Z_PRJCODE,T0.U_Z_FROMDATE,T0.U_Z_TODATE,T1.[U_Z_PRJNAME],T0.U_Z_MODNAME,T0.U_Z_ACTNAME"
            Else
                If strSortBy = "T0.U_Z_FROMDATE" Then
                    strSortBy = "  Order by   " & strSortBy & ",T1.U_Z_PRJCODE,T0.U_Z_TODATE,T0.U_Z_MODNAME,T0.U_Z_ACTNAME"
                ElseIf strSortBy = "T0.U_Z_TODATE" Then
                    strSortBy = "  Order by   " & strSortBy & ",T1.U_Z_PRJCODE,T0.U_Z_FROMDATE,T0.U_Z_MODNAME,T0.U_Z_ACTNAME"
                Else
                    strSortBy = "  Order by   " & strSortBy & ",T0.U_Z_FROMDATE,T0.U_Z_TODATE,T0.U_Z_MODNAME,T0.U_Z_ACTNAME"
                End If
            End If

            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            'dtTemp.ExecuteQuery("Select * from [@Z_CBS1] where U_Z_RefCode='" & ItemCode & "' order by Code")
            Dim st As String = oApplication.Company.UserName
            st = "SELECT empID  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & st & "'"
            st = "SELECT T1.[U_Z_PRJCODE] 'Project Code', T1.[U_Z_PRJNAME] 'Project Name', T0.[U_Z_MODNAME] 'Phase', T0.[U_Z_ACTNAME] 'Activity', T0.[U_Z_TYPE] 'Type', T0.[U_Z_EMPID] 'Employee ID',  T0.[U_Z_FROMDATE] 'Start Date', T0.[U_Z_TODATE] 'End Date', T0.[U_Z_HOURS]/8 'Estimated Days',T0.[U_Z_STATUS] 'Status' FROM [dbo].[@Z_PRJ1]  T0  inner Join  [dbo].[@Z_HPRJ]  T1 on T1.DocEntry=T0.DocEntry where  T0.U_Z_Type='R' and  T0.U_Z_EMPID =(" & st & ")"
            st = st & " and " & strField & strSortBy
            dtTemp.ExecuteQuery(st)
            oGrid.DataTable = dtTemp
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item("Employee ID").Visible = False
            Dim oComboColumn As SAPbouiCOM.ComboBoxColumn
            oGrid.Columns.Item("Type").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("Type")
            oComboColumn.ValidValues.Add("I", "Material")
            oComboColumn.ValidValues.Add("R", "Resource")
            oComboColumn.ValidValues.Add("E", "Expenses")
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("Status")
            oComboColumn.ValidValues.Add("I", "In Process")
            oComboColumn.ValidValues.Add("P", "Pending")
            oComboColumn.ValidValues.Add("C", "Completed")
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            aform.Items.Item("5").Enabled = False



            aform.Items.Item("5").Enabled = False
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    

    Private Sub Databind_Sort_Production(ByVal aform As SAPbouiCOM.Form, ByVal ItemCode As String, ByVal FormType As String)
        Try
            aform.Freeze(True)
            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            Dim strCode, strPHase, strSeq As String
            ' oCombobox = aform.Items.Item("10").Specific
            '  strCode = oCombobox.Selected.Value
            oCombobox = aform.Items.Item("12").Specific
            strPHase = oCombobox.Selected.Value
            oCombobox = aform.Items.Item("14").Specific
            strSeq = oCombobox.Selected.Value
            Dim st As String
            Dim strSortORder As String
            Dim strItemstring, strPhaseString, strSeqString As String
            Dim strItemstring1, strPhaseString1, strSeqString1 As String


            If strPHase = "Phase" And strSeq = "Sequence" Then
                st = "SELECT T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Phase,T0.U_Z_Seq ,T0.ItemCode"
                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 2
            ElseIf strPHase = "Phase" And strSeq = "" Then
                st = "SELECT T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',T0.[Code] 'ItemCode',  T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.U_Z_Phase,T0.U_Z_Seq ,T0.Code,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Phase,T0.U_Z_Seq ,T0.ItemCode"

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 1
            ElseIf strPHase = "Phase" And strSeq = "Phase" Then
                st = "SELECT T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',T0.[Code] 'ItemCode',  T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.U_Z_Phase,T0.U_Z_Seq ,T0.Code,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Phase,T0.U_Z_Seq ,T0.ItemCode"

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 1
            ElseIf strPHase = "" And strSeq = "Phase" Then
                st = "SELECT T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',T0.[Code] 'ItemCode',  T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.U_Z_Phase,T0.U_Z_Seq ,T0.Code,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Phase,T0.U_Z_Seq ,T0.ItemCode"

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 1
            ElseIf strPHase = "Sequence" And strSeq = "Phase" Then
                st = "SELECT T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase', T0.[Code] 'ItemCode',  T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.Code,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT  T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase',T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.ItemCode"

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 2
            ElseIf strPHase = "Sequence" And strSeq = "" Then
                st = "SELECT T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase', T0.[Code] 'ItemCode',  T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.Code,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase',T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.ItemCode"

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 1
            ElseIf strPHase = "Sequence" And strSeq = "Sequence" Then
                st = "SELECT T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase', T0.[Code] 'ItemCode',  T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.Code,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase',T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.ItemCode"

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 1
            ElseIf strPHase = "" And strSeq = "Sequence" Then
                st = "SELECT T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase', T0.[Code] 'ItemCode',  T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.Code,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase', T0.[ItemCode] 'ItemCode',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.U_Z_Seq ,T0.U_Z_Phase,T0.ItemCode"

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 1
            Else
                st = "SELECT T0.[Code] 'ItemCode',T0.[U_Z_Seq] 'Sequence',T0.[U_Z_Phase] 'Phase',   T0.[Quantity] 'Quanity', T0.[Warehouse] 'Warehouse',Count(*) 'Count' FROM ITT1 T0  INNER JOIN OITT T1 ON T0.Father = T1.Code WHERE T0.[Father]='" & ItemCode & "'"
                st = st & " Group by T0.Code,T0.U_Z_Seq ,T0.U_Z_Phase,T0.[Quantity],T0.[Warehouse]"

                st = "SELECT T0.[ItemCode] 'ItemCode',T0.[U_Z_Phase] 'Phase', T0.[U_Z_Seq] 'Sequence',  Sum(T0.[PlannedQty]) 'Planned Quanity',sum(T0.IssuedQty) 'Issued Qty'  FROM wor1 T0  INNER JOIN OWOR T1 on T1.DocEntry=t0.DocEntry  WHERE T1.[DocNum]=" & ItemCode
                st = st & " Group by T0.ItemCode,T0.U_Z_Phase,T0.U_Z_Seq "

                dtTemp.ExecuteQuery(st)
                oGrid.DataTable = dtTemp
                oGrid.CollapseLevel = 1
            End If
            oGrid.Columns.Item("Phase").TitleObject.Sortable = True
            oGrid.Columns.Item("Sequence").TitleObject.Sortable = True
            oGrid.Columns.Item("ItemCode").TitleObject.Sortable = True
            oEditTextColumn = oGrid.Columns.Item("ItemCode")
            oEditTextColumn.LinkedObjectType = "4"
            ' assignLineNo(aform)
            aform.Items.Item("5").Enabled = False
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

    Private Sub assignLineNo(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        oGrid = aForm.Items.Item("5").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next
        aForm.Freeze(False)
    End Sub

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid, ByVal FormType As String)
        If FormType = "SalesOrder" Then
            agrid.Columns.Item("Code").Visible = False
            agrid.Columns.Item("Name").Visible = False
            agrid.Columns.Item("U_Z_LegNo").TitleObject.Caption = "Leg.No"
            agrid.Columns.Item("U_Z_LegNo").Editable = False
            agrid.Columns.Item("U_Z_Dept").TitleObject.Caption = "Departure"
            agrid.Columns.Item("U_Z_DepDate").TitleObject.Caption = "Dep.Date"
            agrid.Columns.Item("U_Z_ETD").TitleObject.Caption = "ETD"
            agrid.Columns.Item("U_Z_Arrival").TitleObject.Caption = "Arrival"
            agrid.Columns.Item("U_Z_ArvDate").TitleObject.Caption = "Arv.Date"
            agrid.Columns.Item("U_Z_ETA").TitleObject.Caption = "ETA"
            agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
            agrid.Columns.Item("U_Z_RefCode").Visible = False
            agrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        Else
            '   agrid.Columns.Item("RowsHeader").Visible = False
            agrid.Columns.Item("Code").Visible = False
            agrid.Columns.Item("Name").Visible = False
            agrid.Columns.Item("U_Z_LegNo").TitleObject.Caption = "Leg.No"
            agrid.Columns.Item("U_Z_LegNo").Editable = False
            agrid.Columns.Item("U_Z_Dept").TitleObject.Caption = "Departure"
            agrid.Columns.Item("U_Z_Dept").Editable = False
            agrid.Columns.Item("U_Z_DepDate").TitleObject.Caption = "Dep.Date"
            agrid.Columns.Item("U_Z_DepDate").Editable = False
            agrid.Columns.Item("U_Z_ETD").TitleObject.Caption = "ETD"
            agrid.Columns.Item("U_Z_ETD").Editable = False
            agrid.Columns.Item("U_Z_Arrival").TitleObject.Caption = "Arrival"
            agrid.Columns.Item("U_Z_Arrival").Editable = False
            agrid.Columns.Item("U_Z_ArvDate").TitleObject.Caption = "Arv.Date"
            agrid.Columns.Item("U_Z_ArvDate").Editable = False
            agrid.Columns.Item("U_Z_ETA").TitleObject.Caption = "ETA"
            agrid.Columns.Item("U_Z_ETA").Editable = False
            agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
            agrid.Columns.Item("U_Z_RefCode").Visible = False
            agrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        End If
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region


    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("5").Specific
        oEditTextColumn = oGrid.Columns.Item("U_Z_Dept")

        Dim strCode As String
        If oGrid.DataTable.Rows.Count - 1 <= 0 Then
            oGrid.DataTable.Rows.Add()
        End If
        'oEditTextColumn = oGrid.Columns.Item("U_Z_HRPeoobjCode")
        strCode = oEditTextColumn.GetText(oGrid.DataTable.Rows.Count - 1).ToString
        ' strCode = oEditTextColumn.GetTex(oGrid.DataTable.Rows.Count - 1).Value
        If strCode <> "" Then
            oGrid.DataTable.Rows.Add()
            If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End If
        assignLineNo(aForm)
        oGrid.Columns.Item("RowsHeader").Click(oGrid.DataTable.Rows.Count - 1, False)
        oGrid.Columns.Item("U_Z_Dept").Click(oGrid.DataTable.Rows.Count - 1, False)
    End Sub
#Region "DeleteRow"
    Private Sub DeleteRow(ByVal aForm As SAPbouiCOM.Form)

        oGrid = aForm.Items.Item("5").Specific
        Dim strCode As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                oTemp.DoQuery("Update [@Z_CBS1] set Name=Name+'_XD' where Code='" & strCode & "'")
                oGrid.DataTable.Rows.Remove(intRow)
                'If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '    aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                'End If
                assignLineNo(aForm)
                Exit Sub
            End If
        Next
    End Sub

#End Region
    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strSql, strETD, strETA, strDepdate, strArrDate As String
        Dim oRec, oRec1 As SAPbobsCOM.Recordset
        Dim Depdate, Arrivedate As Date
        oGrid = aform.Items.Item("5").Specific
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1

            If oGrid.DataTable.GetValue("U_Z_Dept", intRow) <> "" Then
                strDepdate = oGrid.DataTable.GetValue("U_Z_DepDate", intRow)
                strArrDate = oGrid.DataTable.GetValue("U_Z_ArvDate", intRow)
                strETA = oGrid.DataTable.GetValue("U_Z_ETA", intRow)
                strETD = oGrid.DataTable.GetValue("U_Z_ETD", intRow)
                If strDepdate = "" Then
                    oApplication.Utilities.Message("Enter Departure Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If strETD = "" Then
                    oApplication.Utilities.Message("Enter ETD", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If CInt(strETD) = 0 Then
                    oApplication.Utilities.Message("Enter ETD", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If oGrid.DataTable.GetValue("U_Z_Arrival", intRow) = "" Then
                    oApplication.Utilities.Message("Enter Arrival", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If strArrDate = "" Then
                    oApplication.Utilities.Message("Enter Arrival Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If strETD = "" Then
                    oApplication.Utilities.Message("Enter ETA", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If CInt(strETD) = 0 Then
                    oApplication.Utilities.Message("Enter ETA", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                Dim Etime, Earrive As Integer
                Depdate = oGrid.DataTable.GetValue("U_Z_DepDate", intRow)
                Arrivedate = oGrid.DataTable.GetValue("U_Z_ArvDate", intRow)
                Earrive = oGrid.DataTable.GetValue("U_Z_ETA", intRow)
                Etime = oGrid.DataTable.GetValue("U_Z_ETD", intRow)
                strETA = Earrive.ToString("00:00")
                strETD = Etime.ToString("00:00")
                strSql = "Select * from [@Z_CBS1] where '" & Depdate.ToString("yyyy-MM-dd") & "' = '" & Arrivedate.ToString("yyyy-MM-dd") & "'"
                oRec.DoQuery(strSql)
                If oRec.RecordCount > 0 Then
                    strSql = "Select * from [@Z_CBS1] where '" & Depdate.ToString("yyyy-MM-dd") & " " & strETD & "' < '" & Arrivedate.ToString("yyyy-MM-dd") & " " & strETA & "'"
                    oRec1.DoQuery(strSql)
                    If oRec1.RecordCount = 0 Then
                        oApplication.Utilities.Message("Arrival time should be greater than departure time...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item("U_Z_ETA").Click(intRow)
                        Return False
                    End If
                Else
                    strSql = "Select * from [@Z_CBS1] where '" & Depdate.ToString("yyyy-MM-dd") & "' > '" & Arrivedate.ToString("yyyy-MM-dd") & "'"
                    oRec.DoQuery(strSql)
                    If oRec.RecordCount > 0 Then
                        oApplication.Utilities.Message("Arrival Date should be greater than or equal to departure date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item("U_Z_ArvDate").Click(intRow)
                        Return False
                    End If
                End If
            End If
        Next
        Return True
    End Function
#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim oRec As SAPbobsCOM.Recordset
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc As String
        Dim OCHECKBOXCOLUMN As SAPbouiCOM.CheckBoxColumn
        Dim Etime, Earrive As Integer
        oGrid = aform.Items.Item("5").Specific
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_CBS1")
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If (oGrid.DataTable.GetValue("U_Z_Dept", intRow)) <> "" Then
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                Etime = oGrid.DataTable.GetValue("U_Z_ETD", intRow)
                Earrive = oGrid.DataTable.GetValue("U_Z_ETA", intRow)
                If oUserTable.GetByKey(strCode) Then
                    Dim str As String = Etime.ToString("00:00")
                    Dim strArrive As String = Earrive.ToString("00:00")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_Dept").Value = (oGrid.DataTable.GetValue("U_Z_Dept", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_DepDate").Value = (oGrid.DataTable.GetValue("U_Z_DepDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETD").Value = str ' (oGrid.DataTable.GetValue("U_Z_ETD", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_Arrival").Value = (oGrid.DataTable.GetValue("U_Z_Arrival", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ArvDate").Value = (oGrid.DataTable.GetValue("U_Z_ArvDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETA").Value = strArrive ' (oGrid.DataTable.GetValue("U_Z_ETA", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                    If oUserTable.Update() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Else
                    'Dim int As Integer = 320
                    Dim str As String = Etime.ToString("00:00")
                    Dim strArrive As String = Earrive.ToString("00:00")
                    strCode = oApplication.Utilities.getMaxCode("@Z_CBS1", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_Dept").Value = (oGrid.DataTable.GetValue("U_Z_Dept", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_DepDate").Value = (oGrid.DataTable.GetValue("U_Z_DepDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETD").Value = str '(oGrid.DataTable.GetValue("U_Z_ETD", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_Arrival").Value = (oGrid.DataTable.GetValue("U_Z_Arrival", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ArvDate").Value = (oGrid.DataTable.GetValue("U_Z_ArvDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETA").Value = strArrive ' (oGrid.DataTable.GetValue("U_Z_ETA", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                    If oUserTable.Add() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
        Next
        oRec.DoQuery("Delete from [@Z_CBS1] where Name like '%_XD' and U_Z_RefCode='" & oApplication.Utilities.getEdittextvalue(aform, "7") & "'")
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Databind(aform)
    End Function
#End Region

    Private Sub Committrans(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String)
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Update [@Z_CBS1] set Name=Code where Name like '%_XD' and U_Z_RefCode='" & oApplication.Utilities.getEdittextvalue(aForm, "7") & "'")
    End Sub



#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_MyTasks Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "2" Then
                                    '  Committrans(oForm, "Cancel")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oGrid = oForm.Items.Item("5").Specific

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "25" Then
                                    FillValueCombo(oForm)
                                End If
                                If pVal.ItemUID = "20" Then
                                    Databind_Filter(oForm)
                                End If
                                If pVal.ItemUID = "22" Then
                                    Databind_Sort(oForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                


                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        oForm.Update()
                                        'If pVal.ItemUID = "5" Then
                                        '    oGrid = oForm.Items.Item("5").Specific
                                        '    val = oDataTable.GetValue("FormatCode", 0)
                                        '    Try

                                        '        oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        '    Catch ex As Exception
                                        '    End Try
                                        'End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try
                        End Select
                End Select
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID

                Case mnu_MyTasks
                    LoadForm()
                Case mnu_DELETE_ROW
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                'Select Case pVal.MenuUID
                '    Case mnu_LeaveMaster
                '        oMenuobject = New clsEarning
                '        oMenuobject.MenuEvent(pVal, BubbleEvent)
                'End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
