Imports System.IO
Public Class clsUtilities

    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub
    Public Function GetEmployeeID(ByVal aCode As String) As String
        Dim strEmpID As String = ""
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Return strEmpID
    End Function
    Public Function getManagerEmPID(ByVal aCode As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        Dim intManagerid As Integer
        Dim strEmp As String = ""
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("SELECT *  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & aCode & "'")
        If oTest.RecordCount > 0 Then
            intManagerid = oTest.Fields.Item("empID").Value
            Return intManagerid.ToString
        Else
            Return "0"
        End If
    End Function
    Public Function GetEmployeeID_Lists(ByVal aCode As String) As String
        Dim strEmpID As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    End Function
    Public Function getEmpIDforMangers(ByVal aCode As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        Dim intManagerid As Integer
        Dim strEmp As String = ""
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("SELECT *  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & aCode & "'")
        If oTest.RecordCount > 0 Then
            intManagerid = oTest.Fields.Item("empID").Value
            ' strEmp = "'" & intManagerid & "'"
            oTest.DoQuery("Select empId from OHEM where manager=" & intManagerid)
            For intRow As Integer = 0 To oTest.RecordCount - 1

                If strEmp = "" Then
                    strEmp = "'" & oTest.Fields.Item(0).Value & "'"
                Else
                    strEmp = strEmp & " ,'" & oTest.Fields.Item(0).Value & "'"
                End If
                oTest.MoveNext()
            Next
            Return strEmp
        Else
            Return ""
        End If
    End Function

    Public Function getEmpIDforMangers_Reports(ByVal aCode As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        Dim intManagerid As Integer
        Dim strEmp As String = ""
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("SELECT *  FROM OHEM T0  INNER JOIN OUSR T1 ON T0.userId = T1.USERID WHERE T1.[USER_CODE] ='" & aCode & "'")
        If oTest.RecordCount > 0 Then
            intManagerid = oTest.Fields.Item("empID").Value
            strEmp = "'" & intManagerid & "'"
            oTest.DoQuery("Select empId from OHEM where manager=" & intManagerid)
            For intRow As Integer = 0 To oTest.RecordCount - 1

                If strEmp = "" Then
                    strEmp = "'" & oTest.Fields.Item(0).Value & "'"
                Else
                    strEmp = strEmp & " ,'" & oTest.Fields.Item(0).Value & "'"
                End If
                oTest.MoveNext()
            Next
            Return strEmp
        Else
            Return ""
        End If
    End Function

    Private Function LogonLoginDetails(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select U_UID,U_Pwd, U_EmpiD,isnull(U_Approver,'N') from [@Z_Login] where U_UID='" & oApplication.Company.UserName & "'")
        If otemp.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(aForm, "3", otemp.Fields.Item(0).Value)
            oApplication.Utilities.setEdittextvalue(aForm, "5", otemp.Fields.Item(1).Value)
            Return True
        Else
            oApplication.Utilities.setEdittextvalue(aForm, "3", "")
            oApplication.Utilities.setEdittextvalue(aForm, "5", "")
            Return False
        End If



    End Function
    Public Function validateAuthorization(ByVal aUserId As String, ByVal aFormUID As String) As Boolean
        Dim oAuth As SAPbobsCOM.Recordset
        oAuth = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim struserid As String
        struserid = oApplication.Company.UserName
        oAuth.DoQuery("select * from UPT1 where FormId='" & aFormUID & "'")
        If (oAuth.RecordCount <= 0) Then
            Return True
        Else
            Dim st As String
            st = oAuth.Fields.Item("PermId").Value
            st = "Select * from USR3 where PermId='" & st & "' and UserLink=" & aUserId
            oAuth.DoQuery(st)
            If oAuth.RecordCount > 0 Then
                If oAuth.Fields.Item("Permission").Value = "N" Then
                    Return False
                End If
                Return True
            Else
                Return True
            End If

        End If

        Return True

    End Function

#Region "Validate Record Exists"
    Public Function validaterecordexists1(ByVal strValue As String, ByVal strField As String, ByVal strtable As String) As Boolean
        Dim oRec As SAPbobsCOM.Recordset
        Dim strquery As String
        strquery = "Select * from [@" & strtable & "] where " & strField & "='" & strValue.Replace("'", "''") & "'"
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery(strquery)
        If oRec.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function validaterecordexists_Activity(ByVal strValue As String, ByVal strField As String, ByVal strtable As String, ByVal strPhase As String) As Boolean
        Dim oRec As SAPbobsCOM.Recordset
        Dim strquery As String
        strquery = "Select * from [@" & strtable & "] where " & strField & "='" & strValue.Replace("'", "''") & "' and U_Z_MODNAME='" & strPhase.Replace("'", "''") & "'"
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery(strquery)
        If oRec.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If



    End Function
#End Region

#Region "Get item Codes from marketing Documents"
    Public Function getItemCodesFromDocuments(ByVal aForm As SAPbouiCOM.Form) As String
        Dim strDocumentItemCodes As String = ""
        Dim strItem As String
        Dim oMatrix As SAPbouiCOM.Matrix
        oMatrix = aForm.Items.Item("38").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            strItem = getMatrixValues(oMatrix, "1", intRow)
            If strItem <> "" Then
                If strDocumentItemCodes = "" Then
                    strDocumentItemCodes = "'" & strItem & "'"
                Else
                    strDocumentItemCodes = strDocumentItemCodes & ",'" & strItem & "'"
                End If
            End If
        Next
        strDocumentItemCodes = "(" & strDocumentItemCodes & ")"
        strSPItemCode = strDocumentItemCodes
        Return strDocumentItemCodes
    End Function
#End Region

#Region "GetLoggedInUser"
    Public Function GetLoggedUserName() As String
        Dim strUser As String
        Dim ouserRs As SAPbobsCOM.Recordset
        ouserRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strUser = oApplication.Company.UserName
        '  ouserRs.DoQuery("Select * from OUSR where Userid='" & strUser & "'")
        Return strUser

    End Function
#End Region

#Region "Display Condition Records"

    Private Function CheckRecordExists(ByVal aCOn As String, ByVal aCon1 As String, ByVal aCon2 As String) As Boolean
        Dim otemprec As SAPbobsCOM.Recordset
        otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemprec.DoQuery(aCOn)
        If otemprec.RecordCount > 0 Then
            Return True
        End If
        otemprec.DoQuery(aCon1)
        If otemprec.RecordCount > 0 Then
            Return True
        End If
        otemprec.DoQuery(aCon2)
        If otemprec.RecordCount > 0 Then
            Return True
        End If
        Return False
    End Function

    


#End Region

#Region "Check Module/Activity"
    Public Function CheckModule_Activity(ByVal strproject As String, ByVal strTable As String, ByVal strValue As String, ByVal strField As String) As Boolean
        Dim oTe As SAPbobsCOM.Recordset
        Dim strQry As String
        oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strQry = "Select * from " & strTable & " where " & strField & "='" & strValue.ToUpper.Replace("'", "''") & "' and docentry=(Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & strproject.Replace("'", "''") & "')"
        oTe.DoQuery(strQry)
        If oTe.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function CheckActivity(ByVal strproject As String, ByVal strTable As String, ByVal strValue As String, ByVal strvalue1 As String, ByVal strfield As String, ByVal strField1 As String) As Boolean
        Dim oTe As SAPbobsCOM.Recordset
        Dim strQry As String
        oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strQry = "Select * from " & strTable & " where " & strfield & "='" & strValue.ToUpper.Replace("'", "''") & "' and " & strField1 & "='" & strvalue1.Replace("'", "''") & "' and docentry=(Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & strproject.Replace("'", "''") & "')"
        oTe.DoQuery(strQry)
        If oTe.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region


#Region "Fill ComboBoxValues"
    Public Sub FillComboBoxColumn(ByVal aCombo As SAPbouiCOM.ComboBoxColumn, ByVal sql As String)
        Dim oComborec As SAPbobsCOM.Recordset
        oComborec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oComborec.DoQuery(sql)
        For introw As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
            aCombo.ValidValues.Remove(introw)
        Next
        aCombo.ValidValues.Add("", "")
        For introw As Integer = 0 To oComborec.RecordCount - 1
            aCombo.ValidValues.Add(oComborec.Fields.Item(0).Value, oComborec.Fields.Item(1).Value)
            oComborec.MoveNext()
        Next


    End Sub

    Public Sub FillComboBox(ByVal aCombo As SAPbouiCOM.ComboBox, ByVal sql As String)
        Dim oComborec As SAPbobsCOM.Recordset
        oComborec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oComborec.DoQuery(sql)
        For introw As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
            aCombo.ValidValues.Remove(introw)
        Next
        aCombo.ValidValues.Add("", "")
        For introw As Integer = 0 To oComborec.RecordCount - 1
            Try
                aCombo.ValidValues.Add(oComborec.Fields.Item(0).Value, oComborec.Fields.Item(1).Value)
            Catch ex As Exception
            End Try

            oComborec.MoveNext()
        Next
        aCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub
#End Region

#Region "Copy Files"
    Public Sub CopyFilestoCustomers(ByVal aFileName As String, ByVal aLogPath As String)
        Dim otemp As SAPbobsCOM.Recordset
        Dim strFilePath, strDesgfilename, strMessage As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from OCRD where cardtype='C' and U_PharmInt = 'Y'")
        strFilePath = "C:\MYDATA"
        For intRow As Integer = 0 To otemp.RecordCount - 1
            strFilePath = strFilePath & "\" & otemp.Fields.Item("CardCode").Value
            If Directory.Exists(strFilePath) Then
            Else
                Directory.CreateDirectory(strFilePath)
            End If
            strDesgfilename = strFilePath & "\PROMFLQ.mfp"
            If File.Exists(strDesgfilename) Then
                File.Delete(strDesgfilename)
            End If
            File.Copy(aFileName, strDesgfilename)
            '  strFilePath = strExportFilePaty
            '    strMessage = "Exported :  File name : " & strDesgfilename
            ''WriteErrorlog(strMessage, aLogPath)
            otemp.MoveNext()
        Next


    End Sub
#End Region

#Region "Check the Company Settings"
    Public Sub CheckCompanySettings()
        Dim otempRs As SAPbobsCOM.Recordset
        otempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otempRs.DoQuery("Select isnull(U_MasExport,'N'),isnull(U_JEExport,'N') from OADM")
        If otempRs.Fields.Item(0).Value = "Y" Then
            blnMasterExport = True
        Else
            blnMasterExport = False
        End If
        If otempRs.Fields.Item(1).Value = "Y" Then
            blnFEExport = True
        Else
            blnFEExport = False
        End If
    End Sub
#End Region

#Region "Add Controls"

    '*****************************************************************
    'Type               : Procedure   
    'Name               : addControls
    'Parameter          : StrCode
    'Return Value       : string
    'Author             : Senthil Kumar B
    'Created Date       : 03-07-2009
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create Controls in the SAP B1 Screens
    '*****************************************************************
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "")
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                    .Height = objOldItem.Height

                ElseIf position.ToUpper = "DOWN" Then
                    .Top = objOldItem.Top + objOldItem.Height + 1
                    .Left = objOldItem.Left
                ElseIf position.ToUpper = "TOP" Then
                    .Top = objOldItem.Top - objOldItem.Height - 3
                    .Left = objOldItem.Left
                End If
            End If
            '.FromPane = fromPane
            '.ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            ' .ForeColor = 255
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            If ItemUID = "btnDisplay" Then
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 60
            Else
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 60
            End If
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption
            objNewItem.Width = 150
        End If
    End Sub
#End Region

#Region "validate onHandqty"
    Private Function validateOnhand(ByVal aItemCode As String, ByVal aWhs As String, ByVal dblqty As Double) As Boolean
        Dim oTempRs As SAPbobsCOM.Recordset
        Dim dblOnHand As Double
        oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRs.DoQuery("Select * from OITW where itemcode='" & aItemCode & "' and whscode='" & aWhs & "'")
        dblOnHand = 0
        If oTempRs.RecordCount > 0 Then
            dblOnHand = oTempRs.Fields.Item("OnHand").Value - oTempRs.Fields.Item("MinStock").Value
        End If
        If dblOnHand >= dblqty Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

    Public Function CheckSurcharge(ByVal aCardcode As String, ByVal aFieldName As String) As Boolean
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select isnull(" & aFieldName & ",'N') from OCRD where cardcode='" & aCardcode & "'")
        If otemp.Fields.Item(0).Value.ToString.ToUpper() = "YES" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetLocalCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select Maincurncy from OADM"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function GetSystemCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select SysCurrncy from OADM"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function getBPCurrency(ByVal strCardcode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select Currency from OCRD where Cardcode='" & strCardcode & "'"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    Public Function getDateStrin(ByVal DateString As Date) As String
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_DateToString(DateString).Fields.Item(0).Value
    End Function

#Region "Validate Record Exists"
    Public Function validaterecordexists(ByVal strValue As String, ByVal strField As String, ByVal strtable As String) As Boolean
        Dim oRec As SAPbobsCOM.Recordset
        Dim strquery As String
        strquery = "Select * from [@" & strtable & "] where " & strField & "='" & strValue & "'"
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery(strquery)
        If oRec.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If



    End Function
#End Region

#Region "SetDatabind"
    Public Sub setUserDatabind(ByVal aForm As SAPbouiCOM.Form, ByVal UID As String, ByVal strDBID As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aForm.Items.Item(UID).Specific
        objEdit.DataBind.SetBound(True, "", strDBID)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String
        Return Application.StartupPath.Trim
        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "valiate superuser"
    Public Function validateSuperuser() As Boolean
        Dim oTempRs As SAPbobsCOM.Recordset
        Dim oUser, strString As String
        oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUser = oApplication.Company.UserName
        strString = "SELECT *  FROM [OUSR] T0 where isnull(U_Superuser,'N')='Y' and  USER_CODE='" & oUser & "'"
        oTempRs.DoQuery(strString)
        If oTempRs.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If


    End Function

#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

#Region "Assign Fright Expances Details"

    Public Function validateSurchargerequeired(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strCardCode, strDocType As String
        Dim oCombo As SAPbouiCOM.ComboBox
        oCombo = aForm.Items.Item("3").Specific
        strDocType = oCombo.Selected.Value
        If strDocType <> "I" Then
            Return False
        End If
        strCardCode = getEdittextvalue(aForm, "4")

        If CheckSurcharge(strCardCode, "U_Z_SHAK_FLAG") = True Then
            Return True
        ElseIf CheckSurcharge(strCardCode, "U_Z_APOT_FLAG") = True Then
            Return True
        Else
            Return False
        End If
        Return False

    End Function

    Public Sub CalculateSurcharges(ByVal aForm As SAPbouiCOM.Form)
        Dim oTempForm1, oTempForm2 As SAPbouiCOM.Form
        Dim Frtmtrx As SAPbouiCOM.Matrix
        Dim strCardCode, strCurrencyQuery, strCurrency, strFrightValue, strDate, strSurName, strRemarks, strDocBeftotal, strDiscount, strVatcode As String
        Dim dblVatAmount, dblFixedAmount, dblSurPer, dblVatPer, dbLDocDefTotal, dblDiscount As Double
        Dim dtDocdate As Date
        Dim otempRec, otemprs, oSurRecord As SAPbobsCOM.Recordset
        Dim w As Integer
        Dim oCombo As SAPbouiCOM.ComboBox

        If aForm.Type = frm_ARInvoice Or aForm.Type = frm_ARCreditNote Then
            If validateSurchargerequeired(aForm) = False Then
                Exit Sub
            End If
            aSourceForm = aForm
            oCombo = aSourceForm.Items.Item("70").Specific
            strCurrency = oCombo.Selected.Value
            strCardCode = getEdittextvalue(aForm, "4")
            strDocBeftotal = getEdittextvalue(aSourceForm, "22")
            strDiscount = getEdittextvalue(aSourceForm, "42")
            If strDocBeftotal = "" And strDiscount = "" Then
                Exit Sub
            End If
            otemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strCurrencyQuery = ""
            Select Case strCurrency
                Case "C"
                    strCurrencyQuery = "Select Currency from OCRD where Cardcode='" & strCardCode & "'"
                Case "L"
                    strCurrencyQuery = "Select MainCurncy from OADM"
                Case "S"
                    strCurrencyQuery = "Select SysCurncy from OADM"
            End Select
            If strCurrencyQuery <> "" Then
                otemprs.DoQuery(strCurrencyQuery)
                strCurrency = otemprs.Fields.Item(0).Value
            Else
                strCurrency = ""
            End If

            If strDocBeftotal.Length > 3 Then
                If strCurrency <> "##" Then
                    strDocBeftotal = strDocBeftotal.Replace(strCurrency, "")
                Else
                    strDocBeftotal = strDocBeftotal.Substring(3)
                End If
            End If

            If strDiscount.Length > 3 Then
                If strCurrency <> "##" Then
                    strDiscount = strDiscount.Replace(strCurrency, "")
                Else
                    strDiscount = strDiscount.Substring(3)
                End If
            End If
            If strDocBeftotal <> "" Then
                dbLDocDefTotal = CDbl(strDocBeftotal)
            Else
                dbLDocDefTotal = 0
            End If
            If strDiscount <> "" Then
                dblDiscount = CDbl(strDiscount)
            Else
                dblDiscount = 0
            End If


            strDate = getEdittextvalue(aForm, "10")
            dtDocdate = GetDateTimeValue(strDate)
            otempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSurRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempForm1 = aForm
            'oTempForm1.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oTempForm2 = oApplication.SBO_Application.Forms.GetForm("3007", 1)  '// Freight Screen
            oTempForm2 = oApplication.SBO_Application.Forms.ActiveForm()
            If oTempForm2.Type <> 3007 Then
                Exit Sub
            End If
            Dim strFrightName As String
            Try
                oTempForm2.Freeze(True)
                Dim strSQL As String
                If CheckSurcharge(strCardCode, "U_Z_SHAK_FLAG") = True Then
                    strSQL = "Select * from [@Z_SURCHARGES] where U_Z_SUR_BPNM='U_Z_SHAK_FLAG' and '" & dtDocdate.ToString("yyyy-MM-dd") & " '  between U_Z_SUR_FRMDATE and isnull(U_Z_SUR_TODATE,dateadd(m,40,getdate())) order by U_Z_SUR_FRMDATE Desc, U_Z_SUR_TODATE Desc"
                    oSurRecord.DoQuery(strSQL)
                    dblFixedAmount = 0
                    dblVatAmount = 0
                    strRemarks = ""
                    dblVatPer = 0
                    dblSurPer = 0
                    If oSurRecord.RecordCount > 0 Then
                        Frtmtrx = oTempForm2.Items.Item("3").Specific
                        w = 1
                        strVatcode = oSurRecord.Fields.Item("U_Z_SUR_VAT").Value
                        dblVatPer = oSurRecord.Fields.Item("U_Z_SUR_VATPER").Value
                        dblSurPer = oSurRecord.Fields.Item("U_Z_SUR_PER").Value
                        strRemarks = oSurRecord.Fields.Item("U_Z_SUR_REM").Value
                        dblFixedAmount = ((dbLDocDefTotal - dblDiscount) * dblSurPer) / 100
                        dblVatAmount = dblFixedAmount * dblVatPer / 100
                        strSurName = oSurRecord.Fields.Item("U_Z_SUR_NAME").Value
                        strFrightName = ""
                        While w <= Frtmtrx.RowCount
                            Try
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.selected.description
                            Catch ex As Exception
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.value
                                otempRec.DoQuery("SELECT Expnscode,Expnsname  FROM OEXD T0 where ExpnsCode=" & strFrightName)
                                If otempRec.RecordCount > 0 Then
                                    strFrightName = otempRec.Fields.Item(1).Value
                                End If
                            End Try


                            If strFrightName.ToUpper = strSurName.ToUpper Then '//AD
                                Frtmtrx.Columns.Item("2").Cells.Item(w).Specific.value = strRemarks
                                Frtmtrx.Columns.Item("3").Cells.Item(w).Specific.value = dblFixedAmount
                                Try
                                    oCombo = Frtmtrx.Columns.Item("11").Cells.Item(w).Specific
                                    oCombo.Select(strVatcode, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    '    Frtmtrx.Columns.Item("12").Cells.Item(w).Specific.value = oSurRecord.Fields.Item("U_Z_SUR_VATPER").Value

                                    '    Frtmtrx.Columns.Item("17").Cells.Item(w).Specific.value = strVatcode
                                Catch ex As Exception

                                End Try

                                Exit While
                            End If
                            w = w + 1
                        End While
                    End If
                End If
                dblFixedAmount = 0
                dblVatAmount = 0
                strRemarks = ""
                dblVatPer = 0
                dblSurPer = 0
                If CheckSurcharge(strCardCode, "U_Z_APOT_FLAG") = True Then
                    strSQL = "Select * from [@Z_SURCHARGES] where U_Z_SUR_BPNM='U_Z_APOT_FLAG' and '" & dtDocdate.ToString("yyyy-MM-dd") & " '  between U_Z_SUR_FRMDATE and isnull(U_Z_SUR_TODATE,dateadd(m,40,getdate())) order by U_Z_SUR_FRMDATE Desc, U_Z_SUR_TODATE desc"
                    oSurRecord.DoQuery(strSQL)
                    If oSurRecord.RecordCount > 0 Then
                        Frtmtrx = oTempForm2.Items.Item("3").Specific
                        strVatcode = oSurRecord.Fields.Item("U_Z_SUR_VAT").Value
                        dblVatPer = oSurRecord.Fields.Item("U_Z_SUR_VATPER").Value
                        dblSurPer = oSurRecord.Fields.Item("U_Z_SUR_PER").Value
                        strRemarks = oSurRecord.Fields.Item("U_Z_SUR_REM").Value
                        dblFixedAmount = ((dbLDocDefTotal - dblDiscount) * dblSurPer) / 100
                        dblVatAmount = dblFixedAmount * dblVatPer / 100
                        w = 1
                        strSurName = oSurRecord.Fields.Item("U_Z_SUR_NAME").Value
                        strFrightName = ""
                        While w <= Frtmtrx.RowCount
                            '  If Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.selected.description = strSurName Then '//AD
                            Try
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.selected.description
                            Catch ex As Exception
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.value
                                otempRec.DoQuery("SELECT Expnscode,Expnsname  FROM OEXD T0 where ExpnsCode=" & strFrightName)
                                If otempRec.RecordCount > 0 Then
                                    strFrightName = otempRec.Fields.Item(1).Value
                                End If
                            End Try
                            If strFrightName.ToUpper = strSurName.ToUpper Then '//AD
                                Frtmtrx.Columns.Item("2").Cells.Item(w).Specific.value = strRemarks
                                Frtmtrx.Columns.Item("3").Cells.Item(w).Specific.value = dblFixedAmount
                                Try
                                    oCombo = Frtmtrx.Columns.Item("11").Cells.Item(w).Specific
                                    oCombo.Select(strVatcode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                    '   Frtmtrx.Columns.Item("17").Cells.Item(w).Specific.value = strVatcode
                                Catch ex As Exception

                                End Try

                                Exit While
                            End If
                            w = w + 1
                        End While
                    End If
                End If
                oTempForm2.Freeze(False)
                If blnDocumentItem = True Then
                    If oTempForm2.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        oTempForm2.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oTempForm2.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Else
                        oTempForm2.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oTempForm2.Freeze(False)
            End Try
            blnDocumentItem = False
        End If
    End Sub
#End Region


#Region "Check Condition Type"
    Public Function CheckConditionType(ByVal aCode As String) As String
        Dim oCheckRs As SAPbobsCOM.Recordset
        oCheckRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCheckRs.DoQuery("Select isnull(U_Z_COND_TYPE,'') from [@Z_CONDITIONS] where U_Z_COND_NAME='" & aCode & "'")
        Return (oCheckRs.Fields.Item(0).Value)
    End Function
#End Region

#Region "calculate Discount"

#Region "Check COndition Status"
    Private Function CheckConditionStatus(ByVal aCode As String) As Boolean
        Dim Temprec As SAPbobsCOM.Recordset
        Temprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Temprec.DoQuery("Select * from [@Z_CONDITIONS] where U_Z_COND_CODE='" & aCode & "'")
        If Temprec.RecordCount > 0 Then
            If Temprec.Fields.Item("U_Z_COND_STATUS").Value = "Y" Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function
#End Region

    Private Function getQuantityDiscount(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aCardCode As String, ByVal aItemGroup As String, ByVal aDate As Date, ByVal LineQty As Double) As Double
        Dim oQtyRS, oItemRS As SAPbobsCOM.Recordset
        Dim strQtyRS, strItemCode, strLineItemGroup, strMainItemGroup As String
        Dim dblDis, dblScalqty, dblLineQty, dblScaleToQty As Double
        oItemRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        dblLineQty = 0

        strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET] FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
        strQtyRS = strQtyRS & " where T0.[U_Z_DISC_CCODE]='" & aCardCode & "' and T0.[U_Z_DISC_ICODE]='" & aItemGroup & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code)"
        oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oQtyRS.DoQuery(strQtyRS)
        If oQtyRS.RecordCount > 0 Then
            strMainItemGroup = oQtyRS.Fields.Item(0).Value
        Else
            Return dblLineQty
        End If
        For IntLoop As Integer = 1 To aMatrix.RowCount
            strItemCode = getMatrixValues(aMatrix, "1", IntLoop)
            If strItemCode <> "" Then
                'oItemRS.DoQuery("Select * from OITM where Itemcode='" & strItemCode & "'")
                strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET] FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
                strQtyRS = strQtyRS & " where T0.[U_Z_DISC_CCODE]='" & aCardCode & "' and T0.[U_Z_DISC_ICODE]='" & strItemCode & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code)"
                oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oQtyRS.DoQuery(strQtyRS)
                oItemRS.DoQuery(strQtyRS)
                If oItemRS.RecordCount > 0 Then
                    strLineItemGroup = oItemRS.Fields.Item(0).Value
                    If strMainItemGroup = strLineItemGroup Then
                        dblLineQty = dblLineQty + getDocumentQuantity(getMatrixValues(aMatrix, "11", IntLoop))
                    End If
                End If
            End If
        Next
        LineQty = dblLineQty
        strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET],isnull(T0.[U_Z_QTY_TO_SCALE],0) 'ScaleTo' FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
        ' strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET],isnull(T0.[U_Z_QTY_TO_SCALE]," & LineQty & ") 'ScaleTo' FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
        strQtyRS = strQtyRS & " where T0.[U_Z_DISC_CCODE]='" & aCardCode & "' and T0.[U_Z_DISC_ICODE]='" & aItemGroup & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code) Desc"
        oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oQtyRS.DoQuery(strQtyRS)
        dblDis = 0
        If oQtyRS.RecordCount > 0 Then
            'oQtyRS.MoveLast()
            For intRow As Integer = 0 To oQtyRS.RecordCount - 1
                If CheckConditionStatus(oQtyRS.Fields.Item(0).Value) = True Then
                    dblScalqty = oQtyRS.Fields.Item("U_Z_QTY_FROM_SCALE").Value
                    '  dblScaleToQty = oQtyRS.Fields.Item("ScaleTo").Value
                    'If dblScaleToQty = 0 Then
                    '    dblScaleToQty = LineQty
                    'End If
                    If dblScalqty <= LineQty Then
                        ' If LineQty <= dblScaleToQty And LineQty >= dblScalqty Then
                        dblDis = oQtyRS.Fields.Item("U_Z_QTY_SCALE_DIS").Value
                        Exit For
                    End If
                End If
                oQtyRS.MoveNext()
            Next
        End If
        If dblDis = 0 Then
            oQtyRS.DoQuery(strQtyRS)
            If oQtyRS.RecordCount > 0 Then
                oQtyRS.MoveLast()
                'dblDis = oQtyRS.Fields.Item("U_Z_QTY_SCALE_DIS").Value
                dblDis = 0
            End If
        End If
        Return dblDis
    End Function

    Private Function getConditiongroup(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aCardCode As String, ByVal aItemGroup As String, ByVal aDate As Date) As Double
        Dim oQtyRS, oItemRS As SAPbobsCOM.Recordset
        Dim strQtyRS, strItemCode, strLineItemGroup, strMainItemGroup As String
        Dim dblDis, dblScalqty, dblLineQty, dblScaleToQty As Double
        oItemRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        dblLineQty = 0
        oItemRS.DoQuery("Select * from OCRD where Cardcode='" & aCardCode & "'")
        aCardCode = oItemRS.Fields.Item("GroupCode").Value
        strQtyRS = "SELECT T0.[U_Z_DISC_LINK],  T0.[U_Z_DISC_PERC], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET] FROM [dbo].[@Z_DISC_BP_ITM_GROUP]  T0 "
        strQtyRS = strQtyRS & " where T0.[U_Z_DISC_BP_GROUP]='" & aCardCode & "' and T0.[U_Z_DISC_ITM_GROUP]='" & aItemGroup & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code)"
        oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oQtyRS.DoQuery(strQtyRS)
        dblDis = 0
        If oQtyRS.RecordCount > 0 Then
            'oQtyRS.MoveLast()
            For intRow As Integer = 0 To oQtyRS.RecordCount - 1
                If CheckConditionStatus(oQtyRS.Fields.Item(0).Value) = True Then
                    dblDis = dblDis + oQtyRS.Fields.Item("U_Z_DISC_PERC").Value
                End If
                oQtyRS.MoveNext()
            Next
        End If
        Return dblDis
    End Function
    Public Sub CalculateDiscount(ByVal aForm As SAPbouiCOM.Form)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim strCardCode, strItemCode, strSQL, strTempQuery, strConditionCode, strConditionQuery, strPostingdate As String
        Dim dtPostingdate As Date
        Dim oItemRs, oConditionGroup, oConditionType, oTempRs As SAPbobsCOM.Recordset
        Try
            Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            aForm.Freeze(True)
            oItemRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oConditionGroup = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oConditionType = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = aForm.Items.Item("38").Specific
            strCardCode = getEdittextvalue(aForm, "4")
            strPostingdate = getEdittextvalue(aForm, "10")
            If strPostingdate <> "" Then
                dtPostingdate = GetDateTimeValue(strPostingdate)
            End If
            oCombo = aForm.Items.Item("3").Specific
            strTempQuery = ""

            'Condition Type Discount for Price 

            strTempQuery = "SELECT *  FROM [dbo].[@Z_DISCOUNT_GROUP]  T0 inner join  [dbo].[@Z_CONDITIONS]  "
            ' strTempQuery = strTempQuery & " T1 on T0.U_Z_Disc_Link=T1.U_Z_COND_CODE and T1.U_Z_COND_STATUS='Y' where U_Z_DISC_CCODE='" & strCardCode & "'"
            strTempQuery = strTempQuery & " T1 on T0.U_Z_Disc_Link=T1.U_Z_COND_CODE   where U_Z_DISC_CCODE='" & strCardCode & "'"
            strTempQuery = strTempQuery & "  and ('" & dtPostingdate.ToString("yyyy-MM-dd") & "' between U_Z_DISC_DATEF and U_Z_DISC_DATET )"
            oConditionType.DoQuery(strTempQuery)
            Dim dblCumdiscount, dblDiscount As Double
            If oConditionType.RecordCount > 0 Then
                If strCardCode <> "" And oCombo.Selected.Value = "I" And strPostingdate <> "" Then
                    For intRow As Integer = 1 To oMatrix.RowCount
                        Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        strItemCode = getMatrixValues(oMatrix, "1", intRow)
                        dblCumdiscount = 0
                        dblDiscount = 0
                        If strItemCode <> "" Then
                            strSQL = ""
                            strSQL = strTempQuery & " and U_Z_DISC_ICODE='" & strItemCode & "'"
                            oTempRs.DoQuery(strSQL)
                            If oTempRs.RecordCount > 0 Then
                                If oTempRs.Fields.Item("U_Z_COND_STATUS").Value = "Y" Then
                                    For intLoop As Integer = 0 To oTempRs.RecordCount - 1
                                        dblDiscount = 0
                                        dblDiscount = oTempRs.Fields.Item("U_Z_DISC_PERC").Value
                                        dblCumdiscount = dblCumdiscount + dblDiscount
                                        oTempRs.MoveNext()
                                    Next
                                End If
                                Dim dbllinelineqty, dblScalediscount As Double
                                SetMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intRow, dblCumdiscount)
                                dblCumdiscount = dblCumdiscount + dblScalediscount
                                dblCumdiscount = dblCumdiscount
                                SetMatrixValues(oMatrix, "15", intRow, dblCumdiscount.ToString)
                            End If
                        End If
                    Next
                End If
            Else
                For intLoop As Integer = 1 To oMatrix.RowCount
                    strItemCode = getMatrixValues(oMatrix, "1", intLoop)
                    If strItemCode <> "" Then
                        dblCumdiscount = 0
                        SetMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intLoop, dblCumdiscount)
                    End If
                Next
            End If

            'Condition Group Discount 
            If strCardCode <> "" And oCombo.Selected.Value = "I" And strPostingdate <> "" Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strItemCode = getMatrixValues(oMatrix, "1", intRow)
                    dblCumdiscount = 0
                    dblDiscount = 0
                    If strItemCode <> "" Then
                        strSQL = ""
                        oItemRs.DoQuery("Select * from OITM where Itemcode='" & strItemCode & "'")
                        Dim dblScalediscount As Double
                        dblCumdiscount = getDocumentQuantity(getMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intRow))
                        dblScalediscount = getConditiongroup(oMatrix, strCardCode, oItemRs.Fields.Item("ItmsGrpCod").Value, dtPostingdate)
                        'dblScalediscount = dblCumdiscount + dblScalediscount
                        SetMatrixValues(oMatrix, "U_Z_DISCOUNT_GROUP", intRow, dblScalediscount)
                    End If
                Next
            End If



            'Condition Type Discount for Scales
            If strCardCode <> "" And oCombo.Selected.Value = "I" And strPostingdate <> "" Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strItemCode = getMatrixValues(oMatrix, "1", intRow)
                    dblCumdiscount = 0
                    dblDiscount = 0
                    If strItemCode <> "" Then
                        strSQL = ""
                        oItemRs.DoQuery("Select * from OITM where Itemcode='" & strItemCode & "'")
                        Dim dbllinelineqty, dblScalediscount, dblGroupPrice As Double
                        dbllinelineqty = getDocumentQuantity(getMatrixValues(oMatrix, "11", intRow))
                        dblCumdiscount = getDocumentQuantity(getMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intRow))
                        dblGroupPrice = getDocumentQuantity(getMatrixValues(oMatrix, "U_Z_DISCOUNT_GROUP", intRow))
                        dblScalediscount = getQuantityDiscount(oMatrix, strCardCode, oItemRs.Fields.Item("ItemCode").Value, dtPostingdate, dbllinelineqty)
                        SetMatrixValues(oMatrix, "U_Z_DISCOUNT_SCALE", intRow, dblScalediscount)
                        dblCumdiscount = dblCumdiscount + dblScalediscount + dblGroupPrice
                        If dblCumdiscount <> 0 Then
                            SetMatrixValues(oMatrix, "15", intRow, dblCumdiscount.ToString)
                        Else
                            GetB1Price(strItemCode, strCardCode, oMatrix, intRow)
                        End If
                        'End If
                    End If
                Next
            End If
            aForm.Freeze(False)
            Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
        Catch ex As Exception
            aForm.Freeze(False)
            Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
#End Region
    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Price "
    Private Sub GetB1Price(ByVal StrItem As String, ByVal strBP As String, ByVal amatrix As SAPbouiCOM.Matrix, ByVal intRow As Integer)
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItems As SAPbobsCOM.Items
        Dim oRec, oREc1, oRecTemp, oRecDiscount As SAPbobsCOM.Recordset
        Dim strSQL, strSQL1, strDiscount, strBPCod As String
        Dim price, discount As Double
        Dim intFlag As Integer
        Dim intPriceList As Integer
        Dim blnDiscountflag As Boolean
        '  Dim oBP As SAPbobsCOM.BusinessPartners
        Dim objForm As SAPbouiCOM.Form
        ' Dim oRec As SAPbobsCOM.Recordset
        Dim oStatic As SAPbouiCOM.StaticText
        Dim oItem, oItem1 As SAPbouiCOM.Item
        price = 0
        discount = 0
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        oItems = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oREc1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecDiscount = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        intFlag = 0
        blnDiscountflag = False
        If oBP.GetByKey(strBP) Then
            'Find discount in Special Price Table
            strSQL = "SELECT T0.[ItemCode], T0.[CardCode], T0.[Discount], T0.[ListNum] FROM OSPP T0 where T0.Cardcode='" & strBP & "' and T0.ItemCode='" & StrItem & "'"
            oRec.DoQuery(strSQL)
            If oRec.RecordCount > 0 Then
                discount = oRec.Fields.Item(2).Value
                intFlag = 1
                blnDiscountflag = True
            End If
            ' Exit Sub

            'Find discount in Discount Group for given BP
            strSQL = "SELECT T0.[CardCode], T0.[ObjType], T0.[ObjKey], T0.[Discount] FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
            oRec.DoQuery(strSQL)
            If blnDiscountflag = False And oRec.RecordCount > 0 Then
                If Convert.ToDouble(oRec.Fields.Item(1).Value) = 52 Then 'Item Group
                    If oItems.GetByKey(StrItem) Then
                        strSQL1 = "SELECT T0.[CardCode], T0.[ObjType], T0.[ObjKey], T0.[Discount] FROM OSPG T0 where T0.Cardcode='" & strBP & "' and T0.ObjKey=" & oItems.ItemsGroupCode
                        oREc1.DoQuery(strSQL1)
                        If oREc1.RecordCount > 0 Then
                            discount = oREc1.Fields.Item(3).Value
                            intFlag = 1
                            blnDiscountflag = True
                        End If
                    End If
                ElseIf Convert.ToDouble(oRec.Fields.Item(1).Value) = 8 Then ' Item Property
                    Dim strProperty, strD As String
                    strBPCod = "Select dscntrel from ocrd where Cardcode='" & strBP & "'"
                    oRecTemp.DoQuery(strBPCod)
                    If oRecTemp.RecordCount > 0 Then
                        strDiscount = oRecTemp.Fields.Item(0).Value
                        Select Case strDiscount
                            Case "L" ' Lowest
                                strD = "Select Min(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                            Case "H" 'Highest
                                strD = "Select Max(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                            Case "A" 'Average
                                strD = "Select Avg(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                            Case "S" 'Discount Total
                                strD = "Select Sum(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                        End Select
                        oRecDiscount.DoQuery(strD)

                        For IntTemp As Integer = 0 To oRec.RecordCount - 1
                            strProperty = oRec.Fields.Item(2).Value
                            strProperty = "QryGroup" & strProperty
                            strSQL1 = "select " & strProperty & " from OITM where Itemcode='" & StrItem & "'"
                            oREc1.DoQuery(strSQL1)
                            If oREc1.RecordCount > 0 Then
                                If oREc1.Fields.Item(0).Value = "Y" Then
                                    discount = oRecDiscount.Fields.Item(0).Value
                                    intFlag = 1
                                    blnDiscountflag = True
                                    Exit For
                                End If
                            End If
                            oRec.MoveNext()
                        Next
                    End If

                ElseIf Convert.ToDouble(oRec.Fields.Item(1).Value) = 43 Then 'Manufacture
                    strSQL1 = "SELECT T0.[CardCode], T0.[ObjType], T0.[ObjKey], T0.[Discount] FROM OSPG T0 where T0.Cardcode='" & strBP & "' T0.ObjKey=" & oItems.Manufacturer
                    oREc1.DoQuery(strSQL1)
                    If oREc1.RecordCount > 0 Then
                        intFlag = 1
                        discount = oREc1.Fields.Item(3).Value
                        blnDiscountflag = True
                    End If
                End If
            End If

            'Find Discount in Hierarchies for given Item code
            strSQL = "SELECT T0.[ItemCode], T0.[CardCode], T0.[ListNum], T0.[Discount], T0.[FromDate], T0.[ToDate]  FROM SPP1 T0 where   T0.Itemcode='" & StrItem & "' and Getdate() between T0. Fromdate and T0.ToDate "
            oRec.DoQuery(strSQL)
            If blnDiscountflag = False And oRec.RecordCount > 0 Then
                discount = oRec.Fields.Item(3).Value
                intPriceList = Convert.ToInt64(oRec.Fields.Item(2).Value)
                strSQL1 = "SELECT T1.[ItemCode], T1.[PriceList], T1.[Price] FROM OPLN T0  INNER JOIN ITM1 T1 ON T0.ListNum = T1.PriceList where T1.Itemcode='" & StrItem & "' and T1.PriceList=" & intPriceList
                oREc1.DoQuery(strSQL)
                If oREc1.RecordCount > 0 Then
                    price = Convert.ToDouble(oREc1.Fields.Item(2).Value)
                    intFlag = 2
                End If
                blnDiscountflag = True
            End If

            If intFlag <> 2 Then 'Take to price for BP Price List
                strSQL = "SELECT T1.[ItemCode], T1.[PriceList], T1.[Price] FROM OPLN T0  INNER JOIN ITM1 T1 ON T0.ListNum = T1.PriceList where T1.Itemcode='" & StrItem & "' and T1.PriceList=" & oBP.PriceListNum
                oRec.DoQuery(strSQL)
                If oRec.RecordCount > 0 Then
                    price = Convert.ToDouble(oRec.Fields.Item(2).Value)
                End If
            End If
        End If
        ' amatrix.Columns.Item("14").Cells.Item(intRow).Specific.value = price
        amatrix.Columns.Item("15").Cells.Item(intRow).Specific.value = discount
        oBP = Nothing
        oItem = Nothing

    End Sub
#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#Region "GetDocumentQuantity"
    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strTempQuantity As String
        strTemp = CompanyDecimalSeprator
        strTempQuantity = strQuantity
        If strQuantity = "" Then
            Return 0
        End If
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("select CurrCode,* from OCRN")
        For introw As Integer = 0 To otest.RecordCount - 1
            strQuantity = strQuantity.Replace(otest.Fields.Item(0).Value.ToString, "")
            otest.MoveNext()
        Next
        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", CompanyDecimalSeprator)
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = Convert.ToDouble(strTempQuantity)
        End Try

        Return dblQuant
    End Function
#End Region

#Region "GetAccount"
    Public Function GetAccount(ByVal aCode As String, ByVal aFieldName As String, ByVal aCondField As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select isnull(" & aFieldName & ",'') from OACT where " & aCondField & "='" & aCode & "'")
        Return oTest.Fields.Item(0).Value
    End Function

    Public Function getExpAccountCode(ByVal aCode As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select isnull(U_Z_ActCode,'') from [@Z_Expances] where U_Z_ExpName='" & aCode & "'")
        Return oTest.Fields.Item(0).Value
    End Function
#End Region

#Region "Get LocalcurrencyAmount"
    Public Sub updatelocalamount(ByVal aCode As String)
        Dim oAmtRS, oExchangeRS As SAPbobsCOM.Recordset
        Dim strCurrency, strsql, strcode, stramount As String
        Dim dblExchangerate, dblAmount As Double
        Dim dblDate As Date
        oAmtRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oAmtRS.DoQuery("SELECT T0.[U_Z_DOCDATE], T1.[Code], T1.[U_Z_CURRENCY], T1.[U_Z_AMOUNT] FROM [dbo].[@Z_OEXP]  T0  inner join  [dbo].[@Z_EXP1]  T1 on T1.[U_Z_REFCODE]=T0.Code where T0.Code='" & aCode & "'")
        oAmtRS.DoQuery("SELECT T0.[U_Z_DOCDATE], T1.[Code], T1.[U_Z_CURRENCY], T1.[U_Z_AMOUNT] FROM [dbo].[@Z_OEXP]  T0  inner join  [dbo].[@Z_EXP1]  T1 on T1.[U_Z_REFCODE]=T0.Code where T1.[U_Z_APPROVED]='P'")
        For intRow As Integer = 0 To oAmtRS.RecordCount - 1
            strcode = oAmtRS.Fields.Item("Code").Value
            dblAmount = oAmtRS.Fields.Item("U_Z_AMOUNT").Value
            strCurrency = oAmtRS.Fields.Item("U_Z_CURRENCY").Value
            dblDate = oAmtRS.Fields.Item("U_Z_DOCDATE").Value
            If strCurrency <> LocalCurrency Then
                strsql = "select isnull(rate,1) from [ortt] where currency='" & strCurrency & "'and convert(varchar(10),ratedate,105)='" & dblDate.ToString("dd-MM-yyyy") & "'"

                oExchangeRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oExchangeRS.DoQuery(strsql)
                dblAmount = dblAmount / oExchangeRS.Fields.Item(0).Value
                dblAmount = Math.Round(dblAmount, 3)
                stramount = LocalCurrency & " " & dblAmount.ToString
                strsql = "Update [@Z_EXP1] set U_Z_LocAmt='" & stramount & "' where code='" & strcode & "'"
                oExchangeRS.DoQuery(strsql)
            Else
                oExchangeRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stramount = LocalCurrency & " " & dblAmount.ToString
                strsql = "Update [@Z_EXP1] set U_Z_LocAmt='" & stramount & "' where code='" & strcode & "'"
                oExchangeRS.DoQuery(strsql)
            End If
            oAmtRS.MoveNext()
        Next
    End Sub
#End Region

#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.string = strvalue
    End Sub
#End Region

    Public Function createPurchaseRequest(ByVal aCode As String) As Boolean
        Dim oRec, oRec1, orec2 As SAPbobsCOM.Recordset
        Dim oDoc As SAPbobsCOM.Documents
        Dim strSQL, strSQL1, strSQL2 As String
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        orec2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "Select * from ""@Z_PRJ2"" where isnull(""U_Z_PR"",'N')='Y' and ""U_Z_BOQREF""='" & aCode & "'"
        oRec.DoQuery(strSQL)
        If oRec.RecordCount > 0 Then
            strSQL = "Select ""U_Z_VENDOR"",Count(*) from ""@Z_PRJ2"" where ""U_Z_VENDOR""<>'' and isnull(""U_Z_DOCENTRY"",0)<=0 and isnull(""U_Z_PR"",'N')='Y' and ""U_Z_BOQREF""='" & aCode & "' group by ""U_Z_VENDOR"""
            oRec.DoQuery(strSQL)
            For intRow As Integer = 0 To oRec.RecordCount - 1
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest)
                oDoc.ReqType = 12
                Dim oUser As SAPbobsCOM.Users
                oUser = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
                If oUser.GetByKey(oApplication.Company.UserSignature) Then
                    Dim intre As String = oUser.UserName
                    '  MsgBox(oUser.UserName)
                    oDoc.Requester = oUser.UserCode
                    oDoc.Comments = "Based  on BoQ Ref : " & aCode

                    '    oDoc.RequesterBranch = oUser.Branch
                    '   oDoc.RequesterDepartment = oUser.Department
                    Dim blnLineExists As Boolean = False
                    Dim aCodeList As String = ""
                    oRec1.DoQuery("Select * from ""@Z_PRJ2"" where isnull(""U_Z_DOCENTRY"",0)<=0 and  ""U_Z_BOQREF""='" & aCode & "' and ""U_Z_VENDOR""='" & oRec.Fields.Item(0).Value & "' and ""U_Z_PR""='Y'")
                    For intLoop As Integer = 0 To oRec1.RecordCount - 1
                        If intLoop > 0 Then
                            oDoc.Lines.Add()
                        End If
                        If aCodeList = "" Then
                            aCodeList = "'" & oRec1.Fields.Item("Code").Value & "'"
                        Else
                            aCodeList = aCodeList & ",'" & oRec1.Fields.Item("Code").Value & "'"
                        End If
                        oDoc.Lines.SetCurrentLine(intLoop)
                        oDoc.Lines.ItemCode = oRec1.Fields.Item("U_Z_ITEMCODE").Value
                        oDoc.Lines.VendorNum = oRec.Fields.Item("U_Z_VENDOR").Value
                        oDoc.Lines.Quantity = oRec1.Fields.Item("U_Z_REQQTY").Value
                        oDoc.Lines.RequiredDate = oRec1.Fields.Item("U_Z_REQDATE").Value
                        oDoc.RequriedDate = oRec1.Fields.Item("U_Z_REQDATE").Value
                        oDoc.Lines.ProjectCode = oRec1.Fields.Item("U_Z_PRJCODE").Value
                        oDoc.Lines.UserFields.Fields.Item("U_Z_MDNAME").Value = oRec1.Fields.Item("U_Z_MODNAME").Value
                        oDoc.Lines.UserFields.Fields.Item("U_Z_ACTNAME").Value = oRec1.Fields.Item("U_Z_ACTNAME").Value
                        oDoc.Lines.UserFields.Fields.Item("U_Z_CNTID").Value = oRec1.Fields.Item("U_Z_CNTID").Value
                        oDoc.Lines.UserFields.Fields.Item("U_Z_CUSTCNTID").Value = oRec1.Fields.Item("U_Z_CUSTCNTID").Value
                        oDoc.Lines.UserFields.Fields.Item("U_Z_BOQREF").Value = aCode
                        oDoc.Lines.LineTotal = oRec1.Fields.Item("U_Z_ESTCOST").Value
                        blnLineExists = True
                        oRec1.MoveNext()
                    Next
                    If blnLineExists = True Then
                        If oDoc.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            Dim strdoc As String
                            oApplication.Company.GetNewObjectCode(strdoc)

                            oRec1.DoQuery("Update PRQ1 set LineVendor='" & oRec.Fields.Item("U_Z_VENDOR").Value & "' where ""DocEntry""=" & CInt(strdoc))
                            oRec1.DoQuery("Select * from OPRQ where ""DocEntry""=" & CInt(strdoc))
                            orec2.DoQuery("Update ""@Z_PRJ2"" set ""U_Z_DOCENTRY""=" & oRec1.Fields.Item("DocEntry").Value & " ,""U_Z_DOCNUM""='" & oRec1.Fields.Item("DocNum").Value & "' where ""Code"" in (" & aCodeList & ")")

                        End If
                    End If
                End If
                oRec.MoveNext()
            Next
        End If
    End Function

#Region "Get Edit Text"
    Public Function getEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String) As String
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        Return objEdit.String
    End Function
    Public Sub setEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String, ByVal newvalue As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        objEdit.String = newvalue
    End Sub
#End Region

#End Region

#Region "Write to LogFile"
    Public Sub WriteToLogFile(ByVal strMsg As String)
        Dim dtdate As Date
        Dim strFileName As String
        Dim FS As FileStream
        Try
            ErrorLogFile = System.Windows.Forms.Application.StartupPath & "\Log.txt"
            strFileName = ErrorLogFile
            If File.Exists(strFileName) Then
                FS = New FileStream(strFileName, FileMode.Append)
            Else
                FS = New FileStream(strFileName, FileMode.Create, FileAccess.ReadWrite)
            End If
            Dim SW As New StreamWriter(FS)
            strMsg = strMsg
            SW.WriteLine(strMsg)
            SW.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function



End Class
