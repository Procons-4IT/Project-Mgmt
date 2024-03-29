Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab.ToUpper()
                oUserTablesMD.TableDescription = strDesc.ToUpper()
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OHPS" Or strTab = "OCRD" Or strTab = "OHEM" Or strTab = "OITB" Or strTab = "JDT1" Or strTab = "INV1" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OINV" Or strTab = "QUT1" Or strTab = "OPRJ") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUserFieldMD.Description = strDesc.ToUpper()
                oUserFieldMD.Name = strCol.ToUpper
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If TableName.StartsWith("@") Or TableName = "OADM" Or TableName = "OPRJ" Or TableName = "OCRD" Or TableName = "OITM" Or TableName = "OITB" Or TableName = "ORDR" Or TableName = "OUSR" Or TableName = "OHEM" Then
            Else
                TableName = "@" & TableName
            End If

            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName.ToUpper()
                objUserFieldMD.Description = ColDescription.ToUpper()
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                If SetValidValue <> "" Then
                    objUserFieldMD.DefaultValue = SetValidValue
                End If
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            objUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE upper(TableID) = '" & Table.ToUpper & "' AND upper(AliasID) = '" & Column.ToUpper & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal strChildTb2 As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES


                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                ' oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                'oUserObjectMD.LogTableName = ""
                ' oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If
                If strChildTb2 <> "" Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.Add()
                    End If
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If

                ' oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub
    Public Function UDOActivity(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_ActName"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_ActName"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    ' oUserObjects.FindColumns.Add()
                    'oUserObjects.FindColumns.SetCurrentLine(1)
                    'oUserObjects.FindColumns.ColumnAlias = strName
                    'oUserObjects.FindColumns.Add()
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function

    Public Function UDOModule(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_ModName"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_ModName"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Status"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Status"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    ' oUserObjects.FindColumns.Add()
                    'oUserObjects.FindColumns.SetCurrentLine(1)
                    'oUserObjects.FindColumns.ColumnAlias = strName
                    'oUserObjects.FindColumns.Add()
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function
    Public Function UDOExpances(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_ExpName"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_ExpName"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Status"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Status"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "U_Z_ActCode"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_ActCode"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function

    Public Function UDOLeaveType(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                'oUserObjects.FormColumns.FormColumnAlias = "U_Z_LeaveType"
                'oUserObjects.FormColumns.FormColumnDescription = "U_Z_LeaveType"
                'oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Status"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Status"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Name"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function

    Public Function UDOLogin(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_UID"
                oUserObjects.FormColumns.FormColumnDescription = "U_UID"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_PWD"
                oUserObjects.FormColumns.FormColumnDescription = "U_PWD"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_EMPID"
                oUserObjects.FormColumns.FormColumnDescription = "U_EMPID"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "U_SUPERUSER"
                oUserObjects.FormColumns.FormColumnDescription = "U_SUPERUSER"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_APPROVER"
                oUserObjects.FormColumns.FormColumnDescription = "U_APPROVER"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_EXPAPPROVER"
                oUserObjects.FormColumns.FormColumnDescription = "U_EXPAPPROVER"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    ' oUserObjects.FindColumns.Add()
                    'oUserObjects.FindColumns.SetCurrentLine(1)
                    'oUserObjects.FindColumns.ColumnAlias = strName
                    'oUserObjects.FindColumns.Add()
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function
#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        ' Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try
            oApplication.Utilities.Message("Project Mgmt Add-on - Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            'AddFields("OHPS", "Daily_rate", "Daily rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)
            'AddFields("OHPS", "Hr_Rate", "Hourly rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)
            AddFields("OHEM", "CardCode", "Business Partner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OHEM", "Daily_rate", "Daily rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)
            AddFields("OHEM", "Hr_Rate", "Hourly rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)


            AddTables("Z_Activity", "Activity - Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddTables("Z_Module", "Business Phase - Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddTables("Z_Expances", "Expences - Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddTables("Z_LeaveType", "Leave Type - Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            AddTables("Z_Login", "Login Details", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("Z_Login", "UID", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , )
            AddFields("Z_Login", "PWD", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha)
            AddFields("Z_Login", "EMPID", "Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha)
            addField("@Z_Login", "SUPERUSER", "Superuser", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_Login", "APPROVER", "Time Sheet Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_Login", "EXPAPPROVER", "Expenses Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_Login", "LEAVAPPROVER", "Expenses Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
        
            'U_ExpApprover


            AddTables("Z_HPRJ", "Project Budeget Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_PRJ1", "Project Budeget-Modules", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("Z_OEXP", "Expences Entry Header", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddTables("Z_EXP1", "Expences Line Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)


            AddTables("Z_OTIM", "Time Sheet Entry Header", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddTables("Z_TIM1", "Time Sheet Line Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)


            AddTables("Z_OLEV", "Leave Request Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OLEV", "Z_EMPCODE", "Employee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_OLEV", "Z_EMPNAME", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_OLEV", "Z_DocDate", "Requeseted Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OLEV", "Z_TYPE", "Leave Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OLEV", "Z_FROMDATE", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OLEV", "Z_TODATE", "To Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OLEV", "Z_DAYS", "Number of days", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OLEV", "Z_REASON", "Reason", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_OLEV", "Z_Approved", "Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,D,P", "Approved,Declined,Pending", "P")
            AddFields("Z_OLEV", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 80)
            AddFields("Z_OLEV", "Z_SubEmp", "Sub level employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)


            AddFields("Z_Activity", "Z_ActName", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_Activity", "Z_ModName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("Z_Activity", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_Module", "Z_ModName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_Module", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_Expances", "Z_ExpName", "Expences Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_Expances", "Z_ActCode", "Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("Z_Expances", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")


            '  AddFields("Z_LeaveType", "Z_LeaveType", "Leave Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 5)
            AddFields("Z_LeaveType", "Z_Name", "Leave Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_LeaveType", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

            AddFields("Z_OEXP", "Z_EMPCODE", "Employee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_OEXP", "Z_EMPNAME", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OEXP", "Z_DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("Z_EXP1", "Z_EXPNAME", "Expences Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_EXP1", "Z_EXPTYPE", "Expences type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,N", "Project,Normal", "N")
            AddFields("Z_EXP1", "Z_PRJCODE", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_EXP1", "Z_PRJNAME", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_EXP1", "Z_CURRENCY", "Transaction Currencty", SAPbobsCOM.BoFieldTypes.db_Alpha, , 5)
            AddFields("Z_EXP1", "Z_AMOUNT", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_EXP1", "Z_DATE", "Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_EXP1", "Z_RefCode", "Document reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            addField("@Z_EXP1", "Z_Approved", "Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,D,P", "Approved,Declined,Pending", "P")
            AddFields("Z_EXP1", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 80)
            AddFields("Z_EXP1", "Z_LocAmt", "Amount in Local Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_EXP1", "Z_SysAmt", "Amount in System Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_EXP1", "Z_Attachment", "Attachment File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_EXP1", "Z_Ref1", "Ref1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 230)
            AddFields("Z_EXP1", "Z_Flag", "Ready for Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, 1)
            AddFields("Z_EXP1", "Z_DocNum", "Journal Entry Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_EXP1", "Z_LocAmount", "Converted Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_EXP1", "Z_ModName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_EXP1", "Z_ActName", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_EXP1", "Z_EMPID", "Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_EXP1", "Z_EMPNAME", "Employee NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_EXP1", "Z_Attachment1", "Attachment File Name", SAPbobsCOM.BoFieldTypes.db_Memo)



            AddFields("Z_OTIM", "Z_EMPCODE", "Employee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_OTIM", "Z_EMPNAME", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OTIM", "Z_DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("Z_TIM1", "Z_PRJCODE", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_TIM1", "Z_PRJNAME", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_TIM1", "Z_PRCNAME", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_TIM1", "Z_ACTNAME", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_TIM1", "Z_DATE", "Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_TIM1", "Z_HOURS", "HOURS", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_TIM1", "Z_RefCode", "Document reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            addField("@Z_TIM1", "Z_Approved", "Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,D,P", "Approved,Declined,Pending", "P")
            AddFields("Z_TIM1", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 80)
            AddFields("Z_TIM1", "Z_BdgQty", "Budgeted Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_TIM1", "Z_Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_TIM1", "Z_Measure", "Measure", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            addField("@Z_TIM1", "Z_Type", "Activity Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "I,R", "Item,Resource", "R")


           
            AddFields("Z_HPRJ", "Z_PRJCODE", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_HPRJ", "Z_PRJNAME", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_HPRJ", "Z_BUDGET", "Project Budgets in hours", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_HPRJ", "Z_FromDate", "Start Date of Project", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_HPRJ", "Z_ToDate", "End Date of Project", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_HPRJ", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "E,X,I,H,C", "Estimation,Execution,In Process,On Hold,Completed", "E")
            AddFields("Z_HPRJ", "Z_TotalExpense", "Total Estimated Expenses", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_HPRJ", "Z_Approval", "Timesheet approval requires", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")



            AddFields("Z_PRJ1", "Z_ModName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRJ1", "Z_ActName", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRJ1", "Z_Days", "Estimated Days", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRJ1", "Z_Hours", "Estimated Hours", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRJ1", "Z_EmpID", "Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PRJ1", "Z_Position", "Position", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_PRJ1", "Z_Order", "Sales Order", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_PRJ1", "Z_OrdEntry", "Sales Order Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PRJ1", "Z_OrdNum", "Sales Order Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PRJ1", "Z_Amount", "Estimated Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_PRJ1", "Z_Status", "Activity Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "I,P,C", "In Process,Pending,Completed", "P")
            'addField("@Z_PRJ1", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "E,X,I,H,C", "Estimation,Execution,In Process,On Hold,Completed", "I")
            AddFields("Z_PRJ1", "Z_FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PRJ1", "Z_ToDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PRJ1", "Z_Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PRJ1", "Z_Measure", "Measure", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_PRJ1", "Z_CmpDate", "Completion Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PRJ1", "Z_BOQ", "BOQ RefNumber", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            addField("@Z_PRJ1", "Z_Type", "Activity Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "I,R,E", "Item,Resource,Expenses", "R")
            AddFields("Z_PRJ1", "Z_ExpType", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)


            AddTables("Z_PRJ2", "Bill of Quantity Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_PRJ2", "Z_PRJCODE", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PRJ2", "Z_PRJNAME", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRJ2", "Z_ModName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRJ2", "Z_ActName", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRJ2", "Z_BOQRef", "BOQ Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_PRJ2", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRJ2", "Z_ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRJ2", "Z_ItemName", "Item Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRJ2", "Z_UOM", "Item UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PRJ2", "Z_ReqQty", "Required Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PRJ2", "Z_EstCost", "Estimated Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRJ2", "Z_ReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PRJ2", "Z_Vendor", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRJ2", "Z_VendorName", "Vendorname", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRJ2", "Z_Comments", "Comments", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_PRJ2", "Z_UnitPrice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)


            AddFields("INV1", "Z_ActName", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("INV1", "Z_MDName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("INV1", "Z_BOQRef", "BoQ Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddFields("JDT1", "Z_ActName", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("JDT1", "Z_MDName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("JDT1", "Z_BOQRef", "BoQ Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            'AddFields("Z_TimeSheet", "Z_PRJCODE", "PROJECT CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("Z_TimeSheet", "Z_PRJNAME", "PROJECT NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("Z_TimeSheet", "Z_EMPID", "EMPLOYEE ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            'AddFields("Z_TimeSheet", "Z_EMPNAME", "EMPLOYEE NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_TimeSheet", "Z_MODNAME", "BUSINESS PROCESS NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            'AddFields("Z_TimeSheet", "Z_ACTNAME", "ACTIVITY NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            'AddFields("Z_TimeSheet", "Z_DATE", "ACTIVITY DATE", SAPbobsCOM.BoFieldTypes.db_Date)
            'AddFields("Z_TimeSheet", "Z_HOURS", "NUMBER OF HOURS", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)




            AddTables("Z_DHPRJ", "Budeget Template Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_DPRJ1", "Budeget Template Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("Z_DHPRJ", "Z_PRJCODE", "Template Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_DHPRJ", "Z_PRJNAME", "Template Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_DHPRJ", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "E,X,I,H,C", "Estimation,Execution,In Process,On Hold,Completed", "E")
            addField("@Z_DHPRJ", "Z_ACTIVE", "ACTIVE", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")


            AddFields("Z_DPRJ1", "Z_ModName", "Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DPRJ1", "Z_ActName", "Activity Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DPRJ1", "Z_Days", "Estimated Days", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_DPRJ1", "Z_Hours", "Estimated Hours", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_DPRJ1", "Z_EmpID", "Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_DPRJ1", "Z_Position", "Position", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("@Z_DPRJ1", "Z_Order", "Sales Order", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_DPRJ1", "Z_OrdEntry", "Sales Order Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_DPRJ1", "Z_OrdNum", "Sales Order Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_DPRJ1", "Z_Amount", "Estimated Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_DPRJ1", "Z_Status", "Activity Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "I,P,C", "In Process,Pending,Completed", "P")
            'addField("@Z_PRJ1", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "E,X,I,H,C", "Estimation,Execution,In Process,On Hold,Completed", "I")
            AddFields("Z_DPRJ1", "Z_FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_DPRJ1", "Z_ToDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_DPRJ1", "Z_Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_DPRJ1", "Z_Measure", "Measure", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_DPRJ1", "Z_CmpDate", "Completion Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_DPRJ1", "Z_BOQ", "BOQ RefNumber", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            addField("@Z_DPRJ1", "Z_Type", "Activity Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "I,R,E", "Item,Resource,Expenses", "R")
            AddFields("Z_DPRJ1", "Z_ExpType", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)


            'Phase II Fields Creations
            AddFields("OPRJ", "Z_CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OPRJ", "Z_CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddFields("Z_HPRJ", "Z_CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_HPRJ", "Z_CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("OADM", "Z_EmpRate", "Project Mgmt Employee Rate", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,R", "Payroll,CostRate", "R")


            'Contract Agreement
            AddTables("Z_OPAT", "Contract Agreement", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OPAT", "Z_CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_OPAT", "Z_CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_OPAT", "Z_Type", "Contract Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "C,S", "Customer,Supplier", "C")

            AddFields("Z_OPAT", "Z_CntctCode", "Contact person", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPAT", "Z_Telephone", "Telephone", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPAT", "Z_EMail", "Email", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OPAT", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPAT", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPAT", "Z_TermDate", "Termination Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPAT", "Z_SignDate", "Signing date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPAT", "Z_Amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OPAT", "Z_Retention", "Retention %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_OPAT", "Z_DPAmount", "Down Payment Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OPAT", "Z_Link", "Link All documents", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            addField("@Z_OPAT", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,O,D,T", "Approved,OnHold,Draft,Terminated", "O")
            AddFields("Z_OPAT", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)

            AddFields("Z_OPAT", "Z_Rate", "Daily Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)



            AddTables("Z_PAT1", "Contract Attachment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PAT1", "Z_Path", "File Path", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PAT1", "Z_FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PAT1", "Z_Date", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("Z_PRJ2", "Z_DocEntry", "Purchase Request Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PRJ2", "Z_DocNum", "Purchae Request DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("Z_PRJ2", "Z_PR", "Create Purchase Request", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_PRJ2", "Z_CntID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRJ2", "Z_Position", "Position", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRJ2", "Z_CustCntID", "Customer Contract ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_PRJ1", "Z_CntID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRJ1", "Z_CustCntID", "Customer Contract ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_HPRJ", "Z_CustCntID", "Customer Contract ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("OHEM", "Z_IsProject", "Project Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")


            AddFields("INV1", "Z_CntID", "Sub.Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("INV1", "Z_CntName", "Sub Contractor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("INV1", "Z_CustCntID", "Customer Agreement Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)

            AddFields("OADM", "Z_EmpRate", "Project Mgmt Employee Costs", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            addField("OHEM", "Z_IsProject", "Project Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("OHEM", "Z_IsProject", "Project Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            AddTables("Z_PRJ3", "Project Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PRJ3", "Z_FilePath", "File Path", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PRJ3", "Z_FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            addField("@Z_TIM1", "Z_EmpApproval", "Employee Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,C", "Pending,Confirmed", "P")

            AddFields("OPRJ", "Z_CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OPRJ", "Z_CARDNAME", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)


            AddFields("OPRJ", "Z_EmpID", "Delivery Manager", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OPRJ", "Z_EMPNAME", "Delivery Manager Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("OPRJ", "Z_Internal", "Internal Project", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_HPRJ", "Z_EmpID", "Delivery Manager", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_HPRJ", "Z_EMPNAME", "Delivery Manager Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("Z_HPRJ", "Z_Internal", "Internal Project", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

            AddFields("Z_HPRJ", "Z_TotHours", "Total Hours", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_HPRJ", "Z_TotCost", "Total Estimated Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_HPRJ", "Z_SlpCode", "Sales Person Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_HPRJ", "Z_SlpName", "Sales Person Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)


            AddFields("Z_EXP1", "Z_CNTID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_EXP1", "Z_DisRule", "Distribution Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)



            AddTables("Z_PAT2", "Contract Employee Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PAT2", "Z_EmpCode", "Employee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PAT2", "Z_EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PAT2", "Z_Designation", "Designation", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)



            AddFields("Z_OTIM", "Z_CntID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("@Z_OTIM", "Z_Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "C,E", "Contractor,Employee", "E")


             '---- User Defined Object's
            oApplication.Utilities.Message("Project Mgmt  Add-on - Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            CreateUDO()


            'If oApplication.Company.InTransaction() Then
            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'End If

        Catch ex As Exception
            'If oApplication.Company.InTransaction() Then
            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Public Sub CreateUDO()
        Try
            ' AddUDO("Z_Activity", "Activity Master", "Z_Activity", , , , SAPbobsCOM.BoUDOObjType.boud_MasterData)
            AddUDO("Z_PRJ", "Project Budget-Details", "Z_HPRJ", "U_Z_PRJCODE", "U_Z_PRJNAME", "Z_PRJ1", "Z_PRJ3", SAPbobsCOM.BoUDOObjType.boud_Document)
            UDOActivity("Z_Activity", "Activity-master", "Z_Activity", 1, "U_Z_ActName", )
            UDOModule("Z_Module", "Phase - Master", "Z_Module", 1, "U_Z_ModName")
            UDOExpances("Z_Expances", "Expences - Master", "Z_Expances", 1, "U_Z_ExpName")
            UDOLogin("Z_LOGIN", "Login Setup", "Z_LOGIN", 1, "U_UID")
            UDOLeaveType("Z_LeaveType", "Leave Type  - Master", "Z_LeaveType", 1, "U_Z_NAME")

            AddUDO("Z_DPRJ", "Project Budget-Template", "Z_DHPRJ", "U_Z_PRJCODE", "U_Z_PRJNAME", "Z_DPRJ1", , SAPbobsCOM.BoUDOObjType.boud_Document)

            'Phase II
            AddUDO("Z_OPAT", "Project Contract Agreement", "Z_OPAT", "DocEntry", "U_Z_CardCode", "Z_PAT1", "Z_PAT2", SAPbobsCOM.BoUDOObjType.boud_Document)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
