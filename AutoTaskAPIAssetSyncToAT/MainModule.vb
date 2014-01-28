Imports AutoTaskAPIAssetSyncToAT.autotaskwebservices
Imports AutoTaskAPIAssetSyncToAT.My.MySettings
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports AutoTaskAPIAssetSyncToAT


Module MainModule
    Dim myService As New ATWS '// this call can be issued to any zone, we have chosen '// webservices.autotask.net (North American zone)
    Dim ITSPSAAutoTaskAPIStr As String
    Dim ITSAssetDBStr As String
    Dim ITSupport247dbStr As String
    Dim ITSPatchDBStr As String
    Dim ITSMACAssetDBStr As String
    Dim ITSVMWareAssetDBStr As String
    Dim ITSMDMgmtDBStr As String
    Dim ITSRMMVaultDBStr As String
    Dim ITSLinuxAssetDBStr As String
    Dim ITSPassVaultDBStr As String
    Dim ITSSADDBStr As String

    Dim strConTO As String
    Dim strCmdTO As String

    Dim strAssetFullSyncTimeFrom As Date
    Dim strAssetFullSyncTimeTo As Date

    Dim ConnITSPSAAutoTaskAPI As New SqlConnection
    Dim ConnITSAssetDB As New SqlConnection
    Dim ConnItsupport247DB As New SqlConnection
    Dim ConnITSPatchDB As New SqlConnection
    Dim ConnITSMACAssetDB As New SqlConnection
    Dim ConnITSVMWareAssetDB As New SqlConnection
    Dim ConnITSMDMgmtDB As New SqlConnection
    Dim ConnITSRMMVaultDB As New SqlConnection
    Dim ConnITSLinuxAssetDB As New SqlConnection
    Dim ConnITSPassVaultDB As New SqlConnection
    Dim ConnITSSADDB As New SqlConnection

    Dim TotalAResCntToSync As Long
    Dim TotalAResCntSync As Long

    Dim TotalMResCntToSync As Long
    Dim TotalMResCntSync As Long

    Dim TotalVResCntToSync As Long
    Dim TotalVResCntSync As Long


    Dim DTLastSyncData1 As System.Nullable(Of DateTime)
    Dim DTLastSyncData2 As System.Nullable(Of DateTime)
    Dim DTLastSyncData3 As System.Nullable(Of DateTime)

    Dim DTFullSyncData1 As System.Nullable(Of DateTime)
    Dim DTFullSyncData2 As System.Nullable(Of DateTime)
    Dim DTFullSyncData3 As System.Nullable(Of DateTime)


    Dim FileIO As New FileIO
    Dim ClsDBO As New ClsDBO
    Dim URLZone As String
    Dim SyncFlag As String
    Dim MemberidFrom As Long
    Dim MemberidTo As Long
    Dim GlobalMemberId As Long
    Dim GlobalRegID As Long
    Dim IsMigrated As Boolean
    Dim OutOneTimeMapped As Boolean

    ''----- For Entity Product----------------------
    Dim ATProductAllocationCodeID As String
    Dim ATBackupID As Long
    Dim ATDesktopID As Long
    Dim ATMobileID As Long
    Dim ATServerID As Long
    Dim AtVirtualHostID As Long
    '' ---------------------------------------------

    Dim MGetUDFINFO As Long
    Dim PInsertBatch As Integer
    Dim UpdateBatch As Long
    Dim InsertBatch As Long
    Dim Ucount200 As Long
    Dim Icount200 As Long
    Const MaxInsUpdBatch = 5
    Dim NumberOfExternalRequestBuffer As Long
    
    Public ProdCatTable As DataTable
    Public ProdTable As DataTable
    Public ProdConfigTable As DataTable
    Public AssetTable As DataTable
    
    Public ds As DataSet
    Dim ObjRsAssetData As SqlDataReader
    Dim ObjRsMobileDevice As SqlDataReader
    Dim ObjRsBackUpData As SqlDataReader
    Sub Main()
        Try
            Dim AppPath As String
            Dim blnSUCCESS As Boolean
            Dim OutErr As String
            OutErr = ""
            If PrevInstance() Then Exit Sub
            AppPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            blnSUCCESS = FileIO.CreateDir(AppPath + "\\Log", OutErr)
            If blnSUCCESS = True Then
                FileIO.LogNotify("MAIN", FileIO.NotifyType.INFO, "START")
                If OpenConnection() = True Then
                    ATAPIMemberWiseAssetSync()
                    CloseConnection()
                End If
                FileIO.LogNotify("MAIN", FileIO.NotifyType.INFO, "END")
            End If
        Catch ex As Exception

        End Try
        End
    End Sub
    Function ATAPIMemberWiseAssetSync() As Boolean
        Try

            GlobalMemberId = 0
            GlobalRegID = 0

            If Split(UCase(FileIO.ExeFileName), "_")(0) <> "PRAUTOTASKAPIASSETSYNCTOAT" Then
                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.ERR, FileIO.ExeFileName & " Not Match with PRAUTOTASKAPIASSETSYNCTOAT")
                Return False
            End If

            MemberidFrom = CDbl(Split(FileIO.ExeFileName, "_")(1))
            MemberidTo = CDbl(Split(FileIO.ExeFileName, "_")(2))

            If MemberidFrom = 0 Or MemberidTo = 0 Then
                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.ERR, "MemberIdFrom Range Not Valid ", MemberidFrom, MemberidTo)
                Return False
            End If

            Dim Mtime As Date
            Dim MSyncFlag As String
            Dim SplitCommandArg()

            Mtime = Format(Now, "hh:mm tt")

            Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
            Dim arguments As String() = Environment.GetCommandLineArgs()

            For i As Integer = 0 To CommandLineArgs.Count - 1
                'the message box is just a example. you can make Ifs to check the arguments and do the desired options
                SplitCommandArg = Split(CommandLineArgs(i), ",")
                If UBound(SplitCommandArg) > 0 Then
                    strAssetFullSyncTimeFrom = SplitCommandArg(0)
                    strAssetFullSyncTimeTo = SplitCommandArg(1)
                End If
            Next

            FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, Mtime, strAssetFullSyncTimeFrom, strAssetFullSyncTimeTo)

            If Mtime > strAssetFullSyncTimeFrom And Mtime < strAssetFullSyncTimeTo Then
                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, "FULLBack SYNC")
                SyncFlag = 2
                MSyncFlag = "FULLSYNC"
            Else
                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, "NEW SYNC")
                SyncFlag = 1
                MSyncFlag = "NEWSYNC"
            End If


            DTLastSyncData1 = Nothing
            DTLastSyncData2 = Nothing
            DTLastSyncData3 = Nothing

            DTFullSyncData1 = Nothing
            DTFullSyncData2 = Nothing
            DTFullSyncData3 = Nothing


            FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, MSyncFlag & " Process Start For MemberId Range : ", MemberidFrom, MemberidTo)

            If MemberidFrom = 0 Or MemberidTo = 0 Or Trim(MSyncFlag) = "" Then
                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.ERR, "Member Range Not Valid / SyncFlag Is Missing")
                Return False
            End If

            Dim ObjCmdMember As New SqlCommand
            Dim ObjRSMemberList As SqlDataReader

            With ObjCmdMember
                ObjCmdMember.Connection = ConnITSPSAAutoTaskAPI
                ObjCmdMember.CommandTimeout = strCmdTO
                ObjCmdMember.CommandType = CommandType.StoredProcedure
                ObjCmdMember.CommandText = "USP_AT_Get_MemberDetails_PR"
                .Parameters.Add("@InType", System.Data.SqlDbType.Int, 4)
                .Parameters("@InType").Value = 2
                .Parameters.Add("@InMemberFrom", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberFrom").Value = MemberidFrom
                .Parameters.Add("@InMemberTo", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberTo").Value = MemberidTo
                ObjRSMemberList = ObjCmdMember.ExecuteReader
                ObjCmdMember.Dispose()
            End With

            If ObjRSMemberList.HasRows Then

                If TableCreationProcess() = True Then
                    While (ObjRSMemberList.Read())

                        GlobalMemberId = ObjRSMemberList.Item("MemberId")

                        ProdCatTable.Clear()
                        ProdConfigTable.Clear()
                        ProdTable.Clear()
                        AssetTable.Clear()

                        GlobalRegID = 0
                        ATProductAllocationCodeID = ""
                        IsMigrated = False

                        PInsertBatch = 0
                        UpdateBatch = 0
                        InsertBatch = 0
                        Ucount200 = 0
                        Icount200 = 0

                        DTLastSyncData1 = Nothing
                        DTLastSyncData2 = Nothing
                        DTLastSyncData3 = Nothing
                        DTFullSyncData1 = Nothing
                        DTFullSyncData2 = Nothing
                        DTFullSyncData3 = Nothing

                        FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, "Process Start For MemberId : ", ObjRSMemberList.Item("MemberId"))
                        Try
                            If GetAssetDataFromDB() = True Then
                                URLZone = GetZoneInfoAutoTaskAPI(ObjRSMemberList.Item("username"), ObjRSMemberList.Item("Password"))

                                If URLZone <> "" Then
                                    If SetZoneWiseAutoTaskAPI(ObjRSMemberList.Item("username"), ObjRSMemberList.Item("Password"), ObjRSMemberList.Item("MaterialCostCode") & "") = True Then
                                        If getThresholdAndUsageInfo() = True Then

                                            IsMigrated = ObjRSMemberList.Item("IsMigrated")

                                            If ATProductAllocationCodeID <> "" And FetchPicklistValueIDForProdCategory() = True And FetchPicklistValueIDForProdConfig() = True And FetchProductDetails() = True And FetchAssetDetailsFromAutotask() = True Then

                                                TotalAResCntToSync = 0
                                                TotalAResCntSync = 0

                                                TotalMResCntToSync = 0
                                                TotalMResCntSync = 0

                                                TotalVResCntToSync = 0
                                                TotalVResCntSync = 0

                                                AssetSyncToAutoTask(ObjRSMemberList.Item("MemberId"))
                                                CreateAndUpdateMobileProcess()
                                                CreateAndUpdateVaultDataProcess()

                                                If CreateNewProductInBatchOf200() = True Then
                                                    If ProcessOFMakingBatch() = True Then
                                                        CreateAssetInBatchOf200()
                                                        UpdateAssetInBatchOf200()

                                                        If TotalAResCntToSync > 0 And TotalAResCntSync > 0 Then
                                                            ATDeviceSyncStatusInsertUpdate(ObjRSMemberList.Item("MemberId"), 1, DTLastSyncData1, DTFullSyncData1, TotalAResCntToSync, TotalAResCntSync)
                                                        End If

                                                        If TotalMResCntToSync > 0 And TotalMResCntSync > 0 Then
                                                            ATDeviceSyncStatusInsertUpdate(ObjRSMemberList.Item("MemberId"), 2, DTLastSyncData2, DTFullSyncData2, TotalMResCntToSync, TotalMResCntSync)
                                                        End If

                                                        If TotalVResCntToSync > 0 And TotalVResCntSync > 0 Then
                                                            ATDeviceSyncStatusInsertUpdate(ObjRSMemberList.Item("MemberId"), 3, DTLastSyncData3, DTFullSyncData3, TotalVResCntToSync, TotalVResCntSync)
                                                        End If
                                                    Else
                                                        ModLogErrorInsert(GlobalMemberId, GlobalRegID, "ATAPIMemberWiseAssetSync ProcessOFMakingBatch Failed")
                                                    End If
                                                Else
                                                    ModLogErrorInsert(GlobalMemberId, GlobalRegID, "ATAPIMemberWiseAssetSync New product Creation Failed")
                                                End If

                                            Else
                                                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, "AllocationID / ProductCategory ID Not Found")
                                            End If
                                        Else
                                            FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.ERR, "getThresholdAndUsageInfo Error Cannot proceed")
                                        End If
                                    Else
                                        FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.ERR, "SetZoneWiseAutoTaskAPI Failed To Set URLZONE For MyService")
                                    End If
                                Else
                                    FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.ERR, "GetZoneInfoAutoTaskAPI Failed To Retrieve URLZONE")
                                End If
                            Else
                                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, "GetAssetDataFromDB  Not Data Found / For First Time To Process ")
                            End If
                        Catch ex As Exception
                            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
                            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
                            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "ATAPIMemberWiseAssetSync 1 :" & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
                        End Try
                    End While
                Else
                    ModLogErrorInsert(GlobalMemberId, 0, "ATAPIMemberWiseAssetSync Table Creation Failed")
                End If
            Else
                FileIO.LogNotify("ATAPIMemberWiseAssetSync", FileIO.NotifyType.INFO, "USP_AT_Get_MemberDetails_PR No Record Found To Process..")
                Return True
            End If
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "ATAPIMemberWiseAssetSync 2 :" & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function AssetSyncToAutoTask(ByVal InMemberId As String)
        Try
            Dim ObjRsPSAVMWare As SqlDataReader
            Dim ObjRsPSALinux As SqlDataReader

            Dim ObjRsPSAVMAC As SqlDataReader
            Dim ObjRsDeviceLogin As SqlDataReader

            Dim ObjRsDeviceCPU As SqlDataReader
            Dim ObjRSDeviceHD As SqlDataReader
            Dim ObjRsDevicePatch As SqlDataReader
            Dim ObjRsProdKey As SqlDataReader
            Dim ObjRsWarrantyExp As SqlDataReader

            Dim strtOTALpHYSICALMEMORY As String
            Dim strSYSTEMMANUFACTURER As String
            Dim strSYSTEMMODEL As String
            Dim strBIOSSerialNO As String
            Dim StrNUMBEROFPROCESSORS As String
            Dim strCurrentClockSpeed As String
            Dim strHardDiskInfo As String
            Dim strLOGGEDINUSER As String
            Dim strPatchesMissing As String
            Dim strGateWayIP As String
            Dim strInproductKey As String
            Dim I As Integer

            If ObjRsAssetData.HasRows Then
                While (ObjRsAssetData.Read())
                    FileIO.LogNotify("AssetSyncToAutoTask", FileIO.NotifyType.INFO, "Process Start For RegId,ResType : ", ObjRsAssetData.Item("RegId"), ObjRsAssetData.Item("Restype"))
                    Try

                        TotalAResCntToSync = TotalAResCntToSync + 1

                        If GetFailedRegID(InMemberId, ObjRsAssetData.Item("RegId")) = 0 Then


                            strtOTALpHYSICALMEMORY = ""
                            strSYSTEMMANUFACTURER = ""
                            strSYSTEMMODEL = ""
                            strBIOSSerialNO = ""
                            StrNUMBEROFPROCESSORS = ""
                            strCurrentClockSpeed = ""
                            strHardDiskInfo = ""
                            strLOGGEDINUSER = ""
                            strPatchesMissing = ""
                            strGateWayIP = ""
                            strInproductKey = ""

                            Dim StrINWarrantyExpDt As System.Nullable(Of DateTime) = Nothing

                            If ObjRsAssetData.Item("RESTYPE") = 5 Then    ' VMWare
                                Dim ObjCmdPSAVMWare As New SqlCommand
                                With ObjCmdPSAVMWare
                                    .Connection = ConnITSVMWareAssetDB
                                    .CommandType = CommandType.StoredProcedure
                                    .CommandText = "USP_PSA_Integration_VMWare"
                                    .CommandTimeout = strCmdTO
                                    .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                    ObjRsPSAVMWare = ObjCmdPSAVMWare.ExecuteReader
                                    ObjCmdPSAVMWare.Dispose()
                                End With
                                ObjCmdPSAVMWare = Nothing

                                If ObjRsPSAVMWare.HasRows Then
                                    I = 0
                                    Do
                                        I = I + 1
                                        While ObjRsPSAVMWare.Read()
                                            If I = 1 Then
                                                strtOTALpHYSICALMEMORY = ObjRsPSAVMWare![tOTALpHYSICALMEMORY] & ""
                                                strSYSTEMMANUFACTURER = ObjRsPSAVMWare![SYSTEMMANUFACTURER] & ""
                                                strSYSTEMMODEL = ObjRsPSAVMWare![SYSTEMMODEL] & ""
                                                strBIOSSerialNO = ObjRsPSAVMWare![BASEBOARDSRNO] & ""
                                                StrNUMBEROFPROCESSORS = Val(ObjRsPSAVMWare![NUMBEROFPROCESSORS] & "")
                                            End If
                                            If I = 2 Then
                                                strCurrentClockSpeed = strCurrentClockSpeed & IIf(IsDBNull(ObjRsPSAVMWare.Item("CurrentClockSpeed")), "", ObjRsPSAVMWare.Item("CurrentClockSpeed") & "MHz" & "@*@")
                                            End If
                                            If I = 3 Then
                                                strHardDiskInfo = strHardDiskInfo & IIf(IsDBNull(ObjRsPSAVMWare.Item("Caption")), "", ObjRsPSAVMWare.Item("Caption") & ":") _
                                                                                  & IIf(IsDBNull(ObjRsPSAVMWare.Item("FreespaceMB")), "", ObjRsPSAVMWare.Item("FreespaceMB") & "(MB):") _
                                                                                  & IIf(IsDBNull(ObjRsPSAVMWare.Item("SizeMB")), "", ObjRsPSAVMWare.Item("SizeMB") & "(MB)" & "@*@")
                                            End If
                                        End While
                                    Loop While ObjRsPSAVMWare.NextResult()

                                    If Trim$(strCurrentClockSpeed) <> "" Then
                                        strCurrentClockSpeed = Mid$(strCurrentClockSpeed, 1, Len(strCurrentClockSpeed) - 3)
                                        strCurrentClockSpeed = "CPU Detail: " & strCurrentClockSpeed
                                    End If

                                    If Trim$(strHardDiskInfo) <> "" Then
                                        strHardDiskInfo = Mid$(strHardDiskInfo, 1, Len(strHardDiskInfo) - 3)
                                        strHardDiskInfo = "Harddisk Info: " & strHardDiskInfo
                                    End If
                                End If
                                ObjRsPSAVMWare.Close()

                            ElseIf ObjRsAssetData.Item("RESTYPE") = 6 Then   'MAC
                                Dim ObjCmdPSAMAC As New SqlCommand
                                With ObjCmdPSAMAC
                                    .Connection = ConnITSMACAssetDB
                                    .CommandType = CommandType.StoredProcedure
                                    .CommandText = "USP_PSA_Integration_MAC"
                                    .CommandTimeout = strCmdTO
                                    .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                    ObjRsPSAVMAC = ObjCmdPSAMAC.ExecuteReader
                                    ObjCmdPSAMAC.Dispose()
                                End With

                                ObjCmdPSAMAC = Nothing

                                If ObjRsPSAVMAC.HasRows Then
                                    I = 0
                                    Do
                                        I = I + 1
                                        While ObjRsPSAVMAC.Read()
                                            If I = 1 Then
                                                strtOTALpHYSICALMEMORY = ObjRsPSAVMAC![tOTALpHYSICALMEMORY] & ""
                                                strSYSTEMMANUFACTURER = ObjRsPSAVMAC![SYSTEMMANUFACTURER] & ""
                                                strSYSTEMMODEL = ObjRsPSAVMAC![SYSTEMMODEL] & ""
                                                strBIOSSerialNO = ObjRsPSAVMAC![BASEBOARDSRNO] & ""
                                                StrNUMBEROFPROCESSORS = Val(ObjRsPSAVMAC![NUMBEROFPROCESSORS] & "")
                                            End If
                                            If I = 2 Then
                                                strCurrentClockSpeed = strCurrentClockSpeed & IIf(IsDBNull(ObjRsPSAVMAC.Item("CurrentClockSpeed")), "", ObjRsPSAVMAC.Item("CurrentClockSpeed") & "MHz" & "@*@")
                                            End If
                                            If I = 3 Then
                                                strHardDiskInfo = strHardDiskInfo & IIf(IsDBNull(ObjRsPSAVMAC.Item("Caption")), "", ObjRsPSAVMAC.Item("Caption") & ":") _
                                                                                  & IIf(IsDBNull(ObjRsPSAVMAC.Item("FreespaceMB")), "", ObjRsPSAVMAC.Item("FreespaceMB") & "(MB):") _
                                                                                  & IIf(IsDBNull(ObjRsPSAVMAC.Item("SizeMB")), "", ObjRsPSAVMAC.Item("SizeMB") & "(MB)" & "@*@")
                                            End If
                                        End While
                                    Loop While ObjRsPSAVMAC.NextResult()

                                    If Trim$(strCurrentClockSpeed) <> "" Then
                                        strCurrentClockSpeed = Mid$(strCurrentClockSpeed, 1, Len(strCurrentClockSpeed) - 3)
                                        strCurrentClockSpeed = "CPU Detail: " & strCurrentClockSpeed
                                    End If

                                    If Trim$(strHardDiskInfo) <> "" Then
                                        strHardDiskInfo = Mid$(strHardDiskInfo, 1, Len(strHardDiskInfo) - 3)
                                        strHardDiskInfo = "Harddisk Info: " & strHardDiskInfo
                                    End If
                                End If

                                ObjRsPSAVMAC.Close()

                            ElseIf ObjRsAssetData.Item("RESTYPE") = 7 Then   'Linux
                                Dim ObjCmdPSALinux As New SqlCommand
                                With ObjCmdPSALinux
                                    .Connection = ConnITSLinuxAssetDB
                                    .CommandType = CommandType.StoredProcedure
                                    .CommandText = "USP_PSA_Integration_Linux"
                                    .CommandTimeout = strCmdTO
                                    .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                    ObjRsPSALinux = ObjCmdPSALinux.ExecuteReader
                                    ObjCmdPSALinux.Dispose()
                                End With

                                ObjCmdPSALinux = Nothing

                                If ObjRsPSALinux.HasRows Then
                                    I = 0
                                    Do
                                        I = I + 1
                                        While ObjRsPSALinux.Read()
                                            If I = 1 Then
                                                strtOTALpHYSICALMEMORY = ObjRsPSALinux![tOTALpHYSICALMEMORY] & ""
                                                strSYSTEMMANUFACTURER = ObjRsPSALinux![SYSTEMMANUFACTURER] & ""
                                                strSYSTEMMODEL = ObjRsPSALinux![SYSTEMMODEL] & ""
                                                strBIOSSerialNO = ObjRsPSALinux![BASEBOARDSRNO] & ""
                                                StrNUMBEROFPROCESSORS = Val(ObjRsPSALinux![NUMBEROFPROCESSORS] & "")
                                            End If
                                            If I = 2 Then
                                                strCurrentClockSpeed = strCurrentClockSpeed & IIf(IsDBNull(ObjRsPSALinux.Item("CurrentClockSpeed")), "", ObjRsPSALinux.Item("CurrentClockSpeed") & "MHz" & "@*@")
                                            End If
                                            If I = 3 Then
                                                strHardDiskInfo = strHardDiskInfo & IIf(IsDBNull(ObjRsPSALinux.Item("Caption")), "", ObjRsPSALinux.Item("Caption") & ":") _
                                                                                  & IIf(IsDBNull(ObjRsPSALinux.Item("FreespaceMB")), "", ObjRsPSALinux.Item("FreespaceMB") & "(MB):") _
                                                                                  & IIf(IsDBNull(ObjRsPSALinux.Item("SizeMB")), "", ObjRsPSALinux.Item("SizeMB") & "(MB)" & "@*@")
                                            End If
                                        End While
                                    Loop While ObjRsPSALinux.NextResult()

                                    If Trim$(strCurrentClockSpeed) <> "" Then
                                        strCurrentClockSpeed = Mid$(strCurrentClockSpeed, 1, Len(strCurrentClockSpeed) - 3)
                                        strCurrentClockSpeed = "CPU Detail: " & strCurrentClockSpeed
                                    End If

                                    If Trim$(strHardDiskInfo) <> "" Then
                                        strHardDiskInfo = Mid$(strHardDiskInfo, 1, Len(strHardDiskInfo) - 3)
                                        strHardDiskInfo = "Harddisk Info: " & strHardDiskInfo
                                    End If
                                End If

                                ObjRsPSALinux.Close()

                            Else

                                strtOTALpHYSICALMEMORY = ObjRsAssetData.Item("tOTALpHYSICALMEMORY") & ""
                                strSYSTEMMANUFACTURER = ObjRsAssetData.Item("SYSTEMMANUFACTURER") & ""
                                strSYSTEMMODEL = ObjRsAssetData.Item("SYSTEMMODEL") & ""
                                strBIOSSerialNO = ObjRsAssetData.Item("BIOSSerialNO") & ""
                                StrNUMBEROFPROCESSORS = ObjRsAssetData.Item("NUMBEROFPROCESSORS") & ""

                                '---------------For USP CWSync_DeviceListLogin-------------
                                Dim ObjCmdDeviceLogin As New SqlCommand
                                With ObjCmdDeviceLogin
                                    .Connection = ConnItsupport247DB
                                    .CommandType = CommandType.StoredProcedure
                                    .CommandText = "CWSync_DeviceListLogin_RV"
                                    .CommandTimeout = strCmdTO
                                    .Parameters.Add("@InMemberID", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InMemberID").Value = ObjRsAssetData.Item("MemberId")
                                    .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                    ObjRsDeviceLogin = ObjCmdDeviceLogin.ExecuteReader
                                    ObjCmdDeviceLogin.Dispose()
                                End With
                                ObjCmdDeviceLogin = Nothing

                                If ObjRsDeviceLogin.HasRows Then
                                    While (ObjRsDeviceLogin.Read())
                                        strLOGGEDINUSER = ObjRsDeviceLogin.Item("LOGGEDINUSER") & ""
                                    End While
                                End If
                                ObjRsDeviceLogin.Close()
                                '---------------End USP CWSync_DeviceListLogin-------------

                                '---------------For USP CWSync_DeviceListCPU-------------
                                Dim ObjCmdDeviceCPU As New SqlCommand
                                With ObjCmdDeviceCPU
                                    .Connection = ConnITSAssetDB
                                    .CommandType = CommandType.StoredProcedure
                                    .CommandText = "CWSync_DeviceListCPU_RV"
                                    .CommandTimeout = strCmdTO
                                    .Parameters.Add("@InMemberID", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InMemberID").Value = ObjRsAssetData.Item("MemberId")
                                    .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                    ObjRsDeviceCPU = ObjCmdDeviceCPU.ExecuteReader
                                    ObjCmdDeviceCPU.Dispose()
                                End With

                                ObjCmdDeviceCPU = Nothing

                                If ObjRsDeviceCPU.HasRows Then
                                    While (ObjRsDeviceCPU.Read())
                                        strCurrentClockSpeed = strCurrentClockSpeed & IIf(IsDBNull(ObjRsDeviceCPU.Item("CurrentClockSpeed")), "", ObjRsDeviceCPU.Item("CurrentClockSpeed") & "MHz" & "@*@")
                                    End While
                                End If

                                If Trim$(strCurrentClockSpeed) <> "" Then
                                    strCurrentClockSpeed = Mid$(strCurrentClockSpeed, 1, Len(strCurrentClockSpeed) - 3)
                                    strCurrentClockSpeed = "CPU Detail: " & strCurrentClockSpeed
                                End If

                                ObjRsDeviceCPU.Close()

                                ''----------------END(CWSync_DeviceListCPU)------------------------------

                                ''---------------For USP CWSync_DeviceListLD-------------
                                Dim ObjCmdDeviceHD As New SqlCommand
                                With ObjCmdDeviceHD
                                    .Connection = ConnITSAssetDB
                                    .CommandType = CommandType.StoredProcedure
                                    .CommandText = "CWSync_DeviceListLD_RV"
                                    .CommandTimeout = strCmdTO
                                    .Parameters.Add("@InMemberID", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InMemberID").Value = ObjRsAssetData.Item("MemberId")
                                    .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                    ObjRSDeviceHD = ObjCmdDeviceHD.ExecuteReader
                                    ObjCmdDeviceHD.Dispose()
                                End With
                                ObjCmdDeviceHD = Nothing

                                If ObjRSDeviceHD.HasRows Then
                                    While (ObjRSDeviceHD.Read())
                                        strHardDiskInfo = strHardDiskInfo & IIf(IsDBNull(ObjRSDeviceHD.Item("Caption")), "", ObjRSDeviceHD.Item("Caption") & ":") _
                                                                         & IIf(IsDBNull(ObjRSDeviceHD.Item("FreespaceMB")), "", ObjRSDeviceHD.Item("FreespaceMB") & "(MB):") _
                                                                         & IIf(IsDBNull(ObjRSDeviceHD.Item("SizeMB")), "", ObjRSDeviceHD.Item("SizeMB") & "(MB)" & "@*@")
                                    End While
                                End If

                                If Trim$(strHardDiskInfo) <> "" Then
                                    strHardDiskInfo = Mid$(strHardDiskInfo, 1, Len(strHardDiskInfo) - 3)
                                    strHardDiskInfo = "Harddisk Info: " & strHardDiskInfo
                                End If

                                ObjRSDeviceHD.Close()
                                ''----------------END(CWSync_DeviceListLD)------------------------------

                                ''---------------For USP CWSync_DeviceListPathMissing_RV-------------
                                Dim ObjCmdDevicePatch As New SqlCommand
                                With ObjCmdDevicePatch
                                    .Connection = ConnITSPatchDB
                                    .CommandType = CommandType.StoredProcedure
                                    .CommandText = "CWSync_DeviceListPathMissing_RV"
                                    .CommandTimeout = strCmdTO
                                    .Parameters.Add("@InMemberID", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InMemberID").Value = ObjRsAssetData.Item("MemberId")
                                    .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                    .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                    ObjRsDevicePatch = ObjCmdDevicePatch.ExecuteReader
                                    ObjCmdDevicePatch.Dispose()
                                End With
                                ObjCmdDevicePatch = Nothing
                                If ObjRsDevicePatch.HasRows Then
                                    While (ObjRsDevicePatch.Read())
                                        strPatchesMissing = ObjRsDevicePatch.Item("COUNTUP") & ""
                                    End While
                                End If

                                ObjRsDevicePatch.Close()
                                '----------------END( USP CWSync_DeviceListPathMissing_RV)------------------------------
                            End If


                            ''---------------For Product Key-------------
                            Dim ObjCmdProdKey As New SqlCommand
                            With ObjCmdProdKey
                                ObjCmdProdKey.Connection = ConnITSPassVaultDB
                                ObjCmdProdKey.CommandType = CommandType.StoredProcedure
                                ObjCmdProdKey.CommandText = "USP_ATAPI_GetProductkey"
                                ObjCmdProdKey.CommandTimeout = strCmdTO
                                ObjCmdProdKey.Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                ObjCmdProdKey.Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                ObjRsProdKey = ObjCmdProdKey.ExecuteReader
                                ObjCmdProdKey.Dispose()
                            End With
                            ObjCmdProdKey = Nothing
                            If ObjRsProdKey.HasRows Then
                                While (ObjRsProdKey.Read())
                                    strInproductKey = strInproductKey & "  * " & IIf(IsDBNull(ObjRsProdKey.Item("ProductName")), "", ObjRsProdKey.Item("ProductName") & ":") _
                                     & IIf(IsDBNull(ObjRsProdKey.Item("Prdkey")), "", ObjRsProdKey.Item("Prdkey") & vbCrLf)
                                End While
                            End If

                            If Trim$(strInproductKey) <> "" Then
                                strInproductKey = "ProductKey Info: " & vbCrLf & strInproductKey
                            End If


                            ObjRsProdKey.Close()
                            '----------------END( USP USP_ATAPI_GetProductkey)------------------------------

                            ''---------------For WarrantyExp Date-------------
                            Dim ObjCmdWarrantyExp As New SqlCommand
                            With ObjCmdWarrantyExp
                                .Connection = ConnITSSADDB
                                .CommandType = CommandType.StoredProcedure
                                .CommandText = "USP_CW_WarrantyDetails"
                                .CommandTimeout = strCmdTO
                                .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                                .Parameters("@InRegId").Value = ObjRsAssetData.Item("RegId")
                                ObjRsWarrantyExp = ObjCmdWarrantyExp.ExecuteReader
                                ObjCmdWarrantyExp.Dispose()
                            End With
                            ObjCmdWarrantyExp = Nothing
                            If ObjRsWarrantyExp.HasRows Then
                                While (ObjRsWarrantyExp.Read())
                                    StrINWarrantyExpDt = IIf(IsDBNull(ObjRsWarrantyExp("enddate")), "", ObjRsWarrantyExp("enddate"))
                                End While
                            End If

                            ObjRsWarrantyExp.Close()
                            '----------------END( USP USP_ATAPI_GetProductkey)------------------------------

                            If CreateAndUpdateAssetProcess(ObjRsAssetData.Item("MemberID"), ObjRsAssetData.Item("SITEID"), ObjRsAssetData.Item("Regid"), ObjRsAssetData.Item("ResourceName"), ObjRsAssetData.Item("CREATEDON"), ObjRsAssetData.Item("IPAddresses") & "", ObjRsAssetData.Item("DefaultGateway") & "", ObjRsAssetData.Item("DomainRole") & "", ObjRsAssetData.Item("OSType") & "", ObjRsAssetData.Item("OSInformation") & "", strLOGGEDINUSER, strtOTALpHYSICALMEMORY, strSYSTEMMANUFACTURER, strSYSTEMMODEL, strBIOSSerialNO, StrNUMBEROFPROCESSORS, strCurrentClockSpeed, strHardDiskInfo, strPatchesMissing, strInproductKey, StrINWarrantyExpDt, ObjRsAssetData.Item("ATConfigCategType"), ObjRsAssetData.Item("ATConfigItemType")) = False Then
                                FileIO.LogNotify("AssetSyncToAutoTask", FileIO.NotifyType.ERR, "Asset Sync Failed")
                            Else

                            End If
                        Else
                            FileIO.LogNotify("AssetSyncToAutoTask", FileIO.NotifyType.INFO, "RegID Found Failed Not Considering For Process")
                        End If

                    Catch ex As Exception
                        ModLogErrorInsert(GlobalMemberId, GlobalRegID, "AssetSyncToAutoTask 1 : " & ex.Message.ToString)
                    End Try
                End While
            Else
                FileIO.LogNotify("AssetSyncToAutoTask", FileIO.NotifyType.INFO, "ATSync_DeviceList_NewDevice_RV_API No Record Found To Process..")
                If ObjRsAssetData.IsClosed = False Then
                    ObjRsAssetData.Close()
                End If
                Return True
            End If
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "AssetSyncToAutoTask 2 : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function

    Function GetAssetDataFromDB() As Boolean
        Try

            Dim ObjCmdGetLastSyncData1 As New SqlCommand
            Dim ObjCmdGetLastSyncData2 As New SqlCommand
            Dim ObjCmdGetLastSyncData3 As New SqlCommand

            Dim ObjRsLastSyncData1 As SqlDataReader
            Dim ObjRsLastSyncData2 As SqlDataReader
            Dim ObjRsLastSyncData3 As SqlDataReader

            Dim ObjCmdDevice1 As New SqlCommand
            Dim ObjCmdDevice2 As New SqlCommand
            Dim ObjCmdDevice3 As New SqlCommand

            Dim CountFlag As Integer

            '''''''''''''''' For Asset Desktop,Server & Virtual Host
            With ObjCmdGetLastSyncData1
                .Connection = ConnITSPSAAutoTaskAPI
                .CommandTimeout = strCmdTO
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_AT_Get_LatestSynDt"
                .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberId").Value = GlobalMemberId
                .Parameters.Add("@InDeviceID", System.Data.SqlDbType.Int, 4)
                .Parameters("@InDeviceID").Value = 1    '' For Asset
                ObjRsLastSyncData1 = ObjCmdGetLastSyncData1.ExecuteReader
                ObjCmdGetLastSyncData1.Dispose()
            End With


            If ObjRsLastSyncData1.HasRows Then
                While (ObjRsLastSyncData1.Read())

                    If IsDate(ObjRsLastSyncData1.Item("LastSyncDT")) = True Then
                        DTLastSyncData1 = ObjRsLastSyncData1.Item("LastSyncDT")
                        If SyncFlag = 1 Then
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId - 1", DTLastSyncData1)
                        End If
                    End If

                    If IsDate(ObjRsLastSyncData1.Item("FullSyncDt")) = True Then
                        DTFullSyncData1 = ObjRsLastSyncData1.Item("FullSyncDt")
                    End If

                End While
            Else
                If SyncFlag = 1 Then CountFlag = 1
                FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt No Data For DeviceId - 1")
            End If

            ObjRsLastSyncData1.Close()

            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId 2")


            '''''''''''''''' For Mobile
            With ObjCmdGetLastSyncData2
                .Connection = ConnITSPSAAutoTaskAPI
                .CommandTimeout = strCmdTO
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_AT_Get_LatestSynDt"
                .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberId").Value = GlobalMemberId
                .Parameters.Add("@InDeviceID", System.Data.SqlDbType.Int, 4)
                .Parameters("@InDeviceID").Value = 2    '' For Mobile
                ObjRsLastSyncData2 = ObjCmdGetLastSyncData2.ExecuteReader
                ObjCmdGetLastSyncData2.Dispose()
            End With

            If ObjRsLastSyncData2.HasRows Then
                While (ObjRsLastSyncData2.Read())

                    If IsDate(ObjRsLastSyncData2.Item("LastSyncDT")) = True Then
                        DTLastSyncData2 = ObjRsLastSyncData2.Item("LastSyncDT")
                        If SyncFlag = 1 Then
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId - 2", DTLastSyncData2)
                        End If
                    End If

                    If IsDate(ObjRsLastSyncData2.Item("FullSyncDt")) = True Then
                        DTFullSyncData2 = ObjRsLastSyncData2.Item("FullSyncDt")
                    End If

                End While
            Else
                If SyncFlag = 1 Then CountFlag = CountFlag + 1
                FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt No Data For DeviceId - 2")
            End If

            ObjRsLastSyncData2.Close()


            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId 2")

            '''''''''''''''' For Backup
            With ObjCmdGetLastSyncData3
                .Connection = ConnITSPSAAutoTaskAPI
                .CommandTimeout = strCmdTO
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_AT_Get_LatestSynDt"
                .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberId").Value = GlobalMemberId
                .Parameters.Add("@InDeviceID", System.Data.SqlDbType.Int, 4)
                .Parameters("@InDeviceID").Value = 3    '' For Mobile
                ObjRsLastSyncData3 = ObjCmdGetLastSyncData3.ExecuteReader
                ObjCmdGetLastSyncData3.Dispose()
            End With

            If ObjRsLastSyncData3.HasRows Then
                While (ObjRsLastSyncData3.Read())

                    If IsDate(ObjRsLastSyncData3.Item("LastSyncDT")) = True Then
                        DTLastSyncData3 = ObjRsLastSyncData3.Item("LastSyncDT")
                        If SyncFlag = 1 Then
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId - 3", DTLastSyncData3)
                        End If
                    End If

                    If IsDate(ObjRsLastSyncData3.Item("FullSyncDt")) = True Then
                        DTFullSyncData3 = ObjRsLastSyncData3.Item("FullSyncDt")
                    End If

                End While
            Else
                If SyncFlag = 1 Then CountFlag = CountFlag + 1
                FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt No Data For DeviceId - 3")
            End If

            ObjRsLastSyncData3.Close()

            If CountFlag < 3 Then

                FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt Count Flag Is Less Than 3", CountFlag)

                With ObjCmdDevice1
                    .Connection = ConnITSAssetDB
                    .CommandTimeout = strCmdTO
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "ATSync_DeviceList_NewDevice_RV_API"
                    .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                    .Parameters("@InMemberId").Value = GlobalMemberId
                    .Parameters.Add("@InLastFullSync", System.Data.SqlDbType.DateTime, 8)
                    If SyncFlag = 1 Then
                        .Parameters("@InLastFullSync").Value = IIf(IsDBNull(DTLastSyncData1), Now, DTLastSyncData1)
                        DTLastSyncData1 = Now
                    Else
                        If IsDate(DTFullSyncData1) = True Then

                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId - 1", DTLastSyncData1)

                            If DateDiff("d", DTFullSyncData1, Now) > 0 Then
                                FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "DateDiff > 0 Passing Null Value")
                                .Parameters("@InLastFullSync").Value = DBNull.Value
                                DTLastSyncData1 = Now
                                DTFullSyncData1 = Now
                            End If

                            If DateDiff("d", DTFullSyncData1, Now) <= 0 Then
                                FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "DateDiff <=0 Passing InLastFullSync + 1")
                                .Parameters("@InLastFullSync").Value = DateAdd("d", 1, Now)
                            End If

                        Else
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "Passing Null For Device 1")
                            .Parameters("@InLastFullSync").Value = DBNull.Value
                            DTLastSyncData1 = Now
                            DTFullSyncData1 = Now
                        End If
                    End If

                    ObjRsAssetData = ObjCmdDevice1.ExecuteReader
                    ObjCmdDevice1.Dispose()
                End With

                With ObjCmdDevice2
                    .Connection = ConnITSMDMgmtDB
                    .CommandTimeout = strCmdTO
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_MDM_AT_Device_Details"
                    .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                    .Parameters("@InMemberId").Value = GlobalMemberId
                    .Parameters.Add("@InLastFullSync", System.Data.SqlDbType.DateTime, 8)
                    If SyncFlag = 1 Then
                        .Parameters("@InLastFullSync").Value = IIf(IsDBNull(DTLastSyncData2), Now, DTLastSyncData2)
                        DTLastSyncData2 = Now
                    Else
                        If IsDate(DTFullSyncData2) = True Then
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId - 2", DTLastSyncData2)
                            If DateDiff("d", DTFullSyncData2, Now) > 0 Then
                                .Parameters("@InLastFullSync").Value = DBNull.Value
                                DTLastSyncData2 = Now
                                DTFullSyncData2 = Now
                            End If

                            If DateDiff("d", DTFullSyncData2, Now) <= 0 Then
                                .Parameters("@InLastFullSync").Value = DateAdd("d", 1, Now)
                            End If

                        Else
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "Passing Null For Device 2")
                            .Parameters("@InLastFullSync").Value = DBNull.Value
                            DTLastSyncData2 = Now
                            DTFullSyncData2 = Now
                        End If
                    End If
                    ObjRsMobileDevice = ObjCmdDevice2.ExecuteReader
                    ObjCmdDevice2.Dispose()
                End With

                With ObjCmdDevice3
                    .Connection = ConnITSRMMVaultDB
                    .CommandTimeout = strCmdTO
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_Vault_AutoTaskAPI_Data"
                    .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                    .Parameters("@InMemberId").Value = GlobalMemberId
                    .Parameters.Add("@Intype", System.Data.SqlDbType.Int, 4)
                    .Parameters("@Intype").Value = 1
                    .Parameters.Add("@InLastFullSync", System.Data.SqlDbType.DateTime, 8)
                    If SyncFlag = 1 Then
                        .Parameters("@InLastFullSync").Value = IIf(IsDBNull(DTLastSyncData3), Now, DTLastSyncData3)
                        DTLastSyncData3 = Now
                    Else
                        If IsDate(DTFullSyncData3) = True Then
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "USP_AT_Get_LatestSynDt For DeviceId - 3", DTLastSyncData3)
                            If DateDiff("d", DTFullSyncData3, Now) > 0 Then
                                .Parameters("@InLastFullSync").Value = DBNull.Value
                                DTLastSyncData3 = Now
                                DTFullSyncData3 = Now
                            End If

                            If DateDiff("d", DTFullSyncData3, Now) <= 0 Then
                                .Parameters("@InLastFullSync").Value = DateAdd("d", 1, Now)
                            End If
                        Else
                            FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "Passing Null For Device 3")
                            .Parameters("@InLastFullSync").Value = DBNull.Value
                            DTLastSyncData3 = Now
                            DTFullSyncData3 = Now
                        End If

                    End If
                    ObjRsBackUpData = ObjCmdDevice3.ExecuteReader
                    ObjCmdDevice3.Dispose()
                End With

                If ObjRsAssetData.HasRows Or ObjRsMobileDevice.HasRows Or ObjRsBackUpData.HasRows Then
                    GetAssetDataFromDB = True
                End If
            Else
                FileIO.LogNotify("GetAssetDataFromDB", FileIO.NotifyType.INFO, "Member Is For the First Time Sync It will Sync In Full Sync")
                GetAssetDataFromDB = False
            End If

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "GetAssetDataFromDB :  " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function CreateAndUpdateAssetProcess(ByVal InMemberId, ByVal InSiteId, ByVal InRegId, ByVal InResourceName, ByVal InCREATEDON, ByVal InIPAddresses, ByVal InGateWayIP, ByVal InDomainRole, ByVal InOSType, ByVal InOSInformation, ByVal InLOGGEDINUSER, ByVal IntOTALpHYSICALMEMORY, ByVal InSYSTEMMANUFACTURER, ByVal InSYSTEMMODEL, ByVal InBASEBOARDSRNO, ByVal InNUMBEROFPROCESSORS, ByVal InCurrentClockSpeed, ByVal InHardDiskInfo, ByVal InPatchesMissing, ByVal InProductKey, ByVal InWarrantyExpDT, ByVal ATConfigCategType, ByVal ATConfigItemType) As Boolean
        ''''''Install Product Configuration.
        Try

            Dim MProductId
            Dim MATAccountID
            Dim AuTotaskInstalledProductId
            Dim strQuery As String
            Dim IsMapped As String
            Dim MProductCatID
            Dim MProdConfigItemTypeID

            MProductId = ""
            MATAccountID = ""
            AuTotaskInstalledProductId = ""
            strQuery = ""
            IsMapped = ""
            MProductCatID = ""
            MProdConfigItemTypeID = ""

            Dim FoundProdCatRows() As DataRow
            Dim FoundProdConfigTypeRows() As DataRow
            Dim FoundProdRows() As DataRow

            If InSYSTEMMODEL = "" Then InSYSTEMMODEL = "Unavailable"

            FoundProdCatRows = ProdCatTable.Select("ProductCategoryName='" & UCase(ATConfigCategType) & "'")
            FoundProdConfigTypeRows = ProdConfigTable.Select("ProdConfigItemTypeName='" & UCase(ATConfigItemType) & "'")

            If FoundProdCatRows.Length > 0 And FoundProdConfigTypeRows.Length > 0 Then

                For Each row As DataRow In FoundProdCatRows
                    MProductCatID = row("ProductCategoryID")
                Next

                For Each row As DataRow In FoundProdConfigTypeRows
                    MProdConfigItemTypeID = row("ProdConfigItemTypeID")
                Next

                FoundProdRows = ProdTable.Select("ProductName='" & UCase(InSYSTEMMODEL) & "' and ProductCategoryID ='" & MProductCatID & "' And ProductAllocationCodeID ='" & ATProductAllocationCodeID & "'")

                If FoundProdRows.Length > 0 Then
                    For Each row As DataRow In FoundProdRows
                        MProductId = row("ProductId")
                    Next
                Else
                    ProdTable.Rows.Add(GlobalMemberId, "", UCase(InSYSTEMMODEL), ATProductAllocationCodeID, MProductCatID, "I", 0)
                End If

                ProdTable.AcceptChanges()

                MATAccountID = FetchAccountIdBasedOnContinuumSiteId(InMemberId, InSiteId)

                If MATAccountID >= 0 Then
                    IsMapped = GetIsMappedDeviceId(InSiteId, InRegId)
                    If IsMapped <> "" Then

                        Dim foundRows() As DataRow

                        Dim I As Integer
                        Dim StrCriteria As String
                        Dim MMappedFlag As Integer
                        MMappedFlag = 0
                        StrCriteria = ""
                        I = 1

                        If IsMigrated = True And IsMapped = "N" And (MGetUDFINFO = 1 Or MGetUDFINFO = 0) Then ' Migration Partner
                            FileIO.LogNotify("CreateAndUpdateAssetProcess", FileIO.NotifyType.INFO, "IsMigrated = True And IsMapped =N And UDF Found Continuum Device ID Mapping Based On ContinuumDeviceID")

                            StrCriteria = "ContinuumDeviceID = '" & InRegId & "'"
                            foundRows = AssetTable.Select(StrCriteria)

                            If foundRows.Length > 0 Then
                                FileIO.LogNotify("CreateAndUpdateAssetProcess", FileIO.NotifyType.INFO, "Data Found IsMigrated = True And IsMapped =N And UDF Found Continuum Device ID Mapping Based On ContinuumDeviceID")
                                MMappedFlag = 1
                            Else
                                StrCriteria = "DeviceID = '" & InRegId & "' And ATAccountID='" & MATAccountID & "'"
                                foundRows = AssetTable.Select(StrCriteria)
                                If foundRows.Length > 0 Then
                                    FileIO.LogNotify("CreateAndUpdateAssetProcess", FileIO.NotifyType.INFO, "Data Found IsMigrated = True And IsMapped =N And UDF Not Found Continuum Device ID Mapping Based On ContinuumDeviceID")
                                End If
                            End If
                        End If

                        If (IsMigrated = True And IsMapped = "Y" And (MGetUDFINFO = 1 Or MGetUDFINFO = 0)) Or IsMigrated = False Then
                            MMappedFlag = 0
                            FileIO.LogNotify("CreateAndUpdateAssetProcess", FileIO.NotifyType.INFO, "Mapping Based On UDF Device Id")
                            StrCriteria = "DeviceID = '" & InRegId & "' And ATAccountID='" & MATAccountID & "'"
                        End If

                        If StrCriteria <> "" Then
                            foundRows = AssetTable.Select(StrCriteria)
                            If foundRows.Length > 0 Then
                                For Each row As DataRow In foundRows
                                    If I = 1 Then   '' This Is added Because If We are getting Multiple Then Take First Once ' FIFO Method
                                        If ((CStr(InRegId) <> row("DeviceID")) Or (CStr(InRegId) = row("DeviceID"))) And row("InstalledProductID") <> "" Then

                                            row("ProductCategoryID") = MProductCatID
                                            row("ProdConfigItemTypeID") = MProdConfigItemTypeID
                                            row("ProductID") = MProductId
                                            row("ProductName") = UCase(InSYSTEMMODEL)
                                            'row("InstalledProductID")=
                                            row("DeviceID") = InRegId
                                            row("IUflag") = "U"
                                            row("IsMapped") = MMappedFlag
                                            row("IsProcessed") = 0
                                            row("Batch") = 0

                                            row("DeviceTypeID") = 1
                                            row("ATAccountID") = MATAccountID
                                            row("SiteID") = InSiteId

                                            row("InstallDate") = InCREATEDON
                                            row("SerialNumber") = InBASEBOARDSRNO
                                            row("ReferenceTitle") = InResourceName

                                            If IsDate(InWarrantyExpDT) = True Then
                                                row("WarrantyExpirationDate") = InWarrantyExpDT
                                            Else
                                                row("WarrantyExpirationDate") = DBNull.Value
                                            End If

                                            row("IPAddress") = InIPAddresses
                                            row("LastLoginBy") = InLOGGEDINUSER
                                            row("NOOfSecurityPatch") = InPatchesMissing
                                            row("OSType") = InOSType
                                            row("OSName") = InOSInformation
                                            row("CPUCount") = InNUMBEROFPROCESSORS
                                            row("CPUDetail") = InCurrentClockSpeed
                                            row("HardDiskInfo") = InHardDiskInfo
                                            row("MemorySize") = IntOTALpHYSICALMEMORY
                                            row("Manufacturer") = InSYSTEMMANUFACTURER
                                            row("productKeyInfo") = InProductKey
                                        End If
                                    End If
                                    I = I + 1
                                Next
                                AssetTable.AcceptChanges()
                            Else
                                AssetTable.Rows.Add(0, MProductCatID, MProdConfigItemTypeID, MProductId, UCase(InSYSTEMMODEL), "", "", InRegId, "I", 0, 0, 0, 1, MATAccountID, InSiteId, InCREATEDON, InBASEBOARDSRNO, InResourceName, InWarrantyExpDT, InIPAddresses, InLOGGEDINUSER, InPatchesMissing, InOSType, InOSInformation, InNUMBEROFPROCESSORS, InCurrentClockSpeed, InHardDiskInfo, IntOTALpHYSICALMEMORY, InSYSTEMMANUFACTURER, InProductKey)
                                AssetTable.AcceptChanges()
                            End If
                        End If
                        Return True
                    Else
                        FileIO.LogNotify("CreateAndUpdateAssetProcess", FileIO.NotifyType.INFO, "GetIsMappedDeviceId Function Retrieve Failed")
                    End If
                Else
                    FileIO.LogNotify("CreateAndUpdateAssetProcess", FileIO.NotifyType.INFO, "FetchAccountIdBasedOnContinuumSiteId Function Retrieve Failed SiteId :", InSiteId)
                    ATDeviceSyncStatusInsertUpdateDetails(InMemberId, InSiteId, InRegId, "Pending", "Site Not Found", DBNull.Value, 0)
                End If
            Else
                FileIO.LogNotify("CreateAndUpdateAssetProcess", FileIO.NotifyType.INFO, "ProductCategory / ProductConfigType Not Found ", UCase(ATConfigCategType), UCase(ATConfigItemType))
            End If

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "CreateAndUpdateAssetprocess Main Line No ." & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            ATDeviceSyncStatusInsertUpdateDetails(InMemberId, InSiteId, InRegId, "Failed", "CreateAndUpdateAssetprocess Main Line No ." & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString, DBNull.Value, 0)
            Return False
        End Try
    End Function
    Function CreateAndUpdateMobileProcess() As Boolean
        ''''''Install Product Configuration.
        Try

            While (ObjRsMobileDevice.Read())
                Try
                    GlobalRegID = ObjRsMobileDevice.Item("RegId")

                    FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "Process Start For RegId : ", ObjRsMobileDevice.Item("RegId"))
                    TotalMResCntToSync = TotalMResCntToSync + 1

                    If GetFailedRegID(ObjRsMobileDevice.Item("MemberId"), ObjRsMobileDevice.Item("RegId")) = 0 Then

                        Dim MProductId
                        Dim MATAccountID
                        Dim AuTotaskInstalledProductId
                        Dim strQuery As String
                        Dim IsMapped As String
                        Dim MProductCatID
                        Dim MProdConfigItemTypeID

                        MProductId = -1
                        MATAccountID = -1
                        AuTotaskInstalledProductId = ""
                        strQuery = ""
                        IsMapped = ""
                        MProductCatID = ""
                        MProdConfigItemTypeID = ""

                        Dim FoundProdCatRows() As DataRow
                        Dim FoundProdConfigTypeRows() As DataRow
                        Dim FoundProdRows() As DataRow


                        FoundProdCatRows = ProdCatTable.Select("ProductCategoryName='" & UCase(ObjRsMobileDevice.Item("ATConfigCategType")) & "'")
                        FoundProdConfigTypeRows = ProdConfigTable.Select("ProdConfigItemTypeName='" & UCase(ObjRsMobileDevice.Item("ATConfigItemType")) & "'")

                        If FoundProdCatRows.Length > 0 And FoundProdConfigTypeRows.Length > 0 Then
                            For Each row As DataRow In FoundProdCatRows
                                MProductCatID = row("ProductCategoryID")
                            Next

                            For Each row As DataRow In FoundProdConfigTypeRows
                                MProdConfigItemTypeID = row("ProdConfigItemTypeID")
                            Next

                            FoundProdRows = ProdTable.Select("ProductName='" & UCase(ObjRsMobileDevice.Item("ProductName")) & "' and ProductCategoryID ='" & MProductCatID & "' And ProductAllocationCodeID ='" & ATProductAllocationCodeID & "'")

                            If FoundProdRows.Length > 0 Then
                                For Each row As DataRow In FoundProdRows
                                    MProductId = row("ProductId")
                                Next
                            Else
                                ProdTable.Rows.Add(GlobalMemberId, "", UCase(ObjRsMobileDevice.Item("ProductName")), ATProductAllocationCodeID, MProductCatID, "I", 0)
                            End If

                            ProdTable.AcceptChanges()

                            MATAccountID = FetchAccountIdBasedOnContinuumSiteId(ObjRsMobileDevice.Item("MemberId"), ObjRsMobileDevice.Item("SiteId"))

                            If MATAccountID >= 0 Then
                                IsMapped = GetIsMappedDeviceId(ObjRsMobileDevice.Item("SiteId"), ObjRsMobileDevice.Item("RegId"))
                                If IsMapped <> "" Then

                                    Dim foundRows() As DataRow

                                    Dim I As Integer
                                    Dim StrCriteria As String
                                    Dim MMappedFlag As Integer
                                    MMappedFlag = 0
                                    StrCriteria = ""
                                    I = 1

                                    If IsMigrated = True And IsMapped = "N" And (MGetUDFINFO = 1 Or MGetUDFINFO = 0) Then ' Migration Partner
                                        FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "IsMigrated = True And IsMapped =N And UDF Found Continuum Device ID Mapping Based On ContinuumDeviceID")

                                        StrCriteria = "ContinuumDeviceID = '" & ObjRsMobileDevice.Item("RegId") & "'"
                                        foundRows = AssetTable.Select(StrCriteria)

                                        If foundRows.Length > 0 Then
                                            FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "Data Found IsMigrated = True And IsMapped =N And UDF Found Continuum Device ID Mapping Based On ContinuumDeviceID")
                                            MMappedFlag = 1
                                        Else
                                            StrCriteria = "DeviceID = '" & ObjRsMobileDevice.Item("RegId") & "' And ATAccountID='" & MATAccountID & "'"
                                            foundRows = AssetTable.Select(StrCriteria)
                                            If foundRows.Length > 0 Then
                                                FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "Data Found IsMigrated = True And IsMapped =N And UDF Not Found Continuum Device ID Mapping Based On ContinuumDeviceID")
                                            End If
                                        End If
                                    End If

                                    If (IsMigrated = True And IsMapped = "Y" And (MGetUDFINFO = 1 Or MGetUDFINFO = 0)) Or IsMigrated = False Then
                                        MMappedFlag = 0
                                        FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "Mapping Based On UDF Device Id")
                                        StrCriteria = "DeviceID = '" & ObjRsMobileDevice.Item("RegId") & "' And ATAccountID='" & MATAccountID & "'"
                                    End If

                                    If StrCriteria <> "" Then
                                        foundRows = AssetTable.Select(StrCriteria)
                                        If foundRows.Length > 0 Then
                                            For Each row As DataRow In foundRows
                                                If I = 1 Then   '' This Is added Because If We are getting Multiple Then Take First Once ' FIFO Method
                                                    If ((CStr(ObjRsMobileDevice.Item("RegId")) <> row("DeviceID")) Or (CStr(ObjRsMobileDevice.Item("RegId")) = row("DeviceID"))) And row("InstalledProductID") <> "" Then

                                                        row("ProductCategoryID") = MProductCatID
                                                        row("ProdConfigItemTypeID") = MProdConfigItemTypeID
                                                        row("ProductID") = MProductId
                                                        row("ProductName") = UCase(ObjRsMobileDevice.Item("ProductName"))
                                                        'row("InstalledProductID")=
                                                        row("DeviceID") = ObjRsMobileDevice.Item("RegId")
                                                        row("IUflag") = "U"
                                                        row("IsMapped") = MMappedFlag
                                                        row("IsProcessed") = 0
                                                        row("Batch") = 0

                                                        row("DeviceTypeID") = 2
                                                        row("ATAccountID") = MATAccountID
                                                        row("SiteID") = ObjRsMobileDevice.Item("SiteId")

                                                        row("InstallDate") = ObjRsMobileDevice.Item("InstalledDate")
                                                        row("SerialNumber") = ObjRsMobileDevice.Item("IMEIEsn") & ""
                                                        row("ReferenceTitle") = ObjRsMobileDevice.Item("DeviceName")

                                                        If IsDate(ObjRsMobileDevice.Item("WarrantyExpireDT")) = True Then
                                                            row("WarrantyExpirationDate") = ObjRsMobileDevice.Item("WarrantyExpireDT")
                                                        Else
                                                            row("WarrantyExpirationDate") = DBNull.Value
                                                        End If

                                                        row("OSName") = ObjRsMobileDevice.Item("OSName") & ""
                                                        row("Manufacturer") = ObjRsMobileDevice.Item("Manufacturer") & ""

                                                        row("AppleSerialNumber") = ObjRsMobileDevice.Item("AppleSerialNumber") & ""
                                                        row("ComplianceState") = ObjRsMobileDevice.Item("ComplianceState") & ""
                                                        row("CurrCarrier") = ObjRsMobileDevice.Item("CurrCarrier") & ""
                                                        row("DataRoaming") = ObjRsMobileDevice.Item("DataRoaming") & ""
                                                        row("DeviceType") = ObjRsMobileDevice.Item("DeviceType") & ""
                                                        row("FreeIntStorageInGB") = ObjRsMobileDevice.Item("FreeIntStorageInGB") & ""
                                                        row("HardwareEncryption") = ObjRsMobileDevice.Item("HardwareEncryption") & ""
                                                        row("HomeCarrier") = ObjRsMobileDevice.Item("HomeCarrier") & ""
                                                        row("DeviceJailBroken") = ObjRsMobileDevice.Item("DeviceJailBroken") & ""

                                                        If IsDate(ObjRsMobileDevice.Item("LastReported")) = True Then
                                                            row("LastReported") = ObjRsMobileDevice.Item("LastReported")
                                                        Else
                                                            row("LastReported") = DBNull.Value
                                                        End If

                                                        row("Maas360ManagedStatus") = ObjRsMobileDevice.Item("Maas360ManagedStatus") & ""
                                                        row("ModemFirmwareVersion") = ObjRsMobileDevice.Item("ModemFirmwareVersion") & ""
                                                        row("OSVersion") = ObjRsMobileDevice.Item("OSVersion") & ""
                                                        row("Ownership") = ObjRsMobileDevice.Item("Ownership") & ""
                                                        row("PlatformName") = ObjRsMobileDevice.Item("PlatformName") & ""
                                                        row("MDMPolicy") = ObjRsMobileDevice.Item("MDMPolicy") & ""
                                                        row("TotIntStorageInGB") = ObjRsMobileDevice.Item("TotIntStorageInGB") & ""
                                                        row("WiFiMacAddress") = ObjRsMobileDevice.Item("WiFiMacAddress") & ""

                                                    End If
                                                End If
                                                I = I + 1
                                            Next
                                            AssetTable.AcceptChanges()
                                        Else
                                            AssetTable.Rows.Add(0, MProductCatID, MProdConfigItemTypeID, MProductId, UCase(ObjRsMobileDevice.Item("ProductName")), "", "", ObjRsMobileDevice.Item("RegId"), "I", 0, 0, 0, 2, MATAccountID, ObjRsMobileDevice.Item("SiteId"), ObjRsMobileDevice.Item("InstalledDate"), ObjRsMobileDevice.Item("IMEIEsn"), ObjRsMobileDevice.Item("DeviceName"), ObjRsMobileDevice.Item("WarrantyExpireDT"), "", "", "", "", ObjRsMobileDevice.Item("OSName"), "", "", "", "", ObjRsMobileDevice.Item("Manufacturer"), "", ObjRsMobileDevice.Item("AppleSerialNumber"), ObjRsMobileDevice.Item("ComplianceState"), ObjRsMobileDevice.Item("CurrCarrier"), ObjRsMobileDevice.Item("DataRoaming"), ObjRsMobileDevice.Item("DeviceType"), ObjRsMobileDevice.Item("FreeIntStorageInGB"), ObjRsMobileDevice.Item("HardwareEncryption"), ObjRsMobileDevice.Item("HomeCarrier"), ObjRsMobileDevice.Item("DeviceJailBroken"), ObjRsMobileDevice.Item("LastReported"), ObjRsMobileDevice.Item("Maas360ManagedStatus"), ObjRsMobileDevice.Item("ModemFirmwareVersion"), ObjRsMobileDevice.Item("OSVersion"), ObjRsMobileDevice.Item("Ownership"), ObjRsMobileDevice.Item("PlatformName"), ObjRsMobileDevice.Item("MDMPolicy"), ObjRsMobileDevice.Item("TotIntStorageInGB"), ObjRsMobileDevice.Item("WiFiMacAddress"))
                                            AssetTable.AcceptChanges()
                                        End If
                                    End If
                                Else
                                    FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "GetIsMappedDeviceId Function Retrieve Failed")
                                End If
                            Else
                                ATDeviceSyncStatusInsertUpdateDetails(ObjRsMobileDevice.Item("MemberId"), ObjRsMobileDevice.Item("SiteId"), ObjRsMobileDevice.Item("RegId"), "Pending", "Site Not Found", DBNull.Value, 0)
                            End If
                        Else
                            FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "ProductCategory / ProductConfigType Not Found ", UCase(ObjRsMobileDevice.Item("ATConfigCategType")), UCase(ObjRsMobileDevice.Item("ATConfigItemType")))
                        End If
                    Else
                        FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "RegID Found Failed Not Considering For Process")
                    End If
                Catch ex As Exception
                    Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
                    Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
                    ATDeviceSyncStatusInsertUpdateDetails(ObjRsMobileDevice.Item("MemberId"), ObjRsMobileDevice.Item("SiteId"), ObjRsMobileDevice.Item("RegId"), "Failed", "Record Inner Loop " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString, DBNull.Value, 0)
                End Try
            End While
            ObjRsMobileDevice.Close()
            Return True
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "CreateAndUpdateMobileProcess" & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            ObjRsMobileDevice.Close()
            Return False
        End Try
    End Function
    Function CreateAndUpdateVaultDataProcess() As Boolean
        ''''''Install Product Configuration.
        Try

            While (ObjRsBackUpData.Read())
                Try
                    GlobalRegID = ObjRsBackUpData.Item("RegId")
                    FileIO.LogNotify("CreateAndUpdateMobileData", FileIO.NotifyType.INFO, "Process Start For RegId : ", ObjRsBackUpData.Item("RegId"))

                    TotalVResCntToSync = TotalVResCntToSync + 1

                    If GetFailedRegID(ObjRsBackUpData.Item("MemberId"), ObjRsBackUpData.Item("RegId")) = 0 Then

                        Dim MProductId
                        Dim MATAccountID
                        Dim AuTotaskInstalledProductId
                        Dim strQuery As String
                        Dim IsMapped As String
                        Dim MProductCatID
                        Dim MProdConfigItemTypeID

                        MProductId = -1
                        MATAccountID = -1
                        AuTotaskInstalledProductId = ""
                        strQuery = ""
                        IsMapped = ""
                        MProductCatID = ""
                        MProdConfigItemTypeID = ""

                        Dim FoundProdCatRows() As DataRow
                        Dim FoundProdConfigTypeRows() As DataRow
                        Dim FoundProdRows() As DataRow

                        FoundProdCatRows = ProdCatTable.Select("ProductCategoryName='" & UCase(ObjRsBackUpData.Item("ATProductCategory")) & "'")
                        FoundProdConfigTypeRows = ProdConfigTable.Select("ProdConfigItemTypeName='" & UCase(ObjRsBackUpData.Item("ATConfigItemType")) & "'")

                        If FoundProdCatRows.Length > 0 And FoundProdConfigTypeRows.Length > 0 Then
                            For Each row As DataRow In FoundProdCatRows
                                MProductCatID = row("ProductCategoryID")
                            Next

                            For Each row As DataRow In FoundProdConfigTypeRows
                                MProdConfigItemTypeID = row("ProdConfigItemTypeID")
                            Next

                            FoundProdRows = ProdTable.Select("ProductName='" & UCase(ObjRsBackUpData.Item("ProductName")) & "' and ProductCategoryID ='" & MProductCatID & "' And ProductAllocationCodeID ='" & ATProductAllocationCodeID & "'")

                            If FoundProdRows.Length > 0 Then
                                For Each row As DataRow In FoundProdRows
                                    MProductId = row("ProductId")
                                Next
                            Else
                                ProdTable.Rows.Add(GlobalMemberId, "", UCase(ObjRsBackUpData.Item("ProductName")), ATProductAllocationCodeID, MProductCatID, "I", 0)
                            End If

                            ProdTable.AcceptChanges()

                            MATAccountID = FetchAccountIdBasedOnContinuumSiteId(ObjRsBackUpData.Item("MemberId"), ObjRsBackUpData.Item("SiteId"))

                            If MATAccountID >= 0 Then
                                IsMapped = GetIsMappedDeviceId(ObjRsBackUpData.Item("SiteId"), ObjRsBackUpData.Item("RegId"))
                                If IsMapped <> "" Then

                                    Dim foundRows() As DataRow

                                    Dim I As Integer
                                    Dim StrCriteria As String
                                    Dim MMappedFlag As Integer
                                    MMappedFlag = 0
                                    StrCriteria = ""
                                    I = 1

                                    If IsMigrated = True And IsMapped = "N" And (MGetUDFINFO = 1 Or MGetUDFINFO = 0) Then ' Migration Partner
                                        FileIO.LogNotify("CreateAndUpdateVaultDataProcess", FileIO.NotifyType.INFO, "IsMigrated = True And IsMapped =N And UDF Found Continuum Device ID Mapping Based On ContinuumDeviceID")

                                        StrCriteria = "ContinuumDeviceID = '" & ObjRsBackUpData.Item("RegId") & "'"
                                        foundRows = AssetTable.Select(StrCriteria)

                                        If foundRows.Length > 0 Then
                                            FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "Data Found IsMigrated = True And IsMapped =N And UDF Found Continuum Device ID Mapping Based On ContinuumDeviceID")
                                            MMappedFlag = 1
                                        Else
                                            StrCriteria = "DeviceID = '" & ObjRsBackUpData.Item("RegId") & "' And ATAccountID='" & MATAccountID & "'"
                                            foundRows = AssetTable.Select(StrCriteria)
                                            If foundRows.Length > 0 Then
                                                FileIO.LogNotify("CreateAndUpdateMobileProcess", FileIO.NotifyType.INFO, "Data Found IsMigrated = True And IsMapped =N And UDF Not Found Continuum Device ID Mapping Based On ContinuumDeviceID")
                                            End If
                                        End If

                                    End If

                                    If (IsMigrated = True And IsMapped = "Y" And (MGetUDFINFO = 1 Or MGetUDFINFO = 0)) Or IsMigrated = False Then
                                        MMappedFlag = 0
                                        FileIO.LogNotify("CreateAndUpdateVaultDataProcess", FileIO.NotifyType.INFO, "Mapping Based On UDF Device Id")
                                        StrCriteria = "DeviceID = '" & ObjRsBackUpData.Item("RegId") & "' and ATAccountID='" & MATAccountID & "'"
                                    End If

                                    If StrCriteria <> "" Then
                                        foundRows = AssetTable.Select(StrCriteria)
                                        If foundRows.Length > 0 Then
                                            For Each row As DataRow In foundRows
                                                If I = 1 Then   '' This Is added Because If We are getting Multiple Then Take First Once ' FIFO Method
                                                    If ((CStr(ObjRsBackUpData.Item("RegId")) <> row("DeviceID")) Or (CStr(ObjRsBackUpData.Item("RegId")) = row("DeviceID"))) And row("InstalledProductID") <> "" Then
                                                        row("ProductCategoryID") = MProductCatID
                                                        row("ProdConfigItemTypeID") = MProdConfigItemTypeID
                                                        row("ProductID") = MProductId
                                                        row("ProductName") = UCase(ObjRsBackUpData.Item("ProductName"))
                                                        'row("InstalledProductID")=
                                                        row("DeviceID") = ObjRsBackUpData.Item("RegId")
                                                        row("IUflag") = "U"
                                                        row("IsMapped") = MMappedFlag
                                                        row("IsProcessed") = 0
                                                        row("Batch") = 0

                                                        row("DeviceTypeID") = 3
                                                        row("ATAccountID") = MATAccountID
                                                        row("SiteID") = ObjRsBackUpData.Item("SiteId")

                                                        row("InstallDate") = ObjRsBackUpData.Item("VaultagentProvisionDate")
                                                        row("SerialNumber") = ""
                                                        row("ReferenceTitle") = ObjRsBackUpData.Item("Resourcename")
                                                        row("WarrantyExpirationDate") = DBNull.Value

                                                        row("OrderNumber") = ObjRsBackUpData.Item("OrderNumber") & ""
                                                        row("ProvisionID") = ObjRsBackUpData.Item("ProvisionID") & ""
                                                        row("ApplianceSerialNumber") = ObjRsBackUpData.Item("ApplianceSerialNumber") & ""

                                                        If IsDate(ObjRsBackUpData.Item("ApplianceProvisionDate")) = True Then
                                                            row("ApplianceProvisionDate") = ObjRsBackUpData.Item("ApplianceProvisionDate")
                                                        Else
                                                            row("ApplianceProvisionDate") = DBNull.Value
                                                        End If

                                                        row("OffsiteStorage") = ObjRsBackUpData.Item("OffsiteStorage") & ""
                                                        row("VaultAgentID") = ObjRsBackUpData.Item("VaultAgentID") & ""
                                                        row("HostName") = ObjRsBackUpData.Item("HostName") & ""
                                                        row("VolumeName") = ObjRsBackUpData.Item("VolumeName") & ""

                                                        If IsDate(ObjRsBackUpData.Item("LastRecoveryAtAppliance")) = True Then
                                                            row("LastRecoveryAtAppliance") = ObjRsBackUpData.Item("LastRecoveryAtAppliance")
                                                        Else
                                                            row("LastRecoveryAtAppliance") = DBNull.Value
                                                        End If

                                                        row("TotalSizeAtAppliance") = ObjRsBackUpData.Item("TotalSizeAtAppliance") & ""

                                                        If IsDate(ObjRsBackUpData.Item("LastOffsiteRecovery")) = True Then
                                                            row("LastOffsiteRecovery") = ObjRsBackUpData.Item("LastOffsiteRecovery")
                                                        Else
                                                            row("LastOffsiteRecovery") = DBNull.Value
                                                        End If

                                                        row("TotalSizeOffsite") = ObjRsBackUpData.Item("TotalSizeOffsite") & ""

                                                    End If
                                                End If
                                                I = I + 1
                                            Next
                                            AssetTable.AcceptChanges()
                                        Else
                                            AssetTable.Rows.Add(0, MProductCatID, MProdConfigItemTypeID, MProductId, UCase(ObjRsBackUpData.Item("ProductName")), "", "", ObjRsBackUpData.Item("RegId"), "I", 0, 0, 0, 3, MATAccountID, ObjRsBackUpData.Item("SiteId"), ObjRsBackUpData.Item("VaultagentProvisionDate"), "", ObjRsBackUpData.Item("Resourcename"), DBNull.Value, "", "", "", "", ObjRsBackUpData.Item("OperatingSystem"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", DBNull.Value, "", "", "", "", "", "", "", "", ObjRsBackUpData.Item("OrderNumber"), ObjRsBackUpData.Item("ProvisionID"), ObjRsBackUpData.Item("ApplianceSerialNumber"), ObjRsBackUpData.Item("ApplianceProvisionDate"), ObjRsBackUpData.Item("OffsiteStorage"), ObjRsBackUpData.Item("VaultAgentID"), ObjRsBackUpData.Item("HostName"), ObjRsBackUpData.Item("VolumeName"), ObjRsBackUpData.Item("LastRecoveryAtAppliance"), ObjRsBackUpData.Item("TotalSizeAtAppliance"), ObjRsBackUpData.Item("LastOffsiteRecovery"), ObjRsBackUpData.Item("TotalSizeOffsite"))
                                        End If
                                    End If
                                Else
                                    FileIO.LogNotify("CreateAndUpdateVaultDataProcess", FileIO.NotifyType.INFO, "GetIsMappedDeviceId Function Retrieve Failed")
                                End If
                            Else
                                ATDeviceSyncStatusInsertUpdateDetails(ObjRsBackUpData.Item("MemberId"), ObjRsBackUpData.Item("SiteId"), ObjRsBackUpData.Item("RegId"), "Pending", "Site Not Found", DBNull.Value, 0)
                            End If
                        Else
                            FileIO.LogNotify("CreateAndUpdateVaultDataProcess", FileIO.NotifyType.INFO, "ProductCategory / ProductConfigType Not Found ", UCase(ObjRsBackUpData.Item("ATConfigCategType")), UCase(ObjRsBackUpData.Item("ATConfigItemType")))
                        End If
                    Else
                        FileIO.LogNotify("CreateAndUpdateVaultDataProcess", FileIO.NotifyType.INFO, "RegID Found Failed Not Considering For Process")
                    End If

                Catch ex As Exception
                    Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
                    Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
                    ATDeviceSyncStatusInsertUpdateDetails(ObjRsBackUpData.Item("MemberId"), ObjRsBackUpData.Item("SiteId"), ObjRsBackUpData.Item("RegId"), "Failed", "Record Inner Loop " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString, DBNull.Value, 0)
                End Try
            End While
            ObjRsBackUpData.Close()
            Return True
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "CreateAndUpdateVaultDataProcess" & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            ObjRsBackUpData.Close()
            Return False
        End Try
    End Function
    Function ProcessOFMakingBatch() As Boolean
        Try

            Dim foundRows() As DataRow

            '' For Asset Batch Insert
            Icount200 = 0
            foundRows = AssetTable.Select("IUFlag='I'", "DeviceTypeID")
            For Each row As DataRow In foundRows
                Icount200 = Icount200 + 1
                If InsertBatch = 0 Then InsertBatch = 1
                row("Batch") = InsertBatch
                If Icount200 >= MaxInsUpdBatch Then
                    Icount200 = 0
                    InsertBatch = InsertBatch + 1
                End If
            Next
            AssetTable.AcceptChanges()

            '' For Asset Batch Update
            Ucount200 = 0
            foundRows = AssetTable.Select("IUFlag='U'", "DeviceTypeID")
            For Each row As DataRow In foundRows
                Ucount200 = Ucount200 + 1
                If UpdateBatch = 0 Then UpdateBatch = 1
                row("Batch") = UpdateBatch
                If Ucount200 >= MaxInsUpdBatch Then
                    Ucount200 = 0
                    UpdateBatch = UpdateBatch + 1
                End If
            Next
            AssetTable.AcceptChanges()

            ProcessOFMakingBatch = True

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "ProcessOFMakingBatch : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function CreateNewProductInBatchOf200() As Boolean
        Try

            Dim sResponse As ATWSResponse
            Dim ProductToCreate As Product
            Dim I As Integer
            Dim Icount200 As Integer
            Dim AR As Long
            Dim foundRows() As DataRow

            '' For Product Batch
            foundRows = ProdTable.Select("IUFlag='I'")
            For Each row As DataRow In foundRows
                Icount200 = Icount200 + 1
                If PInsertBatch = 0 Then PInsertBatch = 1
                row("Batch") = PInsertBatch
                If Icount200 >= MaxInsUpdBatch Then
                    Icount200 = 0
                    PInsertBatch = PInsertBatch + 1
                End If
            Next
            ProdTable.AcceptChanges()

            FileIO.LogNotify("CreateNewProductInBatchOf200", FileIO.NotifyType.INFO, "Total Batch :", PInsertBatch)

            For I = 1 To PInsertBatch
                Try
                    AR = 0
                    foundRows = ProdTable.Select("IUFlag = 'I' And Batch=" & I & "")

                    If foundRows.Length > 0 Then
                        FileIO.LogNotify("CreateNewProductInBatchOf200", FileIO.NotifyType.INFO, "Process Start For Batch : ", I, foundRows.Length)

                        Dim ProductArray(foundRows.Length - 1) As Product

                        For Each row As DataRow In foundRows

                            ProductToCreate = New Product

                            ProductToCreate.Active = True
                            ProductToCreate.id = 0
                            ProductToCreate.Name = row("ProductName")
                            ProductToCreate.ProductAllocationCodeID = ATProductAllocationCodeID
                            ProductToCreate.ProductCategory = row("ProductCategoryID")
                            ProductToCreate.Serialized = True
                            ProductToCreate.id = 0                        ' Will get Autogenerated Field

                            ProductArray(AR) = ProductToCreate
                            AR = AR + 1

                        Next
                        Dim entityArray() As Entity = CType(ProductArray, Entity())
                        sResponse = myService.create(entityArray)

                        If sResponse.ReturnCode = 1 Then
                            For Each field As autotaskwebservices.Product In sResponse.EntityResults
                                FileIO.LogNotify("CreateNewProductInBatchOf200", FileIO.NotifyType.INFO, "Product For Batch -> : " & I & " ProductID : ", field.id)
                                foundRows = ProdTable.Select("ProductName = '" & Trim(UCase(field.Name)) & "' and ProductAllocationCodeID ='" & ATProductAllocationCodeID & "' and ProductCategoryID='" & field.ProductCategory & "'")
                                If foundRows.Length > 0 Then
                                    For Each row As DataRow In foundRows
                                        row("ProductId") = field.id
                                    Next
                                End If
                            Next
                            ProdTable.AcceptChanges()
                            CreateNewProductInBatchOf200 = True
                        Else
                            For Each ent As autotaskwebservices.ATWSError In sResponse.Errors
                                FileIO.LogNotify("CreateNewProductInBatchOf200", FileIO.NotifyType.ERR, "ATWSError :" & ent.Message)
                                ModLogErrorInsert(GlobalMemberId, 0, "CreateNewProductInBatchOf200 Failed For BatchID : " & InsertBatch & "ATWSError :" & ent.Message)
                                CreateNewProductInBatchOf200 = False
                            Next
                        End If
                    Else
                        FileIO.LogNotify("CreateNewProductInBatchOf200", FileIO.NotifyType.INFO, "No data Found For Batch 200 Creation")
                        CreateNewProductInBatchOf200 = True
                    End If

                Catch ex As Exception
                    Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
                    Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
                    ModLogErrorInsert(GlobalMemberId, GlobalRegID, "CreateNewProductInBatchOf200 Failed For BatchID : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
                    CreateNewProductInBatchOf200 = False
                End Try
            Next
            CreateNewProductInBatchOf200 = True
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "CreateNewProductInBatchOf200 : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
            Exit Function
        End Try
    End Function
    Function CreateAssetInBatchOf200() As Long
        'Creating(New Account) '  {Local Term :Account} Name
        Try

            Dim sResponse As ATWSResponse
            Dim InstalledProductToCreate As InstalledProduct
            Dim AR As Long
            Dim NoOfProcess As Integer
            Dim MBatchRowId As Long
            Dim foundRows() As DataRow
            Dim FoundProdRows() As DataRow
            Dim FoundDeviceID() As DataRow

            FileIO.LogNotify("CreateAssetInBatchOf200", FileIO.NotifyType.INFO, "Total Batch :", InsertBatch)

            For I = 1 To InsertBatch
                NoOfProcess = 0
Reprocess:
                Try
                    foundRows = AssetTable.Select("IUFlag = 'I' And Batch=" & I & "", "DeviceTypeID")

                    If foundRows.Length > 0 Then
                        AR = 0
                        MBatchRowId = 0

                        FileIO.LogNotify("CreateAssetInBatchOf200", FileIO.NotifyType.INFO, "Process Start For Batch : ", I, foundRows.Length)

                        Dim InstalledArray(foundRows.Length - 1) As InstalledProduct

                        For Each row As DataRow In foundRows
                            MBatchRowId = MBatchRowId + 1
                            row("BatchRowID") = MBatchRowId

                            FileIO.LogNotify("CreateAssetInBatchOf200", FileIO.NotifyType.INFO, "Adding Device ID In Batch: ", row("DeviceID"))

                            InstalledProductToCreate = New InstalledProduct

                            Dim udf As New UserDefinedField
                            Dim udf1 As New UserDefinedField
                            Dim udf2 As New UserDefinedField
                            Dim udf3 As New UserDefinedField

                            Dim udf4 As New UserDefinedField
                            Dim udf5 As New UserDefinedField
                            Dim udf6 As New UserDefinedField
                            Dim udf7 As New UserDefinedField
                            Dim udf8 As New UserDefinedField
                            Dim udf9 As New UserDefinedField
                            Dim udf10 As New UserDefinedField
                            Dim udf11 As New UserDefinedField
                            Dim udf12 As New UserDefinedField
                            Dim udf13 As New UserDefinedField
                            Dim udf14 As New UserDefinedField
                            Dim udf15 As New UserDefinedField
                            Dim udf16 As New UserDefinedField
                            Dim udf17 As New UserDefinedField
                            Dim udf18 As New UserDefinedField
                            Dim udf19 As New UserDefinedField
                            Dim udf20 As New UserDefinedField
                            Dim udf21 As New UserDefinedField
                            Dim udf22 As New UserDefinedField
                            Dim udf23 As New UserDefinedField

                            Dim udf24 As New UserDefinedField
                            Dim udf25 As New UserDefinedField
                            Dim udf26 As New UserDefinedField
                            Dim udf27 As New UserDefinedField
                            Dim udf28 As New UserDefinedField
                            Dim udf29 As New UserDefinedField
                            Dim udf30 As New UserDefinedField
                            Dim udf31 As New UserDefinedField
                            Dim udf32 As New UserDefinedField
                            Dim udf33 As New UserDefinedField
                            Dim udf34 As New UserDefinedField
                            Dim udf35 As New UserDefinedField
                            Dim udf36 As New UserDefinedField

                            InstalledProductToCreate.AccountID = row("ATAccountID")
                            InstalledProductToCreate.Active = True
                            InstalledProductToCreate.CreateDate = Now
                            InstalledProductToCreate.id = 0
                            InstalledProductToCreate.InstallDate = row("InstallDate")

                            FoundProdRows = ProdTable.Select("ProductName='" & UCase(row("ProductName")) & "' and ProductCategoryID ='" & row("ProductCategoryID") & "' And ProductAllocationCodeID ='" & ATProductAllocationCodeID & "'")

                            If FoundProdRows.Length > 0 Then
                                For Each Prow As DataRow In FoundProdRows
                                    InstalledProductToCreate.ProductID = Prow("ProductID")
                                Next
                            Else
                                FileIO.LogNotify("GoforBatchUpdate", FileIO.NotifyType.INFO, row("DeviceID") & " Product ID Not found For ProductName,CategoryID ", UCase(row("ProductName")), row("ProductCategoryID"))
                            End If

                            InstalledProductToCreate.SerialNumber = row("SerialNumber")

                            InstalledProductToCreate.ReferenceTitle = row("ReferenceTitle")

                            If IsDate(row("WarrantyExpirationDate")) = True Then
                                InstalledProductToCreate.WarrantyExpirationDate = row("WarrantyExpirationDate")
                            End If

                            Dim MNotes As String
                            MNotes = ""

                            If row("DeviceTypeID") = 1 Then

                                If Trim(row("OSType")) <> "" Then
                                    MNotes = "OSType: " & row("OSType") & " " & vbCrLf
                                End If

                                If Trim(row("OSName")) <> "" Then
                                    MNotes = "OSDetail: " & row("OSName") & " " & vbCrLf
                                End If

                                If Trim(row("CPUCount")) <> "" Then
                                    MNotes = MNotes & "CPU Count: " & row("CPUCount") & " " & vbCrLf
                                End If

                                If Trim(row("CPUDetail")) <> "" Then
                                    MNotes = MNotes & row("CPUDetail") & " " & vbCrLf
                                End If

                                If Trim(row("MemorySize")) <> "" Then
                                    MNotes = MNotes & "Memory Size: " & row("MemorySize") & " " & vbCrLf
                                End If

                                If Trim(row("Manufacturer")) <> "" Then
                                    MNotes = MNotes & "Manufacturer: " & row("Manufacturer") & " " & vbCrLf
                                End If

                                If Trim(row("HardDiskInfo")) <> "" Then
                                    MNotes = MNotes & row("HardDiskInfo") & " " & vbCrLf
                                End If

                                If Trim(row("productKeyInfo")) <> "" Then
                                    MNotes = MNotes & row("productKeyInfo") & " " & vbCrLf
                                End If

                                If MNotes <> "" Then
                                    InstalledProductToCreate.Notes = MNotes
                                End If

                            End If

                            InstalledProductToCreate.Type = row("ProdConfigItemTypeID")

                            udf.Name = "Device ID"
                            udf.Value = row("DeviceID")

                            udf1.Name = "IP Address"
                            udf1.Value = row("IPAddress")

                            udf2.Name = "LastLoginBy"
                            udf2.Value = row("LastLoginBy")

                            udf3.Name = "Nos Of Security Patch Missing"
                            udf3.Value = row("NOOfSecurityPatch")


                            ''  from here Start Mobile Data UDF value

                            udf4.Name = "Apple Serial Number"
                            udf4.Value = row("AppleSerialNumber") & ""

                            udf5.Name = "Compliance State"
                            udf5.Value = row("ComplianceState") & ""

                            udf6.Name = "Current Carrier"
                            udf6.Value = row("CurrCarrier") & ""

                            udf7.Name = "Data Roaming?"
                            udf7.Value = row("DataRoaming") & ""

                            udf8.Name = "Device Type"
                            udf8.Value = row("DeviceType") & ""

                            udf9.Name = "Free Storage (GB)"
                            udf9.Value = row("FreeIntStorageInGB") & ""

                            udf10.Name = "Hardware Encryption"
                            udf10.Value = row("HardwareEncryption") & ""

                            udf11.Name = "Home Carrier"
                            udf11.Value = row("HomeCarrier") & ""

                            udf12.Name = "Jailbroken or Rooted?"
                            udf12.Value = row("DeviceJailBroken") & ""

                            udf13.Name = "Last Reported"
                            If IsDBNull(row("LastReported")) = False Then
                                udf13.Value = row("LastReported")
                            End If

                            udf14.Name = "Management Status"
                            udf14.Value = row("Maas360ManagedStatus") & ""

                            udf15.Name = "Manufacturer"
                            If row("DeviceTypeId") = 2 Then
                                udf15.Value = row("Manufacturer") & ""
                            Else
                                udf15.Value = ""
                            End If

                            udf16.Name = "Modern Firmware Version"
                            udf16.Value = row("ModemFirmwareVersion") & ""

                            udf17.Name = "OS Name"

                            If row("DeviceTypeId") = 2 Then
                                udf17.Value = row("OSName") & ""
                            Else
                                udf17.Value = ""
                            End If

                            udf18.Name = "OS Version"
                            udf18.Value = row("OSVersion") & ""

                            udf19.Name = "Ownership"
                            udf19.Value = row("Ownership") & ""

                            udf20.Name = "Platform Name"
                            udf20.Value = row("PlatformName") & ""

                            udf21.Name = "Policy"
                            udf21.Value = row("MDMPolicy") & ""

                            udf22.Name = "Total Storage (GB)"
                            udf22.Value = row("TotIntStorageInGB") & ""

                            udf23.Name = "WiFi Mac Address"
                            udf23.Value = row("WiFiMacAddress") & ""

                            ' For Vault Data

                            udf24.Name = "Order Number"
                            udf24.Value = row("OrderNumber") & ""

                            udf25.Name = "Provision ID"
                            udf25.Value = row("ProvisionID") & ""

                            udf26.Name = "Appliance Serial Number"
                            udf26.Value = row("ApplianceSerialNumber") & ""

                            udf27.Name = "Appliance Provision Date"
                            If IsDBNull(row("ApplianceProvisionDate")) = False Then
                                udf27.Value = row("ApplianceProvisionDate")
                            End If

                            udf28.Name = "Operating System"

                            If row("DeviceTypeID") = 3 Then
                                udf28.Value = row("OSName") & ""
                            Else
                                udf28.Value = ""
                            End If

                            udf29.Name = "Offsite Storage?"
                            udf29.Value = row("OffsiteStorage") & ""

                            udf30.Name = "Vault Agent ID"
                            udf30.Value = row("VaultAgentID") & ""

                            udf31.Name = "Host Name"
                            udf31.Value = row("HostName") & ""

                            udf32.Name = "Volume Name"
                            udf32.Value = row("VolumeName") & ""

                            udf33.Name = "Last Recovery At Appliance"
                            If IsDBNull(row("LastRecoveryAtAppliance")) = False Then
                                udf33.Value = row("LastRecoveryAtAppliance")
                            End If

                            udf34.Name = "Total Size At Appliance (GB)"
                            udf34.Value = row("TotalSizeAtAppliance") & ""

                            udf35.Name = "Last Offsite Recovery"
                            If IsDBNull(row("LastOffsiteRecovery")) = False Then
                                udf35.Value = row("LastOffsiteRecovery")
                            End If

                            udf36.Name = "Total Size Offsite (GB)"
                            udf36.Value = row("TotalSizeOffsite") & ""

                            InstalledProductToCreate.UserDefinedFields = New UserDefinedField() {udf, udf1, udf2, udf3, udf4, udf5, udf6, udf7, udf8, udf9, udf10, udf11, udf12, udf13, udf14, udf15, udf16, udf17, udf18, udf19, udf20, udf21, udf22, udf23, udf24, udf25, udf26, udf27, udf28, udf29, udf30, udf31, udf32, udf33, udf34, udf35, udf36}

                            InstalledArray(AR) = InstalledProductToCreate
                            AR = AR + 1
                        Next

                        Dim entityArray() As Entity = CType(InstalledArray, Entity())
                        sResponse = myService.create(entityArray)
                        If sResponse.ReturnCode = 1 Then
                            For Each field As autotaskwebservices.InstalledProduct In sResponse.EntityResults
                                Dim UDFLength As Integer
                                Dim MUdfDeviceID As String
                                MUdfDeviceID = ""
                                UDFLength = field.UserDefinedFields.Length

                                If UDFLength > 0 Then
                                    Dim U As Integer
                                    For U = 0 To (UDFLength - 1)
                                        If field.UserDefinedFields(U).Name = "Device ID" Then
                                            MUdfDeviceID = field.UserDefinedFields(U).Value
                                            Exit For
                                        End If
                                    Next
                                End If

                                FileIO.LogNotify("CreateAssetInBatchOf200", FileIO.NotifyType.INFO, "Account For Batch -> : " & I & " DeviceID, InstalledProductID : ", MUdfDeviceID, field.id)

                                FoundDeviceID = AssetTable.Select("DeviceID='" & MUdfDeviceID & "'")
                                If FoundDeviceID.Length > 0 Then
                                    For Each row As DataRow In FoundDeviceID

                                        If row("DeviceTypeID") = "1" Then
                                            TotalAResCntSync = TotalAResCntSync + 1
                                        End If

                                        If row("DeviceTypeID") = "2" Then
                                            TotalMResCntSync = TotalMResCntSync + 1
                                        End If

                                        If row("DeviceTypeID") = "3" Then
                                            TotalVResCntSync = TotalVResCntSync + 1
                                        End If

                                        ATDeviceSyncStatusInsertUpdateDetails(GlobalMemberId, row("SiteId"), MUdfDeviceID, "Success", "", field.id, row("IsMapped"))

                                    Next
                                End If
                            Next
                        Else

                            Dim StrErrorMessage
                            Dim Maxlength As Long
                            Dim FindPosition As Long
                            Dim MRecordNumber As Long
                            For Each ent As autotaskwebservices.ATWSError In sResponse.Errors     ' The Below Code is to remove Bad records from batch.
                                If InStr(ent.Message, "record number") > 0 Then
                                    Maxlength = Len(ent.Message)
                                    FindPosition = InStr(ent.Message, "[")
                                    If FindPosition > 0 Then
                                        StrErrorMessage = Mid(ent.Message, FindPosition, Maxlength - FindPosition)
                                        StrErrorMessage = Replace(Replace(StrErrorMessage, "[", ""), "]", "")
                                        If IsNumeric(StrErrorMessage) = True Then
                                            MRecordNumber = CInt(StrErrorMessage)
                                            foundRows = AssetTable.Select("IUFlag = 'I' And Batch=" & I & " And BatchRowID=" & MRecordNumber & "")
                                            If foundRows.Length > 0 Then
                                                For Each row As DataRow In foundRows
                                                    row("IUFlag") = ""
                                                    ATDeviceSyncStatusInsertUpdateDetails(GlobalMemberId, row("SiteId"), row("DeviceID"), "Failed", ent.Message, "", row("IsMapped"))
                                                Next
                                                ModLogErrorInsert(GlobalMemberId, 0, "CreateAssetInBatchOf200 Failed For BatchID : " & I & "ATWSError :" & ent.Message)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            If NoOfProcess = 0 Then
                                NoOfProcess = NoOfProcess + 1
                                GoTo Reprocess
                            End If
                        End If
                    Else
                        FileIO.LogNotify("CreateAssetInBatchOf200", FileIO.NotifyType.INFO, "No data Found For Batch 200 Creation")
                    End If
                Catch ex As Exception
                    Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
                    Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
                    ModLogErrorInsert(GlobalMemberId, 0, "CreateAssetInBatchOf200 Failed For BatchID : " & I & " Error : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
                End Try
            Next

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "CreateAssetInBatchOf200 : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return -1
            Exit Function
        End Try
    End Function
    Function UpdateAssetInBatchOf200() As Boolean
        Try
            Dim foundRows() As DataRow
            Dim I As Integer
            Dim BatchTotalRec As Long
            Dim UpdateQueryXml As String
            Dim NoOfProcess As Integer

            FileIO.LogNotify("UpdateAssetInBatchOf200", FileIO.NotifyType.INFO, "Total Batch :", UpdateBatch)
            If UpdateBatch > 0 Then
                For I = 1 To UpdateBatch
                    NoOfProcess = 0
Reprocess:
                    Try
                        foundRows = AssetTable.Select("IUFlag = 'U' And Batch=" & I & "", "DeviceTypeID")
                        UpdateQueryXml = ""
                        If foundRows.Length > 0 Then
                            FileIO.LogNotify("UpdateAssetInBatchOf200", FileIO.NotifyType.INFO, "Process Start For Batch : ", I, foundRows.Length)
                            BatchTotalRec = foundRows.Length - 1
                            For Each row As DataRow In foundRows

                                If UpdateQueryXml = "" Then
                                    UpdateQueryXml = "<queryxml><entity>InstalledProduct</entity><query><condition><field>ID<expression op=""equals"">" & row("InstalledProductID") & "</expression></field></condition>"
                                ElseIf UpdateQueryXml <> "" Then
                                    UpdateQueryXml = UpdateQueryXml & "<condition operator=""OR""><field>ID<expression op=""equals"">" & row("InstalledProductID") & "</expression></field></condition>"
                                End If
                            Next
                            UpdateQueryXml = UpdateQueryXml & "</query></queryxml>"
                            If GoforBatchUpdate(UpdateQueryXml, BatchTotalRec, I) = False Then
                                If NoOfProcess = 0 Then
                                    NoOfProcess = NoOfProcess + 1
                                    GoTo Reprocess
                                End If
                            End If
                        Else
                            FileIO.LogNotify("UpdateAssetInBatchOf200", FileIO.NotifyType.INFO, "No Data Found For Batch 200 Updation")
                        End If
                    Catch ex As Exception
                        Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
                        Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
                        ModLogErrorInsert(GlobalMemberId, 0, "UpdateAssetInBatchOf200 Failed For BatchID : " & I & " Error : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
                    End Try
                Next
            Else
                FileIO.LogNotify("UpdateAssetInBatchOf200", FileIO.NotifyType.INFO, "No Record Found For Update")
            End If

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "UpdateAssetInBatchOf200 : : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
            Exit Function
        End Try
    End Function
    Function GoforBatchUpdate(ByVal InUpdateQueryXml, ByVal NoofRecords, ByVal InBatch) As Boolean
        Try
            Dim sResponse As ATWSResponse
            Dim InstalledProductToUpdate As InstalledProduct
            Dim FoundSiteID() As DataRow
            Dim foundRows() As DataRow
            Dim FoundProdRows() As DataRow

            Dim r1 As autotaskwebservices.ATWSResponse
            Dim FoundFilter() As DataRow
            Dim AR As Long
            Dim MBatchRowId As Long

            r1 = myService.query(InUpdateQueryXml.ToString)
            If r1.EntityResults.Length > 0 Then
                Dim InstalledArray(r1.EntityResults.Length - 1) As InstalledProduct

                For Each field As autotaskwebservices.InstalledProduct In r1.EntityResults
                    InstalledProductToUpdate = New InstalledProduct
                    Dim udf As New UserDefinedField
                    Dim udf1 As New UserDefinedField
                    Dim udf2 As New UserDefinedField
                    Dim udf3 As New UserDefinedField

                    Dim udf4 As New UserDefinedField
                    Dim udf5 As New UserDefinedField
                    Dim udf6 As New UserDefinedField
                    Dim udf7 As New UserDefinedField
                    Dim udf8 As New UserDefinedField
                    Dim udf9 As New UserDefinedField
                    Dim udf10 As New UserDefinedField
                    Dim udf11 As New UserDefinedField
                    Dim udf12 As New UserDefinedField
                    Dim udf13 As New UserDefinedField
                    Dim udf14 As New UserDefinedField
                    Dim udf15 As New UserDefinedField
                    Dim udf16 As New UserDefinedField
                    Dim udf17 As New UserDefinedField
                    Dim udf18 As New UserDefinedField
                    Dim udf19 As New UserDefinedField
                    Dim udf20 As New UserDefinedField
                    Dim udf21 As New UserDefinedField
                    Dim udf22 As New UserDefinedField
                    Dim udf23 As New UserDefinedField

                    Dim udf24 As New UserDefinedField
                    Dim udf25 As New UserDefinedField
                    Dim udf26 As New UserDefinedField
                    Dim udf27 As New UserDefinedField
                    Dim udf28 As New UserDefinedField
                    Dim udf29 As New UserDefinedField
                    Dim udf30 As New UserDefinedField
                    Dim udf31 As New UserDefinedField
                    Dim udf32 As New UserDefinedField
                    Dim udf33 As New UserDefinedField
                    Dim udf34 As New UserDefinedField
                    Dim udf35 As New UserDefinedField
                    Dim udf36 As New UserDefinedField

                    InstalledProductToUpdate = field

                    FoundFilter = AssetTable.Select("InstalledProductID='" & field.id & "'")
                    If FoundFilter.Length > 0 Then

                        For Each row As DataRow In FoundFilter
                            InstalledProductToUpdate.AccountID = row("ATAccountID")
                            InstalledProductToUpdate.Active = True
                            InstalledProductToUpdate.CreateDate = Now
                            InstalledProductToUpdate.id = row("InstalledProductID")
                            InstalledProductToUpdate.InstallDate = row("InstallDate")

                            FoundProdRows = ProdTable.Select("ProductName='" & UCase(row("ProductName")) & "' and ProductCategoryID ='" & row("ProductCategoryID") & "' And ProductAllocationCodeID ='" & ATProductAllocationCodeID & "'")

                            If FoundProdRows.Length > 0 Then
                                For Each Prow As DataRow In FoundProdRows
                                    InstalledProductToUpdate.ProductID = Prow("ProductID")
                                Next
                            Else
                                FileIO.LogNotify("GoforBatchUpdate", FileIO.NotifyType.INFO, row("DeviceID") & " Product ID Not found For ProductName,CategoryID ", UCase(row("ProductName")), row("ProductCategoryID"))
                            End If

                            InstalledProductToUpdate.SerialNumber = row("SerialNumber")

                            InstalledProductToUpdate.ReferenceTitle = row("ReferenceTitle")

                            If IsDate(row("WarrantyExpirationDate")) = True Then
                                InstalledProductToUpdate.WarrantyExpirationDate = row("WarrantyExpirationDate")
                            End If

                            Dim MNotes As String
                            MNotes = ""
                            MBatchRowId = MBatchRowId + 1

                            row("BatchRowID") = MBatchRowId

                            If row("DeviceTypeID") = 1 Then

                                If Trim(row("OSType")) <> "" Then
                                    MNotes = "OSType: " & row("OSType") & " " & vbCrLf
                                End If

                                If Trim(row("OSName")) <> "" Then
                                    MNotes = "OSDetail: " & row("OSName") & " " & vbCrLf
                                End If

                                If Trim(row("CPUCount")) <> "" Then
                                    MNotes = MNotes & "CPU Count: " & row("CPUCount") & " " & vbCrLf
                                End If

                                If Trim(row("CPUDetail")) <> "" Then
                                    MNotes = MNotes & row("CPUDetail") & " " & vbCrLf
                                End If

                                If Trim(row("MemorySize")) <> "" Then
                                    MNotes = MNotes & "Memory Size: " & row("MemorySize") & " " & vbCrLf
                                End If

                                If Trim(row("Manufacturer")) <> "" Then
                                    MNotes = MNotes & "Manufacturer: " & row("Manufacturer") & " " & vbCrLf
                                End If

                                If Trim(row("HardDiskInfo")) <> "" Then
                                    MNotes = MNotes & row("HardDiskInfo") & " " & vbCrLf
                                End If

                                If Trim(row("productKeyInfo")) <> "" Then
                                    MNotes = MNotes & row("productKeyInfo") & " " & vbCrLf
                                End If

                                If MNotes <> "" Then
                                    InstalledProductToUpdate.Notes = MNotes
                                End If

                            End If

                            InstalledProductToUpdate.Type = row("ProdConfigItemTypeID")

                            udf.Name = "Device ID"
                            udf.Value = row("DeviceID")

                            udf1.Name = "IP Address"
                            udf1.Value = row("IPAddress")

                            udf2.Name = "LastLoginBy"
                            udf2.Value = row("LastLoginBy")

                            udf3.Name = "Nos Of Security Patch Missing"
                            udf3.Value = row("NOOfSecurityPatch")

                            ''  from here Start Mobile Data UDF value

                            udf4.Name = "Apple Serial Number"
                            udf4.Value = row("AppleSerialNumber") & ""

                            udf5.Name = "Compliance State"
                            udf5.Value = row("ComplianceState") & ""

                            udf6.Name = "Current Carrier"
                            udf6.Value = row("CurrCarrier") & ""

                            udf7.Name = "Data Roaming?"
                            udf7.Value = row("DataRoaming") & ""

                            udf8.Name = "Device Type"
                            udf8.Value = row("DeviceType") & ""

                            udf9.Name = "Free Storage (GB)"
                            udf9.Value = row("FreeIntStorageInGB") & ""

                            udf10.Name = "Hardware Encryption"
                            udf10.Value = row("HardwareEncryption") & ""

                            udf11.Name = "Home Carrier"
                            udf11.Value = row("HomeCarrier") & ""

                            udf12.Name = "Jailbroken or Rooted?"
                            udf12.Value = row("DeviceJailBroken") & ""

                            udf13.Name = "Last Reported"
                            If IsDBNull(row("LastReported")) = False Then
                                udf13.Value = row("LastReported")
                            End If

                            udf14.Name = "Management Status"
                            udf14.Value = row("Maas360ManagedStatus") & ""

                            udf15.Name = "Manufacturer"
                            If row("DeviceTypeId") = 2 Then
                                udf15.Value = row("Manufacturer") & ""
                            Else
                                udf15.Value = ""
                            End If

                            udf16.Name = "Modern Firmware Version"
                            udf16.Value = row("ModemFirmwareVersion") & ""

                            udf17.Name = "OS Name"

                            If row("DeviceTypeId") = 2 Then
                                udf17.Value = row("OSName") & ""
                            Else
                                udf17.Value = ""
                            End If

                            udf18.Name = "OS Version"
                            udf18.Value = row("OSVersion") & ""

                            udf19.Name = "Ownership"
                            udf19.Value = row("Ownership") & ""

                            udf20.Name = "Platform Name"
                            udf20.Value = row("PlatformName") & ""

                            udf21.Name = "Policy"
                            udf21.Value = row("MDMPolicy") & ""

                            udf22.Name = "Total Storage (GB)"
                            udf22.Value = row("TotIntStorageInGB") & ""

                            udf23.Name = "WiFi Mac Address"
                            udf23.Value = row("WiFiMacAddress") & ""


                            ' For Vault Data

                            udf24.Name = "Order Number"
                            udf24.Value = row("OrderNumber") & ""

                            udf25.Name = "Provision ID"
                            udf25.Value = row("ProvisionID") & ""

                            udf26.Name = "Appliance Serial Number"
                            udf26.Value = row("ApplianceSerialNumber") & ""

                            udf27.Name = "Appliance Provision Date"
                            If IsDate(row("ApplianceProvisionDate")) = True Then
                                udf27.Value = row("ApplianceProvisionDate")
                            End If

                            udf28.Name = "Operating System"

                            If row("DeviceTypeID") = 3 Then
                                udf28.Value = row("OSName") & ""
                            Else
                                udf28.Value = ""
                            End If

                            udf29.Name = "Offsite Storage?"
                            udf29.Value = row("OffsiteStorage") & ""

                            udf30.Name = "Vault Agent ID"
                            udf30.Value = row("VaultAgentID") & ""

                            udf31.Name = "Host Name"
                            udf31.Value = row("HostName") & ""

                            udf32.Name = "Volume Name"
                            udf32.Value = row("VolumeName") & ""

                            udf33.Name = "Last Recovery At Appliance"
                            If IsDate(row("LastRecoveryAtAppliance")) = True Then
                                udf33.Value = row("LastRecoveryAtAppliance")
                            End If

                            udf34.Name = "Total Size At Appliance (GB)"
                            udf34.Value = row("TotalSizeAtAppliance") & ""

                            udf35.Name = "Last Offsite Recovery"
                            If IsDate(row("LastOffsiteRecovery")) = True Then
                                udf35.Value = row("LastOffsiteRecovery")
                            End If

                            udf36.Name = "Total Size Offsite (GB)"
                            udf36.Value = row("TotalSizeOffsite") & ""

                            InstalledProductToUpdate.UserDefinedFields = New UserDefinedField() {udf, udf1, udf2, udf3, udf4, udf5, udf6, udf7, udf8, udf9, udf10, udf11, udf12, udf13, udf14, udf15, udf16, udf17, udf18, udf19, udf20, udf21, udf22, udf23, udf24, udf25, udf26, udf27, udf28, udf29, udf30, udf31, udf32, udf33, udf34, udf35, udf36}

                        Next

                    End If
                    InstalledArray(AR) = InstalledProductToUpdate
                    AR = AR + 1
                Next

                Dim entityArray() As Entity = CType(InstalledArray, Entity())
                sResponse = myService.update(InstalledArray)

                If sResponse.ReturnCode = 1 Then
                    For Each field As autotaskwebservices.InstalledProduct In sResponse.EntityResults
                        Dim UDFLength As Integer
                        Dim MUdfDeviceID As String
                        MUdfDeviceID = ""
                        UDFLength = field.UserDefinedFields.Length

                        If UDFLength > 0 Then
                            Dim U As Integer
                            For U = 0 To (UDFLength - 1)
                                If field.UserDefinedFields(U).Name = "Device ID" Then
                                    MUdfDeviceID = field.UserDefinedFields(U).Value
                                    Exit For
                                End If
                            Next
                        End If
                        FileIO.LogNotify("GoforBatchUpdate", FileIO.NotifyType.INFO, "Account For Batch -> : " & InBatch & " DeviceID,AccountID: ", MUdfDeviceID, field.id)
                        FoundSiteID = AssetTable.Select("DeviceID='" & MUdfDeviceID & "'")
                        If FoundSiteID.Length > 0 Then
                            For Each row As DataRow In FoundSiteID

                                If row("DeviceTypeID") = "1" Then
                                    TotalAResCntSync = TotalAResCntSync + 1
                                End If

                                If row("DeviceTypeID") = "2" Then
                                    TotalMResCntSync = TotalMResCntSync + 1
                                End If

                                If row("DeviceTypeID") = "3" Then
                                    TotalVResCntSync = TotalVResCntSync + 1
                                End If

                                ATDeviceSyncStatusInsertUpdateDetails(GlobalMemberId, row("SiteId"), MUdfDeviceID, "Success", "", field.id, row("IsMapped"))

                            Next
                        End If
                    Next
                    Return True
                Else
                    Dim StrErrorMessage
                    Dim Maxlength As Long
                    Dim FindPosition As Long
                    Dim MRecordNumber As Long
                    Dim StoreError As String
                    StoreError = ""
                    For Each ent As autotaskwebservices.ATWSError In sResponse.Errors
                        StoreError = StoreError & " " & ent.Message
                        If InStr(ent.Message, "record number") > 0 Then
                            Maxlength = Len(ent.Message)
                            FindPosition = InStr(ent.Message, "[")
                            If FindPosition > 0 Then
                                StrErrorMessage = Mid(ent.Message, FindPosition, Maxlength - FindPosition)
                                StrErrorMessage = Replace(Replace(StrErrorMessage, "[", ""), "]", "")
                                If IsNumeric(StrErrorMessage) = True Then
                                    MRecordNumber = CInt(StrErrorMessage)
                                    foundRows = AssetTable.Select("IUFlag = 'U' And Batch=" & InBatch & " And BatchRowID=" & MRecordNumber & "")
                                    If foundRows.Length > 0 Then
                                        For Each row As DataRow In foundRows
                                            row("IUFlag") = ""
                                            ATDeviceSyncStatusInsertUpdateDetails(GlobalMemberId, row("SiteId"), row("DeviceID"), "Failed", StoreError, "", row("IsMapped"))
                                            StoreError = ""
                                        Next
                                    End If
                                End If
                            End If
                        End If
                        ModLogErrorInsert(GlobalMemberId, 0, "GoforBatchUpdate ATWSError Batch -> : " & InBatch & "  " & ent.Message)
                    Next
                    Return False
                End If
            End If

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "GoforBatchUpdate Batch -> : " & InBatch & " " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
            Exit Function
        End Try

    End Function
    Function ATDeviceSyncStatusInsertUpdate(ByVal InMemberId, ByVal InDeviceType, ByVal InLastSyncDt, ByVal InFullSyncDt, ByVal InTotalResCntToSync, ByVal InTotalResCntSync)
        Try
            Dim ObjCmdInsert As New SqlCommand
            With ObjCmdInsert
                .Connection = ConnITSPSAAutoTaskAPI

                .CommandTimeout = strCmdTO
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_ATDeviceSyncStatus_IU"
                .Parameters.Add("@InmemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InmemberId").Value = InMemberId
                .Parameters.Add("@InDeviceType", System.Data.SqlDbType.Int, 4)
                .Parameters("@InDeviceType").Value = InDeviceType
                .Parameters.Add("@InLastSyncDt", System.Data.SqlDbType.DateTime, 8)

                If IsDate(InLastSyncDt) = True Then
                    .Parameters("@InLastSyncDt").Value = IIf(IsDBNull(InLastSyncDt), DBNull.Value, InLastSyncDt)
                Else
                    .Parameters("@InLastSyncDt").Value = DBNull.Value
                End If

                .Parameters.Add("@InFullSyncDt", System.Data.SqlDbType.DateTime, 8)

                If IsDate(InFullSyncDt) = True Then
                    .Parameters("@InFullSyncDt").Value = IIf(IsDBNull(InFullSyncDt), DBNull.Value, InFullSyncDt)
                Else
                    .Parameters("@InFullSyncDt").Value = DBNull.Value
                End If

                .Parameters.Add("@InResCntTotalToSync", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InResCntTotalToSync").Value = InTotalResCntToSync
                .Parameters.Add("@InResCntSync", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InResCntSync").Value = InTotalResCntSync
                .Parameters.Add("@InResCntNotSync", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InResCntNotSync").Value = InTotalResCntToSync - InTotalResCntSync
                .Parameters.Add("@InSyncFlag", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InSyncFlag").Value = SyncFlag
                .Parameters.Add("@OutStatus", System.Data.SqlDbType.Int, 4)
                .Parameters("@OutStatus").Direction = ParameterDirection.Output
                .ExecuteNonQuery()
                If .Parameters("@OutStatus").Value.ToString() = "1" Then

                Else
                    FileIO.LogNotify("ATDeviceSyncStatusInsertUpdate", FileIO.NotifyType.ERR, "For Device Type : " & SyncFlag, " USP_ATDeviceSyncStatus_IU Insert / Update Failed")
                End If
            End With
            ObjCmdInsert.Dispose()
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            FileIO.LogNotify("ATDeviceSyncStatusInsertUpdate", FileIO.NotifyType.ERR, sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function ATDeviceSyncStatusInsertUpdateDetails(ByVal InMemberId, ByVal InSiteId, ByVal InRegId, ByVal InSyncStatus, ByVal InSyncErrordesc, ByVal InAtInstalledProductID, ByVal InIsmapped)
        Try
            Dim ObjCmdInsert As New SqlCommand
            With ObjCmdInsert
                .Connection = ConnITSPSAAutoTaskAPI

                .CommandTimeout = strCmdTO
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_ATDeviceSyncDetails_IU"
                .Parameters.Add("@InmemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InmemberId").Value = InMemberId
                .Parameters.Add("@InSiteId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InSiteId").Value = InSiteId
                .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InRegId").Value = InRegId
                .Parameters.Add("@InDeviceStatus", System.Data.SqlDbType.VarChar, 20)
                .Parameters("@InDeviceStatus").Value = "Active"
                .Parameters.Add("@InSyncStatus", System.Data.SqlDbType.VarChar, 20)
                .Parameters("@InSyncStatus").Value = InSyncStatus
                .Parameters.Add("@InSyncErrordesc", System.Data.SqlDbType.VarChar, 8000)
                .Parameters("@InSyncErrordesc").Value = Mid(InSyncErrordesc, 1, 8000)
                .Parameters.Add("@OutStatus", System.Data.SqlDbType.Int, 4)
                .Parameters("@OutStatus").Direction = ParameterDirection.Output
                .Parameters.Add("@IsMapped", System.Data.SqlDbType.Bit, 1)
                .Parameters("@IsMapped").Value = InIsmapped
                .Parameters.Add("@InAtInstalledProductID", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InAtInstalledProductID").Value = IIf(InAtInstalledProductID.ToString = "", DBNull.Value, InAtInstalledProductID)
                .ExecuteNonQuery()
                If .Parameters("@OutStatus").Value.ToString() = "1" Then

                Else
                    FileIO.LogNotify("ATDeviceSyncStatusInsertUpdateDetails", FileIO.NotifyType.ERR, "For RegID : " & InRegId, " USP_ATDeviceSyncDetails_IU Insert / Update Failed")
                End If
            End With
            ObjCmdInsert.Dispose()
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            FileIO.LogNotify("ATDeviceSyncStatusInsertUpdateDetails", FileIO.NotifyType.ERR, sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            ModLogErrorInsert(GlobalMemberId, InRegId, "ATDeviceSyncStatusInsertUpdateDetails : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    
    Function FetchAllocationValues(ByVal InMaterialCostCode)
        Try
            Dim boolQueryFinished = False
            Dim strCurrentID As String = "0"
            Dim strQuery As String = "<queryxml><entity>AllocationCode</entity> " & _
                                         " <query>" & _
                                         "<condition><field>UseType<expression op=""equals"">4</expression></field></condition>" & _
                                         "<condition><field>Name<expression op=""equals"">" & InMaterialCostCode & " </expression></field></condition>" & _
                                         "</query>" & _
                                         "</queryxml>"
            Dim r As autotaskwebservices.ATWSResponse
            r = myService.query(strQuery)
            If r.EntityResults.Length > 0 Then
                For Each ent As autotaskwebservices.Entity In r.EntityResults ' execute some code on the current account 
                    ATProductAllocationCodeID = ent.id
                Next
            ElseIf r.EntityResults.Length = 0 Then
                ATProductAllocationCodeID = ""
            Else
                For Each ent As autotaskwebservices.ATWSError In r.Errors
                    FileIO.LogNotify("FetchAllocationValues", FileIO.NotifyType.ERR, "ATWSError :" & ent.Message)
                Next
                ATProductAllocationCodeID = ""
            End If
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "FetchAllocationValues :  " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            ATProductAllocationCodeID = ""
        End Try
    End Function

    Function FetchAccountIdBasedOnContinuumSiteId(ByVal InMemberId As Long, ByVal InSiteId As Long) As Long
        Try
            Dim ObjCmdGetAccountId As New SqlCommand
            Dim ObjRsAccountID As SqlDataReader

            With ObjCmdGetAccountId
                ObjCmdGetAccountId.Connection = ConnITSPSAAutoTaskAPI
                ObjCmdGetAccountId.CommandTimeout = strCmdTO
                ObjCmdGetAccountId.CommandType = CommandType.StoredProcedure
                ObjCmdGetAccountId.CommandText = "USP_AT_Get_SiteAcctId"
                .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberId").Value = InMemberId
                .Parameters.Add("@InSiteId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InSiteId").Value = InSiteId
                ObjRsAccountID = ObjCmdGetAccountId.ExecuteReader
                ObjCmdGetAccountId.Dispose()
            End With

            If ObjRsAccountID.HasRows Then
                While (ObjRsAccountID.Read())
                    Return ObjRsAccountID.Item("ATAcctId")
                End While
            Else
                FileIO.LogNotify("FetchAccountIdBasedOnContinuumSiteId", FileIO.NotifyType.INFO, InSiteId & " Not found In AT_AcctIdSitewise / Not Sync Yet")
                Return -2
            End If
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "FetchAccountIdBasedOnContinuumSiteId : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return -1
        End Try
    End Function
    Function FetchPicklistValueIDForProdCategory() As Boolean
        Try
            Dim RSConfgItem As SqlDataReader
            Dim CmdConfgType As New SqlCommand
            Dim foundRows() As DataRow
            Dim CountData As Integer

            With CmdConfgType
                CmdConfgType.Connection = ConnITSPSAAutoTaskAPI
                CmdConfgType.CommandTimeout = strCmdTO
                CmdConfgType.CommandType = CommandType.StoredProcedure
                CmdConfgType.CommandText = "USP_AT_Get_MemberDetails"
                CmdConfgType.Parameters.AddWithValue("@InType", 4)
                RSConfgItem = CmdConfgType.ExecuteReader
                CmdConfgType.Dispose()
            End With

            While (RSConfgItem.Read())
                ProdCatTable.Rows.Add(0, UCase(RSConfgItem("ProdCatName")))
            End While

            Dim PArray() As String
            Dim FieldInfo() As autotaskwebservices.Field
            FieldInfo = myService.GetFieldInfo("Product")
            For Each fld As autotaskwebservices.Field In FieldInfo
                If fld.IsPickList = True Then
                    If UCase(fld.Name) = UCase("ProductCategory") Then
                        For Each FieldItem In fld.PicklistValues
                            PArray = Split(FieldItem.Label, ">")
                            For i = 0 To UBound(PArray)
                                foundRows = ProdCatTable.Select("ProductCategoryName='" & Trim(UCase(PArray(i))) & "'")
                                If foundRows.Length > 0 Then
                                    For Each row As DataRow In foundRows
                                        row("ProductCategoryID") = FieldItem.Value
                                        CountData = CountData + 1
                                    Next
                                End If
                            Next
                        Next
                    End If
                End If
            Next
            ProdCatTable.AcceptChanges()
            If ProdCatTable.Rows.Count = CountData Then
                FileIO.LogNotify("FetchPicklistValueIDForProdCategory", FileIO.NotifyType.INFO, "Success Total Count :", CountData)
                Return True
            Else
                ModLogErrorInsert(GlobalMemberId, 0, "FetchPicklistValueIDForProdCategory : ATProductCatCount " & ProdCatTable.Rows.Count & " ContinummProductCatCount " & CountData)
            End If

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "FetchPicklistValueIDForProdCategory : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function FetchPicklistValueIDForProdConfig() As Boolean
        Try
            Dim RSConfgItem As SqlDataReader
            Dim CmdConfgType As New SqlCommand
            Dim foundRows() As DataRow
            Dim CountData As Integer

            With CmdConfgType
                CmdConfgType.Connection = ConnITSPSAAutoTaskAPI
                CmdConfgType.CommandTimeout = strCmdTO
                CmdConfgType.CommandType = CommandType.StoredProcedure
                CmdConfgType.CommandText = "USP_AT_Get_MemberDetails"
                CmdConfgType.Parameters.AddWithValue("@InType", 3)
                RSConfgItem = CmdConfgType.ExecuteReader
                CmdConfgType.Dispose()
            End With

            While (RSConfgItem.Read())
                ProdConfigTable.Rows.Add(0, UCase(RSConfgItem("ConfigItemType")))
            End While

            Dim FieldInfo() As autotaskwebservices.Field
            FieldInfo = myService.GetFieldInfo("InstalledProduct")
            For Each fld As autotaskwebservices.Field In FieldInfo
                If fld.IsPickList = True Then
                    If UCase(fld.Name) = UCase("Type") Then
                        For Each FieldItem In fld.PicklistValues
                            foundRows = ProdConfigTable.Select("ProdConfigItemTypeName='" & Trim(UCase(FieldItem.Label)) & "'")
                            If foundRows.Length > 0 Then
                                For Each row As DataRow In foundRows
                                    row("ProdConfigItemTypeID") = FieldItem.Value
                                    CountData = CountData + 1
                                Next
                            End If
                        Next
                    End If
                End If
            Next
            ProdConfigTable.AcceptChanges()
            If ProdConfigTable.Rows.Count = CountData Then
                FileIO.LogNotify("FetchPicklistValueIDForProdConfig", FileIO.NotifyType.INFO, "Success Total Count :", CountData)
                Return True
            Else
                ModLogErrorInsert(GlobalMemberId, 0, "FetchPicklistValueIDForProdConfig : ATProductConfigCount " & ProdConfigTable.Rows.Count & " ContinummProductConfigCount " & CountData)
            End If

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "FetchPicklistValueIDForProdConfig : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function FetchProductDetails() As Boolean
        Try
            Dim boolQueryFinished = False
            Dim strCurrentID As String = "-1"
            Dim foundRows() As DataRow

            While Not (boolQueryFinished)

                Dim strQuery As String = "<queryxml><entity>Product</entity> " & _
                             " <query>" & _
                             "<condition><field>id<expression op=""greaterthan"">" & strCurrentID & "</expression></field></condition>" & _
                             "<condition><field>ProductAllocationCodeID<expression op=""equals"">" & ATProductAllocationCodeID & " </expression></field></condition>" & _
                             "</query>" & _
                             "</queryxml>"

                Dim r As autotaskwebservices.ATWSResponse
                r = myService.query(strQuery)
                If r.EntityResults.Length > 0 Then
                    FileIO.LogNotify("FetchProductDetails", FileIO.NotifyType.INFO, "Total Record Fetch From API ", r.EntityResults.Length)
                    For Each ent As autotaskwebservices.Entity In r.EntityResults
                        strCurrentID = ent.id
                        foundRows = ProdCatTable.Select("ProductCategoryID='" & CType(ent, autotaskwebservices.Product).ProductCategory & "'")
                        If foundRows.Length > 0 And Trim(CType(ent, autotaskwebservices.Product).ProductAllocationCodeID) = ATProductAllocationCodeID And CType(ent, autotaskwebservices.Product).Active = True Then
                            ProdTable.Rows.Add(GlobalMemberId, CType(ent, autotaskwebservices.Product).id, UCase(Trim(CType(ent, autotaskwebservices.Product).Name)), Trim(CType(ent, autotaskwebservices.Product).ProductAllocationCodeID), Trim(CType(ent, autotaskwebservices.Product).ProductCategory), "U", 0)
                        End If
                    Next
                    ProdTable.AcceptChanges()

                ElseIf r.EntityResults.Length = 0 Then
                    boolQueryFinished = True
                    Return True
                Else
                    For Each ent As autotaskwebservices.ATWSError In r.Errors
                        FileIO.LogNotify("FetchProductDetails", FileIO.NotifyType.ERR, "ATWSError :" & ent.Message)
                    Next
                    boolQueryFinished = True
                    Return False
                End If
            End While

            Return True

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "FetchProductDetails : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function FetchAssetDetailsFromAutotask() As Boolean
        Try
            Dim boolQueryFinished = False
            Dim strCurrentID As String = "-1"
            Dim UDFLength As Long
            Dim MUdfContinuumDeviceId
            Dim MUdfDeviceID
            MGetUDFINFO = 0

            While Not (boolQueryFinished)
                Dim strQuery As String = "<queryxml><entity>InstalledProduct</entity>" & _
                                         "<query><field>id<expression op=""greaterthan"">" & strCurrentID & "</expression></field></query></queryxml>"
                Dim r As autotaskwebservices.ATWSResponse
                r = myService.query(strQuery)
                If r.EntityResults.Length > 0 Then

                    FileIO.LogNotify("FetchAssetDetailsFromAutotask", FileIO.NotifyType.INFO, "Total Record Fetch From API ", r.EntityResults.Length)

                    For Each ent As autotaskwebservices.Entity In r.EntityResults

                        MUdfContinuumDeviceId = 0
                        MUdfDeviceID = 0

                        UDFLength = CType(ent, autotaskwebservices.InstalledProduct).UserDefinedFields.Length
                        If UDFLength > 0 Then
                            Dim i As Integer
                            For i = 0 To (UDFLength - 1)

                                If CType(ent, autotaskwebservices.InstalledProduct).UserDefinedFields(i).Name = "Continuum Device ID" Then
                                    MUdfContinuumDeviceId = Trim(CType(ent, autotaskwebservices.InstalledProduct).UserDefinedFields(i).Value)
                                    MGetUDFINFO = 1
                                End If

                                If CType(ent, autotaskwebservices.InstalledProduct).UserDefinedFields(i).Name = "Device ID" Then
                                    MUdfDeviceID = Trim(CType(ent, autotaskwebservices.InstalledProduct).UserDefinedFields(i).Value)
                                End If
                            Next
                        End If

                        '' For Asset Releated   Total 60 Column   On Colum 16,19,40,57 & 59 Is a datecolumn

                        If Val(MUdfContinuumDeviceId) > 0 Or Val(MUdfDeviceID) > 0 Then
                            AssetTable.Rows.Add(0, "", "", "", "", (CType(ent, autotaskwebservices.InstalledProduct).id), MUdfContinuumDeviceId, MUdfDeviceID, "", 0, 0, 0, "", (CType(ent, autotaskwebservices.InstalledProduct).AccountID), "", DBNull.Value, "", "", DBNull.Value, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", DBNull.Value, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", DBNull.Value, "", DBNull.Value, "")
                        End If

                        strCurrentID = ent.id
                        AssetTable.AcceptChanges()
                    Next
                ElseIf r.EntityResults.Length = 0 Then
                    boolQueryFinished = True
                    Return True
                Else
                    For Each ent As autotaskwebservices.ATWSError In r.Errors
                        FileIO.LogNotify("FetchAssetDetailsFromAutotask", FileIO.NotifyType.ERR, "ATWSError :" & ent.Message)
                    Next
                    boolQueryFinished = True
                    Return False
                End If
            End While
            Return True
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "FetchAssetDetailsFromAutotask : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function GetIsMappedDeviceId(ByVal InSiteID As Long, ByVal INRegID As Long) As String
        Try
            GetIsMappedDeviceId = ""

            Dim CmdCheckSiteMapped As New SqlCommand
            With CmdCheckSiteMapped
                .Connection = ConnITSPSAAutoTaskAPI
                .CommandTimeout = strCmdTO
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_AT_CheckDeviceMapped_Status"
                .Parameters.Add("@InSiteId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InSiteId").Value = InSiteID
                .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InRegId").Value = INRegID
                .Parameters.Add("@OutStatus", System.Data.SqlDbType.Int, 4)
                .Parameters("@OutStatus").Direction = ParameterDirection.Output
                .ExecuteNonQuery()
                If .Parameters("@OutStatus").Value.ToString() = "1" Then
                    GetIsMappedDeviceId = "Y"
                Else
                    GetIsMappedDeviceId = "N"
                    FileIO.LogNotify("GetIsMappedDeviceId", FileIO.NotifyType.INFO, "Is Mapped Not Found ...")
                End If
                CmdCheckSiteMapped.Dispose()
            End With
        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "GetIsMappedDeviceId : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return ""
        End Try
    End Function
    Function GetFailedRegID(ByVal InMemberId As Long, ByVal InRegId As Long) As Long
        Try
            Dim ObjCmdDevice1 As New SqlCommand
            Dim ObjRsData As SqlDataReader

            With ObjCmdDevice1
                .Connection = ConnITSPSAAutoTaskAPI
                .CommandTimeout = strCmdTO
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_AT_Get_InstalledPrdId"
                .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberId").Value = GlobalMemberId
                .Parameters.Add("@InRegId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InRegId").Value = InRegId
                .Parameters.Add("@InMode", System.Data.SqlDbType.Int, 4)
                .Parameters("@InMode").Value = 2
                ObjRsData = .ExecuteReader
                .Dispose()
            End With

            If ObjRsData.HasRows Then
                While (ObjRsData.Read())
                    FileIO.LogNotify("GetInstalledProductID", FileIO.NotifyType.INFO, "USP_AT_Get_InstalledPrdId Data Found For RegID Mode 2", InRegId)
                    Return ObjRsData.Item("RegID")
                End While
            Else
                FileIO.LogNotify("GetInstalledProductID", FileIO.NotifyType.INFO, "USP_AT_Get_InstalledPrdId No Data Found For RegID Mode 2", InRegId)
                Return 0
            End If

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "GetInstalledProductID : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            Return -1
        End Try
    End Function
    Public Function TableCreationProcess() As Boolean
        Try

            ProdCatTable = New DataTable
            ProdTable = New DataTable
            ProdConfigTable = New DataTable
            AssetTable = New DataTable

            '' For ProductCatTable
            ProdCatTable.Columns.Add("ProductCategoryID", Type.GetType("System.String"))
            ProdCatTable.Columns.Add("ProductCategoryName", Type.GetType("System.String"))

            '' For Product ConfigType
            ProdConfigTable.Columns.Add("ProdConfigItemTypeID", Type.GetType("System.String"))
            ProdConfigTable.Columns.Add("ProdConfigItemTypeName", Type.GetType("System.String"))

            '' For Product 
            ProdTable.Columns.Add("MemberId", Type.GetType("System.String"))
            ProdTable.Columns.Add("ProductId", Type.GetType("System.String"))
            ProdTable.Columns.Add("ProductName", Type.GetType("System.String"))
            ProdTable.Columns.Add("ProductAllocationCodeID", Type.GetType("System.String"))
            ProdTable.Columns.Add("ProductCategoryID", Type.GetType("System.String"))
            ProdTable.Columns.Add("IUFlag", Type.GetType("System.String"))
            ProdTable.Columns.Add("Batch", Type.GetType("System.Int32"))


            ''For Asset Releated   Total 60 Column  On Colum 15,19,40,57 & 59 Is a datecolumn

            '' Additional field required for Checking and making a batch or Insert / Update.

            AssetTable.Columns.Add("BatchRowID", Type.GetType("System.Int32"))
            AssetTable.Columns.Add("ProductCategoryID", Type.GetType("System.String"))
            AssetTable.Columns.Add("ProdConfigItemTypeID", Type.GetType("System.String"))
            AssetTable.Columns.Add("ProductID", Type.GetType("System.String"))
            AssetTable.Columns.Add("ProductName", Type.GetType("System.String"))

            AssetTable.Columns.Add("InstalledProductID", Type.GetType("System.String"))

            AssetTable.Columns.Add("ContinuumDeviceID", Type.GetType("System.String"))
            AssetTable.Columns.Add("DeviceID", Type.GetType("System.String"))

            AssetTable.Columns.Add("IUFlag", Type.GetType("System.String"))
            AssetTable.Columns.Add("IsMapped", Type.GetType("System.Int32"))
            AssetTable.Columns.Add("IsProcessed", Type.GetType("System.Int32"))
            AssetTable.Columns.Add("Batch", Type.GetType("System.Int32"))


            AssetTable.Columns.Add("DeviceTypeID", Type.GetType("System.String"))  ' Asset / Mobile / Backup  I.e (1,2,3)
            AssetTable.Columns.Add("ATAccountID", Type.GetType("System.String"))
            AssetTable.Columns.Add("SiteID", Type.GetType("System.String"))

            AssetTable.Columns.Add("InstallDate", GetType(DateTime))      '14     'Is Common For Mobile & Vault Data Also
            AssetTable.Columns.Add("SerialNumber", Type.GetType("System.String"))
            AssetTable.Columns.Add("ReferenceTitle", Type.GetType("System.String"))
            AssetTable.Columns.Add("WarrantyExpirationDate", GetType(DateTime))   '17


            '' For Server / Desktop / VirtualHost / Linux UDF's Field
            AssetTable.Columns.Add("IPAddress", Type.GetType("System.String"))
            AssetTable.Columns.Add("LastLoginBy", Type.GetType("System.String"))
            AssetTable.Columns.Add("NOOfSecurityPatch", Type.GetType("System.String"))
            AssetTable.Columns.Add("OSType", Type.GetType("System.String"))
            AssetTable.Columns.Add("OSName", Type.GetType("System.String"))      'OsName is Common For Mobile & Vault Data Also
            AssetTable.Columns.Add("CPUCount", Type.GetType("System.String"))
            AssetTable.Columns.Add("CPUDetail", Type.GetType("System.String"))
            AssetTable.Columns.Add("HardDiskInfo", Type.GetType("System.String"))
            AssetTable.Columns.Add("MemorySize", Type.GetType("System.String"))
            AssetTable.Columns.Add("Manufacturer", Type.GetType("System.String")) 'Manufacturer is Common For Mobile
            AssetTable.Columns.Add("productKeyInfo", Type.GetType("System.String"))

            '' For Mobile  UDF's Field
            AssetTable.Columns.Add("AppleSerialNumber", Type.GetType("System.String"))
            AssetTable.Columns.Add("ComplianceState", Type.GetType("System.String"))
            AssetTable.Columns.Add("CurrCarrier", Type.GetType("System.String"))
            AssetTable.Columns.Add("DataRoaming", Type.GetType("System.String"))
            AssetTable.Columns.Add("DeviceType", Type.GetType("System.String"))
            AssetTable.Columns.Add("FreeIntStorageInGB", Type.GetType("System.String"))
            AssetTable.Columns.Add("HardwareEncryption", Type.GetType("System.String"))
            AssetTable.Columns.Add("HomeCarrier", Type.GetType("System.String"))
            AssetTable.Columns.Add("DeviceJailBroken", Type.GetType("System.String"))
            AssetTable.Columns.Add("LastReported", GetType(DateTime))     ' 38
            AssetTable.Columns.Add("Maas360ManagedStatus", Type.GetType("System.String"))
            AssetTable.Columns.Add("ModemFirmwareVersion", Type.GetType("System.String"))
            AssetTable.Columns.Add("OSVersion", Type.GetType("System.String"))
            AssetTable.Columns.Add("Ownership", Type.GetType("System.String"))
            AssetTable.Columns.Add("PlatformName", Type.GetType("System.String"))
            AssetTable.Columns.Add("MDMPolicy", Type.GetType("System.String"))
            AssetTable.Columns.Add("TotIntStorageInGB", Type.GetType("System.String"))
            AssetTable.Columns.Add("WiFiMacAddress", Type.GetType("System.String"))

            '' For Backup i.e (Vault) UDF's Field
            AssetTable.Columns.Add("OrderNumber", Type.GetType("System.String"))
            AssetTable.Columns.Add("ProvisionID", Type.GetType("System.String"))
            AssetTable.Columns.Add("ApplianceSerialNumber", Type.GetType("System.String"))
            AssetTable.Columns.Add("ApplianceProvisionDate", Type.GetType("System.String"))
            AssetTable.Columns.Add("OffsiteStorage", Type.GetType("System.String"))
            AssetTable.Columns.Add("VaultAgentID", Type.GetType("System.String"))
            AssetTable.Columns.Add("HostName", Type.GetType("System.String"))
            AssetTable.Columns.Add("VolumeName", Type.GetType("System.String"))
            AssetTable.Columns.Add("LastRecoveryAtAppliance", GetType(DateTime))  ' 55
            AssetTable.Columns.Add("TotalSizeAtAppliance", Type.GetType("System.String"))
            AssetTable.Columns.Add("LastOffsiteRecovery", GetType(DateTime))  ' 57
            AssetTable.Columns.Add("TotalSizeOffsite", Type.GetType("System.String"))


            TableCreationProcess = True

        Catch ex As Exception
            Dim st As StackTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim sf As StackFrame = st.GetFrame(st.FrameCount - 1)
            ModLogErrorInsert(GlobalMemberId, 0, "TableCreationProcess : " & sf.GetFileLineNumber() & " Message : " & ex.Message.ToString)
            TableCreationProcess = False
        End Try

    End Function

    Function SetZoneWiseAutoTaskAPI(ByVal InUserName As String, ByVal InPassword As String, ByVal InMaterialCostCode As String) As Boolean
        Try
            myService.Url = URLZone
            Dim cred As New System.Net.NetworkCredential(InUserName, InPassword)
            Dim credCache As New System.Net.CredentialCache
            credCache.Add(New Uri(myService.Url), "Basic", cred)
            myService.Credentials = credCache
            FetchAllocationValues(InMaterialCostCode)
            If ATProductAllocationCodeID <> "" Then
                Return True
            Else
                FileIO.LogNotify("SetZoneWiseAutoTaskAPI", FileIO.NotifyType.ERR, "CheckLoginPwdvalidation")
                Return False
            End If
            Return True
        Catch ex As Exception
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "SetZoneWiseAutoTaskAPI : " & ex.Message.ToString)
            Return False
        End Try
    End Function
    Function GetZoneInfoAutoTaskAPI(ByVal InUserName As String, ByVal InPassword As String) As String
        '// this call can be issued to any zone, we have chosen
        '// webservices.autotask.net (North American zone)
        Try
            myService.Url = "https://webservices.autotask.net/atservices/1.5/atws.asmx"
            Dim cred As New System.Net.NetworkCredential(InUserName, InPassword)
            Dim credCache As New System.Net.CredentialCache
            credCache.Add(New Uri(myService.Url), "Basic", cred)
            myService.Credentials = credCache

            Dim resZoneInfo As autotaskwebservices.ATWSZoneInfo
            resZoneInfo = myService.getZoneInfo(InUserName)
            If resZoneInfo.ErrorCode >= 0 Then
                'Log("ZoneInfo addr:" & resZoneInfo.URL)
                FileIO.LogNotify("GetZoneInfoAutoTaskAPI", FileIO.NotifyType.INFO, "ZoneInfo addr:" & resZoneInfo.URL)
                Return resZoneInfo.URL
            Else
                'Log("ZoneInfo error! [" & resZoneInfo.ErrorCode.ToString & "]:" & resZoneInfo.URL)
                FileIO.LogNotify("GetZoneInfoAutoTaskAPI", FileIO.NotifyType.INFO, "ZoneInfo error! [" & resZoneInfo.ErrorCode.ToString & "]:" & resZoneInfo.URL)
                Return ""
                Exit Function
            End If
        Catch ex As Exception
            'Conlog("Connection Test (ZoneInfo) Failed:" & ex.Message)
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "GetZoneInfoAutoTaskAPI : " & ex.Message.ToString)
            Return ""
            Exit Function
        End Try
    End Function
    Function getThresholdAndUsageInfo() As Boolean
        Try
            Dim SplitData() As String
            Dim ReplaceData As String
            Dim ThresholdOfExternalRequestLimit As Long
            Dim NumberOfExternalRequest As Long

            Dim response As autotaskwebservices.ATWSResponse = myService.getThresholdAndUsageInfo()
            If response.ReturnCode = 1 Then
                For Each er As autotaskwebservices.EntityReturnInfo In response.EntityReturnInfoResults

                    'MsgBox(er.Message) ' Sample Output: ' thresholdOfExternalRequest: 10000; TimeframeOfLimitation: 60; numberOfExternalRequest: 99; 

                    ReplaceData = Replace(Replace(Replace(er.Message, "thresholdOfExternalRequest:", ""), "TimeframeOfLimitation", ""), "numberOfExternalRequest", "")
                    ReplaceData = Replace(Replace(ReplaceData, " ", ""), ";", "")

                    SplitData = Split(ReplaceData, ":")

                    If UBound(SplitData) > 1 Then
                        ThresholdOfExternalRequestLimit = SplitData(0)
                        NumberOfExternalRequest = SplitData(2)
                    End If

                    FileIO.LogNotify("getThresholdAndUsageInfo", FileIO.NotifyType.INFO, er.Message)

                    If NumberOfExternalRequest < (ThresholdOfExternalRequestLimit - NumberOfExternalRequestBuffer) Then
                        getThresholdAndUsageInfo = True
                    Else
                        getThresholdAndUsageInfo = False
                    End If

                Next
            Else
                For Each ex As autotaskwebservices.ATWSError In response.Errors
                    ModLogErrorInsert(GlobalMemberId, 0, "getThresholdAndUsageInfo " & ex.Message.ToString)
                    getThresholdAndUsageInfo = False
                Next
            End If
        Catch ex As Exception
            ModLogErrorInsert(GlobalMemberId, 0, "getThresholdAndUsageInfo : " & ex.Message.ToString)
            getThresholdAndUsageInfo = False
        End Try
    End Function

    Function PrevInstance() As Boolean
        Try
            If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
        End Try
    End Function
    Function GetUDFINFO(ByVal EntityName As String, ByVal InUDFName As String) As Integer
        Try
            Dim FieldInfo() As autotaskwebservices.Field
            FieldInfo = myService.getUDFInfo(EntityName)
            For Each UDF As autotaskwebservices.Field In FieldInfo
                If UCase(Trim(UDF.Name)) = Trim(UCase(InUDFName)) Then
                    Return 1
                    Exit For
                End If
            Next
        Catch ex As Exception
            ModLogErrorInsert(GlobalMemberId, GlobalRegID, "GetUDFINFO : " & ex.Message.ToString)
            Return -1
        End Try
    End Function
    Function ModLogErrorInsert(ByVal InErrMemberId, ByVal InErrRegId, ByVal InErrDesc)
        Try
            Dim ObjCmdError As New SqlCommand

            With ObjCmdError
                ObjCmdError.Connection = ConnITSPSAAutoTaskAPI
                ObjCmdError.CommandTimeout = strCmdTO
                ObjCmdError.CommandType = CommandType.StoredProcedure
                ObjCmdError.CommandText = "USP_AT_ModuleErrorLog_Insert"
                .Parameters.Add("@InMemberId", System.Data.SqlDbType.BigInt, 8)
                .Parameters("@InMemberId").Value = InErrMemberId
                .Parameters.Add("@InModuleName", System.Data.SqlDbType.VarChar, 200)
                .Parameters("@InModuleName").Value = FileIO.ExeFileName
                .Parameters.Add("@InReg_Site_Tkt_RuleID", System.Data.SqlDbType.VarChar, 200)
                .Parameters("@InReg_Site_Tkt_RuleID").Value = InErrRegId
                .Parameters.Add("@InTransactionType", System.Data.SqlDbType.VarChar, 20)
                .Parameters("@InTransactionType").Value = DBNull.Value
                .Parameters.Add("@InErrorDesc", System.Data.SqlDbType.VarChar, Len(InErrDesc) + 1)
                .Parameters("@InErrorDesc").Value = InErrDesc
                .Parameters.Add("@OutStatus", System.Data.SqlDbType.Int).Direction = ParameterDirection.Output
                .ExecuteNonQuery()
                If .Parameters("@OutStatus").Value.ToString() <> "1" Then
                    FileIO.LogNotify("ModLogErrorInsert", FileIO.NotifyType.ERR, "USP_AT_ModuleErrorLog_Insert Failed")
                End If
            End With

        Catch ex As Exception
            FileIO.LogNotify("ModLogErrorInsert", FileIO.NotifyType.ERR, ex.Message.ToString)
            Return False
        End Try
    End Function
    Function OpenConnection() As Boolean
        Try
            If getConDtls() Then
                Dim OutErr As String
                OutErr = ""

                If ClsDBO.openConnection(ConnITSPSAAutoTaskAPI, ITSPSAAutoTaskAPIStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSPSA_AutoTaskAPI")
                    OpenConnection = False
                    Exit Function
                End If


                If ClsDBO.openConnection(ConnITSAssetDB, ITSAssetDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSAssetDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnItsupport247DB, ITSupport247dbStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSupport247db")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSPatchDB, ITSPatchDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSPatchDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSMACAssetDB, ITSMACAssetDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSMACAssetDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSVMWareAssetDB, ITSVMWareAssetDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSVMWareAssetDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSMDMgmtDB, ITSMDMgmtDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSMDMgmtDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSRMMVaultDB, ITSRMMVaultDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSRMMVaultDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSLinuxAssetDB, ITSLinuxAssetDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSLinuxAssetDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSPassVaultDB, ITSPassVaultDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSPassVaultDB")
                    OpenConnection = False
                    Exit Function
                End If

                If ClsDBO.openConnection(ConnITSSADDB, ITSSADDBStr, OutErr, strConTO) = False Then
                    FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, "DBConnection Failure for ITSSADDB")
                    OpenConnection = False
                    Exit Function
                End If

                OpenConnection = True
            Else
                OpenConnection = False
            End If
        Catch ex As Exception
            FileIO.LogNotify("OpenConnection", FileIO.NotifyType.ERR, ex.Message.ToString)
            OpenConnection = False
        End Try
    End Function
    Function getConDtls() As Boolean
        Try
            ITSPSAAutoTaskAPIStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSPSAAutoTaskAPIStr")
            ITSAssetDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSAssetDBStrTest")
            ITSupport247dbStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "Itsupport247dbStr")
            ITSPatchDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSPatchDBStr")
            ITSMACAssetDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSMACAssetDBStr")
            ITSVMWareAssetDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ItsVMWareAssetDBStr")
            ITSMDMgmtDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSMDMgmtDBStrTest")
            ITSRMMVaultDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSRMMVaultDBStrTest")
            ITSLinuxAssetDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSLinuxAssetDBStr")
            ITSPassVaultDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSPassVaultDBStr")
            ITSSADDBStr = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "ITSSADDBStr")
            NumberOfExternalRequestBuffer = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "NumberOfExternalRequestBuffer")

            strConTO = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "conntimeoutsec")
            strCmdTO = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "cmdtimeoutsec")

            ' strAssetFullSyncTimeFrom = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "AssetFullSyncTimeFrom")
            ' strAssetFullSyncTimeTo = FileIO.GetINIKeyValue("AutoTaskAPIIntegration", "AssetFullSyncTimeTo")

            If Trim$(ITSAssetDBStr) = "" Or Trim$(ITSupport247dbStr) = "" Or Trim$(ITSPatchDBStr) = "" Or Trim$(ITSMACAssetDBStr) = "" Or Trim$(ITSVMWareAssetDBStr) = "" Or Trim$(ITSMDMgmtDBStr) = "" Or Trim$(ITSRMMVaultDBStr) = "" Or Trim$(ITSPSAAutoTaskAPIStr) = "" Or Trim$(ITSLinuxAssetDBStr) = "" Or Trim$(ITSPassVaultDBStr) = "" Or Trim$(ITSSADDBStr) = "" Or Trim(NumberOfExternalRequestBuffer) = "" Then
                FileIO.LogNotify("getConDtls", FileIO.NotifyType.INFO, "DBConnection Data missing for AutoTaskAPIIntegration")
            Else
                getConDtls = True
            End If
        Catch ex As Exception
            getConDtls = False
        End Try
    End Function

    Function CloseConnection()
        On Error Resume Next
        ConnITSPSAAutoTaskAPI.Dispose()
        ConnITSPSAAutoTaskAPI = Nothing
        ConnITSAssetDB.Dispose()
        ConnITSAssetDB = Nothing
        ConnItsupport247DB.Dispose()
        ConnItsupport247DB = Nothing
        ConnITSPatchDB.Dispose()
        ConnITSPatchDB = Nothing
        ConnITSMACAssetDB.Dispose()
        ConnITSMACAssetDB = Nothing
        ConnITSVMWareAssetDB.Dispose()
        ConnITSVMWareAssetDB = Nothing
        ConnITSMDMgmtDB.Dispose()
        ConnITSMDMgmtDB = Nothing
        ConnITSRMMVaultDB.Dispose()
        ConnITSRMMVaultDB = Nothing
        ConnITSLinuxAssetDB.Dispose()
        ConnITSLinuxAssetDB = Nothing
        ConnITSPassVaultDB.Dispose()
        ConnITSPassVaultDB = Nothing
        ConnITSSADDB.Dispose()
        ConnITSSADDB = Nothing
    End Function
End Module