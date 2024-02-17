' Copyright 2021 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SapPPMRibbonTimesheet
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapPPMRibbonTimesheet getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP PPM Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PPM")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SapPPMExcelAddin"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP PPM Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PPM")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP PPM Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PPM")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub getTimeAttendance(ByRef pSapCon As SapCon)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aWB As Excel.Workbook

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAP_PPMTimesheet As New SAP_PPMTimesheet(pSapCon, aIntPar)

        aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
        Dim aEMAwsName As String = If(aIntPar.value("WS", "EMA") <> "", aIntPar.value("WS", "EMA"), "IT_EMAIL")
        Dim aInputHelper As New InputHelper
        ' Read the Items
        Try
            log.Debug("SapPPMRibbonTimesheet.getTimeAttendance - " & "processing data - disabling events, screen update, cursor")
            Dim aItems As New TData(aIntPar)
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("DBG", "DUMPOBJNR")), 0)
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPPMExcelAddIn.Application.EnableEvents = False
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = False
            aInputHelper.read(aEMAwsName, aLOff, "EMA", aWB, aItems)
            Dim aTSAP_GetTimeAttendanceData As New TSAP_GetTimeAttendanceData(aPar, aIntPar)
            Dim aET_TIME_ATTENDANCE As IRfcTable = Nothing
            If aTSAP_GetTimeAttendanceData.fillHeader(aItems) And aTSAP_GetTimeAttendanceData.fillData(aItems) Then
                ' check if we should dump this document
                If aDumpObjNr <> 0 Then
                    log.Debug("SapPPMRibbonTimesheet.getTimeAttendance - " & "dumping Object Nr " & CStr(aObjNr))
                    aTSAP_GetTimeAttendanceData.dumpHeader()
                    aTSAP_GetTimeAttendanceData.dumpData()
                End If
                log.Debug("SapPPMRibbonTimesheet.getTimeAttendance - " & "calling SAP_PPMTimesheet.getTimeAttendance")
                aRetStr = aSAP_PPMTimesheet.getTimeAttendance(aTSAP_GetTimeAttendanceData, aET_TIME_ATTENDANCE)
                log.Debug("SapPPMRibbonTimesheet.getTimeAttendance - " & "SAP_PPMTimesheet.get returned, aRetStr=" & aRetStr)
            End If
            ' output result
            Dim aOutputHelper As New OutputHelper
            Dim aTATwsName As String = If(aIntPar.value("WS", "TAT") <> "", aIntPar.value("WS", "TAT"), "ET_TIME_ATTENDANCE")
            aOutputHelper.write(aTATwsName, aLOff, aIntPar, aWB, aET_TIME_ATTENDANCE, pClear:=True, pClearColumn:=1)

            log.Debug("SapPPMRibbonTimesheet.getTimeAttendance - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPPMRibbonTimesheet.getTimeAttendance failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPPMRibbonTimesheet.getTimeAttendance - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub getTsIds(ByRef pSapCon As SapCon)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aWB As Excel.Workbook

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAP_PPMTimesheet As New SAP_PPMTimesheet(pSapCon, aIntPar)

        aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
        Dim aIDSwsName As String = If(aIntPar.value("WS", "IDS") <> "", aIntPar.value("WS", "IDS"), "IT_ZPPM_S_TS_IDS")
        Dim aInputHelper As New InputHelper
        ' Read the Items
        Try
            log.Debug("SapPPMRibbonTimesheet.getTsIds - " & "processing data - disabling events, screen update, cursor")
            Dim aItems As New TData(aIntPar)
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("DBG", "DUMPOBJNR")), 0)
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPPMExcelAddIn.Application.EnableEvents = False
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = False
            aInputHelper.read(aIDSwsName, aLOff, "IDS", aWB, aItems)
            Dim aTSAP_GetTsIdsData As New TSAP_GetTsIdsData(aPar, aIntPar)
            Dim aET_ZPPM_S_TS_IDS As IRfcTable = Nothing
            Dim aET_RETURN As IRfcTable = Nothing
            If aTSAP_GetTsIdsData.fillHeader(aItems) And aTSAP_GetTsIdsData.fillData(aItems) Then
                ' check if we should dump this document
                If aDumpObjNr <> 0 Then
                    log.Debug("SapPPMRibbonTimesheet.getTsIds - " & "dumping Object Nr " & CStr(aObjNr))
                    aTSAP_GetTsIdsData.dumpHeader()
                    aTSAP_GetTsIdsData.dumpData()
                End If
                log.Debug("SapPPMRibbonTimesheet.getTsIds - " & "calling SAP_PPMTimesheet.getTsIds")
                aRetStr = aSAP_PPMTimesheet.getTsIds(aTSAP_GetTsIdsData, aET_ZPPM_S_TS_IDS, aET_RETURN)
                log.Debug("SapPPMRibbonTimesheet.getTsIds - " & "SAP_PPMTimesheet.get returned, aRetStr=" & aRetStr)
            End If
            ' output result
            Dim aOutputHelper As New OutputHelper
            Dim aODSwsName As String = If(aIntPar.value("WS", "ODS") <> "", aIntPar.value("WS", "ODS"), "ET_ZPPM_S_TS_IDS")
            aOutputHelper.write(aODSwsName, aLOff, aIntPar, aWB, aET_ZPPM_S_TS_IDS, pClear:=True, pClearColumn:=1)
            Dim aRETwsName As String = If(aIntPar.value("WS", "RET") <> "", aIntPar.value("WS", "RET"), "ET_RETURN")
            aOutputHelper.write(aRETwsName, aLOff, aIntPar, aWB, aET_RETURN, pClear:=True, pClearColumn:=1)

            log.Debug("SapPPMRibbonTimesheet.getTsIds - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPPMRibbonTimesheet.getTsIds failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPPMRibbonTimesheet.getTsIds - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub postTS(ByRef pSapCon As SapCon)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aWB As Excel.Workbook

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAP_PPMTimesheet As New SAP_PPMTimesheet(pSapCon, aIntPar)

        aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
        Dim aICAwsName As String = If(aIntPar.value("WS", "ICA") <> "", aIntPar.value("WS", "ICA"), "IT_CATS")
        Dim aInputHelper As New InputHelper
        ' Read the Items
        Try
            log.Debug("SapPPMRibbonTimesheet.postTS - " & "processing data - disabling events, screen update, cursor")
            Dim aItems As New TData(aIntPar)
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("DBG", "DUMPOBJNR")), 0)
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPPMExcelAddIn.Application.EnableEvents = False
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = False
            aInputHelper.read(aICAwsName, aLOff, "ICA", aWB, aItems)
            Dim aTSAP_PostTSData As New TSAP_PostTSData(aPar, aIntPar)
            Dim aET_CATS As IRfcTable = Nothing
            Dim aET_RETURN As IRfcTable = Nothing
            If aTSAP_PostTSData.fillHeader(aItems) And aTSAP_PostTSData.fillData(aItems) Then
                ' check if we should dump this document
                If aDumpObjNr <> 0 Then
                    log.Debug("SapPPMRibbonTimesheet.postTS - " & "dumping Object Nr " & CStr(aObjNr))
                    aTSAP_PostTSData.dumpHeader()
                    aTSAP_PostTSData.dumpData()
                End If
                log.Debug("SapPPMRibbonTimesheet.postTS - " & "calling SAP_PPMTimesheet.postTS")
                aRetStr = aSAP_PPMTimesheet.postTS(aTSAP_PostTSData, aET_CATS, aET_RETURN)
                log.Debug("SapPPMRibbonTimesheet.postTS - " & "SAP_PPMTimesheet.get returned, aRetStr=" & aRetStr)
            End If
            ' output result
            Dim aOutputHelper As New OutputHelper
            Dim aECAwsName As String = If(aIntPar.value("WS", "ECA") <> "", aIntPar.value("WS", "ECA"), "ET_CATS")
            aOutputHelper.write(aECAwsName, aLOff, aIntPar, aWB, aET_CATS, pClear:=True, pClearColumn:=1)
            Dim aRETwsName As String = If(aIntPar.value("WS", "RET") <> "", aIntPar.value("WS", "RET"), "ET_RETURN")
            aOutputHelper.write(aRETwsName, aLOff, aIntPar, aWB, aET_RETURN, pClear:=True, pClearColumn:=1)

            log.Debug("SapPPMRibbonTimesheet.postTS - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPPMRibbonTimesheet.postTS failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPPMRibbonTimesheet.postTS - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

End Class
