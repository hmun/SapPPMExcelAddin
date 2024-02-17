' Copyright 2021 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SapPPMRibbonActuals
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapPPMRibbonActuals getGenParametrs - " & "reading Parameter")
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

    Public Sub getActuals(ByRef pSapCon As SapCon)
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

        Dim aSAP_PPMActuals As New SAP_PPMActuals(pSapCon, aIntPar)

        aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
        Dim aInputHelper As New InputHelper
        Dim aWBSwsName As String = If(aIntPar.value("WS", "WBS") <> "", aIntPar.value("WS", "WBS"), "IT_PSP")
        Dim aCERwsName As String = If(aIntPar.value("WS", "CER") <> "", aIntPar.value("WS", "CER"), "IT_CER")
        Dim aORDwsName As String = If(aIntPar.value("WS", "ORD") <> "", aIntPar.value("WS", "ORD"), "IT_ORD")
        ' Read the Items
        Try
            log.Debug("SapPPMRibbonActuals.exec - " & "processing data - disabling events, screen update, cursor")
            Dim aItems As New TData(aIntPar)
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("DBG", "DUMPOBJNR")), 0)
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPPMExcelAddIn.Application.EnableEvents = False
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = False
            aInputHelper.read(aWBSwsName, aLOff, "WBS", aWB, aItems)
            aInputHelper.read(aCERwsName, aLOff, "CER", aWB, aItems)
            aInputHelper.read(aORDwsName, aLOff, "ORD", aWB, aItems)
            Dim aTSAP_GetActualsData As New TSAP_GetActualsData(aPar, aIntPar)
            Dim aET_RETURN As IRfcTable = Nothing
            Dim aET_ACTUALS As IRfcTable = Nothing
            If aTSAP_GetActualsData.fillHeader(aItems) And aTSAP_GetActualsData.fillData(aItems) Then
                ' check if we should dump this document
                If aDumpObjNr <> 0 Then
                    log.Debug("SapPPMRibbonActuals.exec - " & "dumping Object Nr " & CStr(aObjNr))
                    aTSAP_GetActualsData.dumpHeader()
                    aTSAP_GetActualsData.dumpData()
                End If
                log.Debug("SapPPMRibbonActuals.exec - " & "calling SAP_PPMActuals.get")
                aRetStr = aSAP_PPMActuals.getActuals(aTSAP_GetActualsData, aET_ACTUALS, aET_RETURN)
                log.Debug("SapPPMRibbonActuals.exec - " & "SAP_PPMActuals.get returned, aRetStr=" & aRetStr)
                ' message has to be written in all lines that where processed in items
                For Each aKey In aItems.aTDataDic.Keys
                Next
            End If
            ' output result
            Dim aOutputHelper As New OutputHelper
            Dim aACTwsName As String = If(aIntPar.value("WS", "ACT") <> "", aIntPar.value("WS", "ACT"), "ET_ACTUALS")
            aOutputHelper.write(aACTwsName, aLOff, aIntPar, aWB, aET_ACTUALS, pClear:=True, pClearColumn:=12)
            Dim aRETwsName As String = If(aIntPar.value("WS", "RET") <> "", aIntPar.value("WS", "RET"), "ET_RETURN")
            aOutputHelper.write(aRETwsName, aLOff, aIntPar, aWB, aET_RETURN, pClear:=True, pClearColumn:=1)

            log.Debug("SapPPMRibbonActuals.exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPPMRibbonActuals.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPPMRibbonActuals.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

End Class
