' Copyright 2021 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SapPPMRibbonParam
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapPPMRibbonParam getGenParametrs - " & "reading Parameter")
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

    Public Sub getParameter(ByRef pSapCon As SapCon)
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

        Dim aSAP_PPMParam As New SAP_PPMParam(pSapCon, aIntPar)

        aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
        ' Read the Items
        Try
            log.Debug("SapPPMRibbonParam.getParameter - " & "processing data - disabling events, screen update, cursor")
            Dim aItems As New TData(aIntPar)
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("DBG", "DUMPOBJNR")), 0)
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapPPMExcelAddIn.Application.EnableEvents = False
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = False
            Dim aTSAP_GetParamData As New TSAP_GetParamData(aPar, aIntPar)
            Dim aET_ZPPM_PARAM As IRfcTable = Nothing
            Dim aET_RETURN As IRfcTable = Nothing
            If aTSAP_GetParamData.fillHeader(aItems) And aTSAP_GetParamData.fillData(aItems) Then
                ' check if we should dump this document
                If aDumpObjNr <> 0 Then
                    log.Debug("SapPPMRibbonParam.getParameter - " & "dumping Object Nr " & CStr(aObjNr))
                    aTSAP_GetParamData.dumpHeader()
                    aTSAP_GetParamData.dumpData()
                End If
                log.Debug("SapPPMRibbonParam.getParameter - " & "calling SAP_PPMTimesheet.getParameter")
                aRetStr = aSAP_PPMParam.getParam(aTSAP_GetParamData, aET_ZPPM_PARAM, aET_RETURN)
                log.Debug("SapPPMRibbonParam.getParameter - " & "SAP_PPMTimesheet.get returned, aRetStr=" & aRetStr)
            End If
            ' output result
            Dim aOutputHelper As New OutputHelper
            Dim aPARwsName As String = If(aIntPar.value("WS", "PAR") <> "", aIntPar.value("WS", "PAR"), "ET_ZPPM_PARAM")
            aOutputHelper.write(aPARwsName, aLOff, aIntPar, aWB, aET_ZPPM_PARAM, pClear:=True, pClearColumn:=1)
            Dim aRETwsName As String = If(aIntPar.value("WS", "RET") <> "", aIntPar.value("WS", "RET"), "ET_RETURN")
            aOutputHelper.write(aRETwsName, aLOff, aIntPar, aWB, aET_RETURN, pClear:=True, pClearColumn:=1)
            log.Debug("SapPPMRibbonParam.getParameter - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapPPMExcelAddIn.Application.EnableEvents = True
            Globals.SapPPMExcelAddIn.Application.ScreenUpdating = True
            Globals.SapPPMExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPPMRibbonParam.getParameter failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPPMRibbonParam.getParameter - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub
End Class
