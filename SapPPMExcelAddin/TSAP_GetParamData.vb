' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TSAP_GetParamData

    Public aHdrRec As TDataRec
    Public aData As TData
    Public aExt As TData

    Private Hd_Fields() As String = {"IV_PARAM"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private aUseAsEmpty As String = "#"
    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aPar = pPar
        aIntPar = pIntPar
        aUseAsEmpty = If(aIntPar.value("GEN", "USEASEMPTY") <> "", aIntPar.value("GEN", "USEASEMPTY"), "#")
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        aHdrRec = New TDataRec(aIntPar)
        Dim aKvB As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        For Each aKvB In aPar.getData()
            aTStrRec = aKvB.Value
            If valid_Hdr_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            End If
        Next
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        fillData = True
    End Function

    Public Function valid_Hdr_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Hdr_Field = False
        If pTStrRec.Strucname = "" Or pTStrRec.Strucname = "HD" Then
            valid_Hdr_Field = isInArray(pTStrRec.Fieldname, Hd_Fields)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("DBG", "DUMPHEADER") <> "", aIntPar.value("DBG", "DUMPHEADER"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpHeader - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the DBG-DUMPHEADR Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PPM")
                Exit Sub
            End Try
            log.Debug("dumpHeader - " & "dumping to " & dumpHd)
            ' clear the Header
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Header
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aHdrRec.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

    Public Sub dumpData()
        Dim dumpDt As String = If(aIntPar.value("DBG", "DUMPDATA") <> "", aIntPar.value("DBG", "DUMPDATA"), "")
        If dumpDt <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapPPMExcelAddIn.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpDt)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpData - " & "No " & dumpDt & " Sheet in current workbook.")
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PPM")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB As KeyValuePair(Of String, TDataRec)
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec(aIntPar)
            Dim aDataRec_Am As New TDataRec(aIntPar)
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB In aData.aTDataDic
                aDataRec = aKvB.Value
                Dim aFieldArray() As String = {}
                Dim aValueArray() As String = {}
                For Each aTStrRec In aDataRec.aTDataRecCol
                    Array.Resize(aFieldArray, aFieldArray.Length + 1)
                    aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                    Array.Resize(aValueArray, aValueArray.Length + 1)
                    aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                Next
                aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                aRange.Value = aFieldArray
                aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                aRange.Value = aValueArray
                i += 2
            Next
        End If
    End Sub

End Class
