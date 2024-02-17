Imports SAP.Middleware.Connector

Public Class OutputHelper
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub write(aOUTwsName As String, aLOff As UInt64, ByRef aIntPar As SAPCommon.TStr, ByRef aWB As Excel.Workbook, ByRef aTable As IRfcTable,
                          Optional pClear As Boolean = True, Optional pClearColumn As UInt64 = 1)
        Dim aItem As New TDataRec(aIntPar)
        Dim jMaxOUT As UInt64 = 0
        Dim aOUTws As Excel.Worksheet = Nothing
        Dim aFieldname As String
        Dim i As UInt64
        Dim j As UInt64
        Try
            aOUTws = aWB.Worksheets(aOUTwsName)
        Catch Exc As System.Exception
            log.Debug("SapPPMRibbonTimesheet.exec - " & "No " & aOUTwsName & " Sheet in current workbook")
        End Try
        ' output data
        If Not aOUTws Is Nothing Then
            i = aLOff + 1
            Do
                jMaxOUT += 1
            Loop While Not String.IsNullOrEmpty(CStr(aOUTws.Cells(aLOff - 3, jMaxOUT + 1).value))
            ' clear the output area
            Dim aRange As Excel.Range
            Dim aStartCell As String
            If pClear Then
                aStartCell = "A" & i
                If Not String.IsNullOrEmpty(CStr(aOUTws.Cells(i, pClearColumn).Value)) Then
                    aRange = aOUTws.Range(aStartCell)
                    Do
                        i += 1
                    Loop While Not String.IsNullOrEmpty(CStr(aOUTws.Cells(i, pClearColumn).Value))
                    aRange = aOUTws.Range(aRange, aOUTws.Cells(i, 1))
                    aRange.EntireRow.Delete()
                End If
            End If
            i = aLOff + 1
            Dim aValueArray() As String = {}
            Array.Resize(aValueArray, jMaxOUT)
            For k As Integer = 0 To aTable.Count - 1
                For j = 1 To jMaxOUT
                    aFieldname = aItem.name2Fieldname(CStr(aOUTws.Cells(1, j).value))
                    Try
                        aValueArray(j - 1) = CStr(aTable(k).GetValue(aFieldname))
                    Catch Exc As System.Exception
                        log.Debug("SapPPMRibbonTimesheet.exec - " & "Exception=" & Exc.ToString)
                    End Try
                Next j
                aRange = aOUTws.Range(aOUTws.Cells(i, 1), aOUTws.Cells(i, jMaxOUT))
                aRange.Value = aValueArray
                i += 1
            Next k
        End If
    End Sub
End Class
