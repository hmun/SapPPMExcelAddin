Imports SAP.Middleware.Connector

Public Class InputHelper
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub read(aINPwsName As String, aLOff As UInt64, aKeyPart As String, ByRef aWB As Excel.Workbook, ByRef aItems As TData)
        Dim jMaxINP As UInt64 = 0
        Dim aINPws As Excel.Worksheet = Nothing
        Dim j As UInt64
        Dim jMax As UInt64
        Dim aKey As String
        Try
            aINPws = aWB.Worksheets(aINPwsName)
        Catch Exc As System.Exception
            log.Debug("InputHelper - " & "No " & aINPwsName & " Sheet in current workbook")
        End Try
        ' read data
        If Not aINPws Is Nothing Then
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
            Loop While CStr(aINPws.Cells(aLOff - 3, jMax + 1).value) <> ""
        End If
        Dim i As UInt64 = aLOff + 1
        If Not aINPws Is Nothing Then
            ' collect the IDS items
            While CStr(aINPws.Cells(i, 1).Value) <> ""
                aKey = aKeyPart + "_" + CStr(i)
                For j = 1 To jMax
                    If CStr(aINPws.Cells(1, j).value) <> "N/A" And CStr(aINPws.Cells(1, j).value) <> "" Then
                        aItems.addValue(aKey, CStr(aINPws.Cells(aLOff - 3, j).value), CStr(aINPws.Cells(i, j).value),
                                        CStr(aINPws.Cells(aLOff - 2, j).value), CStr(aINPws.Cells(aLOff - 1, j).value), pEmty:=False,
                                        pEmptyChar:="")
                    End If
                Next
                i += 1
            End While
        End If
    End Sub
End Class
