Imports System
Imports Application = NetOffice.ExcelApi.Application
Imports ExcelDna.Integration.CustomUI
Imports NetOffice.ExcelApi


Public Class ExcelController : Implements IDisposable
    private readonly _modelingRibbon As IRibbonUI
    protected readonly _excel As Application

    public Sub New(excel As Application, modelingRibbon As IRibbonUI)
        _modelingRibbon = modelingRibbon
        _excel = excel
    End Sub

    public Sub PressMe()
        Dim activeSheet As Worksheet = CType(_excel.ActiveSheet, Worksheet)
        activeSheet.Range("A1").Value = "Hello, World!"
    End Sub



#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region


End Class

