Imports System
Imports System.IO
Imports System.Resources
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Application = NetOffice.ExcelApi.Application
Imports ExcelDna.Integration.CustomUI

<ComVisible(True)> _ 
Public Class CustomRibbon : Inherits ExcelRibbon
    
    private  _excel As Application
    private  _thisRibbon As IRibbonUI

    Overrides Function GetCustomUI(ribbonId As String) As String
        _excel = new Application(Nothing, ExcelDna.Integration.ExcelDnaUtil.Application)
        Dim ribbonXml As String = GetCustomRibbonXML()
        return ribbonXml        
    End Function

    Private Function GetCustomRibbonXML() As String
        Dim ribbonXml As String = Nothing
        Dim thisAssembly As Assembly = GetType(CustomRibbon).Assembly
        Dim resourceName As String = GetType(CustomRibbon).Namespace + ".CustomRibbon.xml"

        Using stream as Stream = thisAssembly.GetManifestResourceStream(resourceName)
        Using reader As StreamReader = new StreamReader(stream)
            ribbonXml = reader.ReadToEnd()                
        End Using
        End Using

        If (ribbonXml Is Nothing) Then
            throw New MissingManifestResourceException(resourceName)
        End If
        Return ribbonXml
    End Function

    Public Sub OnLoad(ribbon As IRibbonUI)
        _thisRibbon = ribbon
        If (_excel.ActiveWorkbook Is Nothing) Then
            _excel.Workbooks.Add()
        End If
    End Sub

    Public Sub OnPressMe(control As IRibbonControl)
            Using controller As ExcelController = New ExcelController(_excel, _thisRibbon)
                controller.PressMe()            
            End Using
    End Sub

End Class
