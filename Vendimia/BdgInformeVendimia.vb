'Public Class BdgInformeVendimia

'#Region "Constructor"

'    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

'    Public Sub New()
'        MyBase.New(cnEntidad)
'    End Sub


'    Private Const cnEntidad As String = "tbBdgMaestroInformeVendimia"

'#End Region

'#Region "Eventos Entidad"

'#End Region

'#Region "Funciones Públicas"

'    <Task()> Public Shared Function ObtenerDatosAnexo(ByVal data As StDatosAnexo, ByVal services As ServiceProvider) As DataTable
'        Dim filtro As New Filter
'        filtro.Add("AnexoVendimia", data.AnexoVendimia)
'        filtro.Add("Destino", data.Destino)
'        Return New BdgInformeVendimia().Filter(filtro)
'    End Function

'#End Region

'End Class


<Serializable()> _
    Public Class StDatosAnexo
    Public AnexoVendimia As Integer
    Public Destino As Integer

    Public Sub New()
    End Sub

    Public Sub New(ByVal AnexoVendimia As Integer, ByVal Destino As Integer)
        Me.AnexoVendimia = AnexoVendimia
        Me.Destino = Destino
    End Sub

End Class



Public Enum enumBdgAnexoVendimia
    AnexoII = 0
    AnexoIII = 1
    AnexoIIIA = 2
    AnexoIIIB = 3
End Enum

Public Enum enumBdgDestinoAnexoVendimia
    LaRioja = 0
    Navarra = 1
    Alava = 2
    CYL = 3
    LaMancha = 4
End Enum