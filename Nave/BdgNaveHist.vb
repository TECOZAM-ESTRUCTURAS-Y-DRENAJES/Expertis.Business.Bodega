Public Class BdgNaveHist

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgNaveHist"

#End Region

End Class

<Serializable()> _
Public Class _NH
    Public Const IDHist As String = "IDHist"
    Public Const IDNave As String = "IDNave"
    Public Const FechaDesde As String = "FechaDesde"
    Public Const FechaHasta As String = "FechaHasta"
    Public Const Fecha As String = "Fecha"
End Class