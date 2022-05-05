Public Class BdgVinoEstructura

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoEstructura"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function SelOnComponente(ByVal IDVino As Guid, ByVal services As ServiceProvider) As DataTable
        Return New BdgVinoEstructura().Filter(New GuidFilterItem(_VE.IDVinoComponente, FilterOperator.Equal, IDVino))
    End Function

    <Task()> Public Shared Function SelOnVino(ByVal IDVino As Guid, ByVal services As ServiceProvider) As DataTable
        Return New BdgVinoEstructura().Filter(New GuidFilterItem(_VE.IDVino, FilterOperator.Equal, IDVino))
    End Function

    <Serializable()> _
    Public Class StSelOnVinoOperacion
        Public IDVino As Guid
        Public Operacion As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Operacion As String)
            Me.IDVino = IDVino
            Me.Operacion = Operacion
        End Sub
    End Class

    <Task()> Public Shared Function SelOnVinoOperacion(ByVal data As StSelOnVinoOperacion, ByVal services As ServiceProvider) As DataTable
        Dim oFltr As New Filter
        oFltr.Add(New GuidFilterItem(_VE.IDVino, data.IDVino))
        oFltr.Add(New StringFilterItem(_VE.Operacion, data.Operacion))
        Return New BdgVinoEstructura().Filter(oFltr)
    End Function

    <Serializable()> _
    Public Class StSelOnVinoComponente
        Public IDVino As Guid
        Public IDVinoComponente As Guid

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal IDVinoComponente As Guid)
            Me.IDVino = IDVino
            Me.IDVinoComponente = IDVinoComponente
        End Sub
    End Class

    <Task()> Public Shared Function SelOnVinoComponente(ByVal data As StSelOnVinoComponente, ByVal services As ServiceProvider) As DataTable
        Dim oFltr As New Filter
        oFltr.Add(New GuidFilterItem(_VE.IDVino, data.IDVino))
        oFltr.Add(New GuidFilterItem(_VE.IDVinoComponente, data.IDVinoComponente))
        Return New BdgVinoEstructura().Filter(oFltr)
    End Function

#End Region

End Class

<Serializable()> _
Public Class _VE
    Public Const IDVino As String = "IDVino"
    Public Const IDVinoComponente As String = "IDVinoComponente"
    Public Const Operacion As String = "Operacion"
    Public Const Cantidad As String = "Cantidad"
    Public Const Merma As String = "Merma"
    Public Const Factor As String = "Factor"
    Public Const NOperacion As String = "NOperacion"
End Class