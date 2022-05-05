Imports Solmicro.Expertis.Business.Bodega.BdgDAACabecera

Public Class BdgDAALineaPaquete

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbDAALineaPaquete"

#End Region

#Region "Eventos Entidad"
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub


    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDDaaLineaPaquete") = Guid.NewGuid
    End Sub

    <Serializable()> _
    Public Class stCrearDAALineaPaqueteInfo
        Public IDDaaLinea As String
        Public IDTipoEmbalaje As String
        Public NumeroBultos As Double
        Public IdentificadorPrecinto As String
        Public InformacionPrecinto As String

        Public Sub New(ByVal IDDaaLinea As String, ByVal IDTipoEmablaje As String, ByVal numeroBultos As Double, ByVal identificadorPrecinto As String, ByVal informacionPrecinto As String)
            Me.IDDaaLinea = IDDaaLinea
            Me.IDTipoEmbalaje = IDTipoEmablaje
            Me.NumeroBultos = numeroBultos
            Me.IdentificadorPrecinto = identificadorPrecinto
            Me.InformacionPrecinto = informacionPrecinto
        End Sub

    End Class

    <Task()> Public Shared Function CrearDaaLineasPaquete(ByVal data As stCrearDAALineaPaqueteInfo, ByVal services As ServiceProvider) As DataTable
        'por tenerlo separado
        Dim dttLineaPaquete As DataTable = New BdgDAALineaPaquete().AddNewForm
        dttLineaPaquete.Rows(0)("IDDaaLineaPaquete") = Guid.NewGuid
        dttLineaPaquete.Rows(0)("IDDaaLinea") = data.IDDaaLinea
        dttLineaPaquete.Rows(0)("IDTipoEmbalaje") = data.IDTipoEmbalaje
        dttLineaPaquete.Rows(0)("NumeroBultos") = data.NumeroBultos
        dttLineaPaquete.Rows(0)("IdentificadorPrecinto") = data.IdentificadorPrecinto & String.Empty
        dttLineaPaquete.Rows(0)("InformacionPrecinto") = data.InformacionPrecinto & String.Empty
        Return dttLineaPaquete
    End Function

#End Region

End Class