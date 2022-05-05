Public Class BdgTipoOperacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTipoOperacion"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)

        data("ProponerArticulo") = True
        data("PermitirMerma") = True
        data("PorcMermaMaxima") = 0

    End Sub

    Public Overloads Function SelOnPrimaryKey(ByVal IDTipoOperacion As String) As DataTable
        Return MyBase.SelOnPrimaryKey(IDTipoOperacion)
    End Function

    Public Overloads Function GetItemRow(ByVal IDTipoOperacion As String) As DataRow
        Dim dt As DataTable = New BdgTipoOperacion().SelOnPrimaryKey(IDTipoOperacion)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el Tipo de Operación")
        Else : Return dt.Rows(0)
        End If
    End Function

    Public Function GetChilds() As DataSet
        Dim dsResult As New DataSet
        Dim oTOM As New BdgTipoOperacionMaterial
        Dim oTOC As New BdgTipoOperacionCentro
    End Function

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTipoRutaRepetida)
    End Sub

    <Task()> Public Shared Sub ValidarTipoRutaRepetida(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso data("IDTipoRuta") & String.Empty <> data("IDTipoRuta", DataRowVersion.Original) & String.Empty) AndAlso Length(data("IDTipoRuta")) > 0 Then
            Dim dtOperacionRuta As DataTable = New BdgTipoOperacion().Filter(New StringFilterItem("IDTipoRuta", data("IDTipoRuta")))
            If dtOperacionRuta.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Tipo Ruta indicado ya está asociado a otro Tipo de Operación.")
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Function ProponerArticulo(ByVal IDTipoOperacion As String, ByVal services As ServiceProvider) As Boolean
        Dim blProponerArticulo As Boolean = True
        Dim dt As DataTable = New BdgTipoOperacion().SelOnPrimaryKey(IDTipoOperacion)
        If dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el Tipo de Operación | ", IDTipoOperacion)
        Else
            blProponerArticulo = dt.Rows(0)("ProponerArticulo")
        End If
        Return blProponerArticulo
    End Function

End Class


Public Class BdgTipoOperacionInfo
    Inherits ClassEntityInfo

    Public IDTipoOperacion As String
    Public DescTipoOperacion As String

    Public TipoMovimiento As Integer
    Public TipoOrigen As Integer?
    Public TipoDestino As Integer?

    Public IDAnalisis As String
    Public RequiereOF As Boolean
    Public IDRule As Integer?

    Public ImputacionPrevMaterial As Boolean
    Public ImputacionPrevMod As Boolean
    Public ImputacionPrevCentro As Boolean
    Public ImputacionPrevVarios As Boolean

    Public ImputacionRealMaterial As Boolean
    Public ImputacionRealMod As Boolean
    Public ImputacionRealCentro As Boolean
    Public ImputacionRealVarios As Boolean

    Public ProponerArticulo As Boolean
    Public PermitirMerma As Boolean
    Public PorcMermaMaxima As Double

    Public IdClasificacion As Integer?
    Public IDEstadoVino As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New BdgTipoOperacion().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Tipo de Operación | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub
End Class


Public Enum TipoMovimiento
    SinMovimiento
    DeUnoAUno
    DeUnoAVarios
    DeVariosAUno
    SinOrigen   'Se crea un vino huerfano, sin estructura
    CrearOrigen     'se crea la estructura de un vino huerfano
    DeVariosAVarios
    Salida
End Enum

<Serializable()> _
Public Class _TO
    Public Const IDTipoOperacion As String = "IDTipoOperacion"
    Public Const DescTipoOperacion As String = "DescTipoOperacion"
    Public Const TipoMovimiento As String = "TipoMovimiento"
    Public Const TipoOrigen As String = "TipoOrigen"
    Public Const TipoDestino As String = "TipoDestino"
    Public Const IDAnalisis As String = "IDAnalisis"
    Public Const IDEstadoVino As String = "IDEstadoVino"
End Class