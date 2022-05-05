Public Class BdgDepositoInfo
    Inherits ClassEntityInfo

    Public IDDeposito As String
    Public DescDeposito As String

    Public IDContador As String

    Public Capacidad As Double
    Public Ocupacion As Double

    Public PosX As Integer
    Public PosY As Integer
    Public Ancho As Integer
    Public Alto As Integer

    Public Bloqueado As Boolean

    Public IDNave As String
    Public TipoDeposito As Integer
    Public SinLimite As Boolean
    Public CapacidadKg As Double
    Public NSubDep As Integer
    Public TieneSub As Boolean

    Public MultiplesVinos As Boolean

    Public Base As Integer
    Public Altura As Integer
    Public Disposicion As Integer
    Public Orientacion As Integer
    Public Llenado As Integer

    Public IDUDMedida As String

    Public NoRepresentar As Boolean
    Public NoTrazar As Boolean
    Public UsarBarricaComoLote As Boolean

    Public JaulonRequerido As Boolean

    Public IDActivo As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New BdgDeposito().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Depósito | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub
End Class

Public Class BdgDeposito

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgDeposito"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDDeposito As String) As DataRow
        Dim dt As DataTable = New BdgDeposito().SelOnPrimaryKey(IDDeposito)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el depósito |", IDDeposito)
        Else : Return dt.Rows(0)
        End If
    End Function

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarID)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescripcion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarUDMedida)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarNave)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarUDMedida)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub
    <Task()> Public Shared Sub ValidarNave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDNave")) = 0 Then ApplicationService.GenerateError("La Nave es un dato Obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarDescripcion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescDeposito")) = 0 Then ApplicationService.GenerateError("La descripción del depósito es un dato Obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarID(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDDeposito")) = 0 Then ApplicationService.GenerateError("El depósito es un dato Obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarUDMedida(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDUDMedida")) = 0 Then ApplicationService.GenerateError("La Unidad de Medida es un dato Obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            'Si el depósito tiene más de un vino no se puede cambiar el tipo de depósito.
            If (data(_D.TipoDeposito) <> data(_D.TipoDeposito, DataRowVersion.Original)) Then
                Dim dtDepV As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgDepositoVino.SelOnIDDeposito, data(_D.IDDeposito), services)
                If dtDepV.Rows.Count > 1 Then ApplicationService.GenerateError("No se puede realizar la modificación. El depósito {0} contiene actualmente más de un vino.", data(_D.IDDeposito))
            End If

            If Not CBool(data(_D.MultiplesVinos)) Then
            End If
            If data(_D.IDUDMedida) <> Nz(data(_D.IDUDMedida, DataRowVersion.Original), String.Empty) Then
                If data(_D.Ocupacion) <> 0 Then ApplicationService.GenerateError("No se puede cambiar la unidad de medida de un depósito cuando este no está vacio.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtAux As DataTable = New BdgDeposito().SelOnPrimaryKey(data("IDDeposito"))
            If Not DtAux Is Nothing AndAlso DtAux.Rows.Count <> 0 Then
                ApplicationService.GenerateError("El Depósito introducido ya existe en la base de datos", data("IDDeposito"))
            End If
        End If
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_D.IDContador)) > 0 Then
                data(_D.IDDeposito) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data(_D.IDContador), services)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ObtenerOpcionMultipleVino(ByVal tipodeposito As Business.Bodega.TipoDeposito, ByVal services As ServiceProvider) As Integer
        Select Case tipodeposito
            Case Business.Bodega.TipoDeposito.Almacen, Business.Bodega.TipoDeposito.Barricas, Business.Bodega.TipoDeposito.Botellero
                Return 1
            Case Business.Bodega.TipoDeposito.Deposito
                Return 0
        End Select
    End Function

    <Task()> Public Shared Function Almacen(ByVal IDDeposito As String, ByVal services As ServiceProvider) As String
        Return ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgNave.Almacen, New BdgDeposito().GetItemRow(IDDeposito)(_D.IDNave) & String.Empty, services)
    End Function

    <Task()> Public Shared Function ObtenerAlmacenDeposito(ByVal IDDeposito As String, ByVal services As ServiceProvider) As String
        Dim dtDep As DataTable = New BdgDeposito().SelOnPrimaryKey(IDDeposito)
        Dim IDAlmacen As String = String.Empty
        If dtDep.Rows.Count > 0 Then
            IDAlmacen = ProcessServer.ExecuteTask(Of String, String)(AddressOf Business.Bodega.BdgNave.Almacen, dtDep.Rows(0)("IDNave") & String.Empty, services)
        End If
        If Length(IDAlmacen) = 0 Then IDAlmacen = New Parametro().AlmacenPredeterminado
        If Length(IDAlmacen) = 0 Then ApplicationService.GenerateError("No hay un almacén predeterminado para el depósito")
        Return IDAlmacen
    End Function

    <Task()> Public Shared Function UsarBarricaComoLote(ByVal IDDeposito As String, ByVal services As ServiceProvider) As Boolean
        Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
        Dim DptoInfo As BdgDepositoInfo = Depositos.GetEntity(IDDeposito)
        If DptoInfo.TipoDeposito = TipoDeposito.Barricas Then
            Return DptoInfo.UsarBarricaComoLote
        End If
    End Function

    <Task()> Public Shared Function RequerirJaulon(ByVal IDDeposito As String, ByVal services As ServiceProvider) As Boolean
        Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
        Dim DptoInfo As BdgDepositoInfo = Depositos.GetEntity(IDDeposito)
        If DptoInfo.TipoDeposito = TipoDeposito.Botellero Then
            Return DptoInfo.JaulonRequerido
        End If
    End Function

#End Region

End Class

Public Enum TipoDeposito
    Deposito
    Barricas
    Botellero
    Almacen
End Enum

<Serializable()> _
Public Class _D
    Public Const IDDeposito As String = "IDDeposito"
    Public Const DescDeposito As String = "DescDeposito"
    Public Const IDContador As String = "IDContador"
    Public Const TipoDeposito As String = "TipoDeposito"
    Public Const Capacidad As String = "Capacidad"
    Public Const SinLimite As String = "SinLimite"
    Public Const Ocupacion As String = "Ocupacion"
    Public Const CapacidadKg As String = "CapacidadKg"
    Public Const NSubDep As String = "NSubDep"
    Public Const TieneSub As String = "TieneSub"
    Public Const MultiplesVinos As String = "MultiplesVinos"
    Public Const IDNave As String = "IDNave"
    Public Const PosX As String = "PosX"
    Public Const PosY As String = "PosY"
    Public Const Ancho As String = "Ancho"
    Public Const Alto As String = "Alto"
    Public Const Bloqueado As String = "Bloqueado"
    Public Const Color As String = "Color"
    Public Const Base As String = "Base"
    Public Const Altura As String = "Altura"
    Public Const Disposicion As String = "Disposicion"
    Public Const Orientacion As String = "Orientacion"
    Public Const Llenado As String = "Llenado"
    Public Const IDUDMedida As String = "IDUDMedida"
    Public Const NoRepresentar As String = "NoRepresentar"
    Public Const NoTrazar As String = "NoTrazar"
    Public Const UsarBarricaComoLote As String = "UsarBarricaComoLote"
End Class