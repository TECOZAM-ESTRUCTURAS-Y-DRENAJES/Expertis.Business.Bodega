Public Class BdgVariable

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVariable"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDVariable As String) As DataRow
        Dim dt As DataTable = New BdgVariable().SelOnPrimaryKey(IDVariable)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la variable |", IDVariable)
        Else : Return dt.Rows(0)
        End If
    End Function

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data(_Vr.TipoVariable) = BdgTipoVariable.Numerica
        data(_Vr.FactorEscala) = 1.0
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIntervalo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarEscala)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDVariable")) > 0 Then
                If IsNumeric(Left(data("IDVariable"), 1)) Then
                    ApplicationService.GenerateError("El código de la variable no puede empezar por un número.")
                End If
            Else : ApplicationService.GenerateError("El código de la Variable es obligatorio")
            End If
            ProcessServer.ExecuteTask(Of String)(AddressOf ValidateDuplicateKey, data("IDVariable"), services)
        End If
        If Length(data("DescVariable")) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarIntervalo(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dblMax As Double = Nz(data("Maximo"), 0)
        Dim dblMin As Double = Nz(data("Minimo"), 0)
        If dblMax <> 0 Or dblMin <> 0 Then
            If dblMin > dblMax Then ApplicationService.GenerateError("El intervalo de máximo y mínimo no esta definido correctamente")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarEscala(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data(_Vr.TipoVariable), BdgTipoVariable.Numerica) = BdgTipoVariable.Numerica Then
            If Nz(data("FactorEscala"), 0) = 0 Then
                ApplicationService.GenerateError("El factor de escala no es válido.")
            End If
        End If
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDContador")) > 0 Then
                data("IDVariable") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDVariable As String, ByVal services As ServiceProvider)
        If New BdgVariable().SelOnPrimaryKey(IDVariable).Rows.Count = 0 Then
            ApplicationService.GenerateError("La Variable: | ya existe en la tabla: |.", IDVariable, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDVariable As String, ByVal services As ServiceProvider)
        If New BdgVariable().SelOnPrimaryKey(IDVariable).Rows.Count > 0 Then
            ApplicationService.GenerateError("La variable introducida ya existe en la tabla: |.", cnEntidad)
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _Vr
    Public Const IDVariable As String = "IDVariable"
    Public Const DescVariable As String = "DescVariable"
    Public Const IDContador As String = "IDContador"
    Public Const TipoVariable As String = "TipoVariable"
    Public Const Abreviatura As String = "Abreviatura"
    Public Const UdMedida As String = "UdMedida"
    Public Const Maximo As String = "Maximo"
    Public Const Minimo As String = "Minimo"
    Public Const Lista As String = "Lista"
    Public Const ColorMaximo As String = "ColorMaximo"
    Public Const ColorMinimo As String = "ColorMinimo"
    Public Const FactorEscala As String = "FactorEscala"
End Class

Public Enum BdgTipoVariable
    Numerica = 0
    Alfanumerica = 1
End Enum