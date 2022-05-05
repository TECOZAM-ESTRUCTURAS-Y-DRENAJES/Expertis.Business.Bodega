Public Class BdgFincaHistorico
#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbBdgFincaHistorico"

#End Region

#Region "Eventos Entidad"

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
       Public Class stRegistroHistoricoFinca
        Public IDFinca As Guid
        Public FechaMedicion As Date
        Public IDVariedad As String
        Public IDMarco As String
        Public Superficieviñedo As Double
        Public NCepas As Double
        Public IDOrientacion As String
        Public IDPortaInjerto As String
        Public IDTipoViñedo As String
        Public IDRiego As String
        Public IDPoda As String

        Public Clones As DataTable

        Public Sub New(ByVal IDFinca As Guid, ByVal FechaMedicion As Date, ByVal IDVariedad As String, ByVal IDMarco As String, _
                       ByVal Superficieviñedo As Double, ByVal NCepas As Double, ByVal IDOrientacion As String, ByVal IDPortaInjerto As String, _
                       ByVal IDTipoViñedo As String, ByVal IDRiego As String, ByVal IDPoda As String, ByVal clones As DataTable)
            Me.IDFinca = IDFinca
            Me.FechaMedicion = FechaMedicion
            Me.IDVariedad = IDVariedad
            Me.IDMarco = IDMarco
            Me.Superficieviñedo = Superficieviñedo
            Me.NCepas = NCepas
            Me.IDOrientacion = IDOrientacion
            Me.IDPortaInjerto = IDPortaInjerto
            Me.IDTipoViñedo = IDTipoViñedo
            Me.IDPoda = IDPoda
            Me.IDRiego = IDRiego
            Me.Clones = clones
        End Sub
    End Class

    <Task()> _
    Public Shared Sub RegistroHistoricoFinca(ByVal data As stRegistroHistoricoFinca, ByVal services As ServiceProvider)
        Dim bsnBdgFincaHistorico As New BdgFincaHistorico
        Dim dttHistoricoFinca As DataTable = bsnBdgFincaHistorico.AddNewForm
        With dttHistoricoFinca
            .Rows(0)("IDFinca") = data.IDFinca
            .Rows(0)("FechaMedicion") = data.FechaMedicion
            .Rows(0)("IDVariedad") = data.IDVariedad
            .Rows(0)("IDMarco") = data.IDMarco
            .Rows(0)("Superficieviñedo") = data.Superficieviñedo
            .Rows(0)("NCepas") = data.NCepas
            .Rows(0)("IDOrientacion") = data.IDOrientacion
            .Rows(0)("IDPortaInjerto") = data.IDPortaInjerto
            .Rows(0)("IDTipoViñedo") = data.IDTipoViñedo
            .Rows(0)("IDRiego") = data.IDRiego
            .Rows(0)("IDPoda") = data.IDPoda

            If Not data.Clones Is Nothing AndAlso data.Clones.Rows.Count > 0 Then
                .Rows(0)("IDClonVariedad") = data.Clones.Rows(0)("IDClonVariedad")
                .Rows(0)("Clones") = GenerarDescripcionClones(data.Clones)
            End If

        End With
        bsnBdgFincaHistorico.Update(dttHistoricoFinca)
    End Sub

    Protected Shared Function GenerarDescripcionClones(ByVal dttClones As DataTable) As String
        Dim strResult As String = String.Empty
        For Each dtr As DataRow In dttClones.Rows
            If (dtr.RowState <> DataRowState.Deleted) Then
                strResult = String.Format("{0}{1}-{2}; ", strResult, dtr("IDClonVariedad"), xRound(dtr("Superficie"), 2))
            End If
        Next
        Return strResult
    End Function

#End Region


End Class