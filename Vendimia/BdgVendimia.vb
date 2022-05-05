Public Class BdgVendimia

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVendimia"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal Vendimia As Integer) As DataRow
        Dim dt As DataTable = New BdgVendimia().SelOnPrimaryKey(Vendimia)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la vendimia |", Vendimia)
        Else : Return dt.Rows(0)
        End If
    End Function

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarEntradas)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarDatosCartillistas)
    End Sub

    <Task()> Public Shared Sub ComprobarEntradas(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim FilVend As New Filter
        FilVend.Add("Vendimia", FilterOperator.Equal, data("Vendimia"))

        Dim dtEnt As DataTable = New BdgEntrada().Filter(FilVend, , "Count(*) As NumEntradas")
        If dtEnt.Rows.Count > 0 AndAlso Nz(dtEnt.Rows(0)("NumEntradas"), 0) > 0 Then
            ApplicationService.GenerateError("Hay entradas para la vendimia {0}.", Quoted(data("Vendimia")))
        End If
    End Sub

    <Task()> Public Shared Sub BorrarDatosCartillistas(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim UdtDts As New UpdatePackage

        Dim FilVend As New Filter
        FilVend.Add("Vendimia", FilterOperator.Equal, data("Vendimia"))

        Dim DtCV As DataTable = New BdgCartillistaVendimia().Filter(FilVend)
        For Each DrCV As DataRow In DtCV.Select
            DrCV.Delete()
        Next
        UdtDts.Add(DtCV)

        Dim DtCM As DataTable = New BdgCartillistaMunicipio().Filter(FilVend, "IDCartillista")
        For Each DrCM As DataRow In DtCM.Select
            DrCM.Delete()
        Next
        UdtDts.Add(DtCM)

        Dim DtCF As DataTable = New BdgCartillistaFinca().Filter(FilVend)
        For Each DrCF As DataRow In DtCF.Select
            DrCF.Delete()
        Next
        UdtDts.Add(DtCF)

        BusinessHelper.UpdatePackage(UdtDts)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarPrecioUvaT)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarPrecioUvaB)
        'updateProcess.AddTask(Of StCrearVendimia)(AddressOf CrearVendimia)
    End Sub

    <Task()> Public Shared Sub ActualizarPrecioUvaT(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data(_VDM.PrecioUvaT) <> data(_VDM.PrecioUvaT, DataRowVersion.Original) Then
                Dim dblP As Double = Nz(data(_VDM.PrecioUvaT), 0)
                Dim StPrecio As New StSetPrecioUva(data(_VDM.Vendimia), BdgTipoVariedad.Tinta, dblP, data(_VDM.IDArticuloT))
                ProcessServer.ExecuteTask(Of StSetPrecioUva)(AddressOf SetPrecioUva, StPrecio, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPrecioUvaB(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data(_VDM.PrecioUvaB) <> data(_VDM.PrecioUvaB, DataRowVersion.Original) Then
                Dim dblP As Double = Nz(data(_VDM.PrecioUvaB), 0)
                Dim StPrecio As New StSetPrecioUva(data(_VDM.Vendimia), BdgTipoVariedad.Blanca, dblP, data(_VDM.IDArticuloB))
                ProcessServer.ExecuteTask(Of StSetPrecioUva)(AddressOf SetPrecioUva, StPrecio, services)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function UltimaVendimia(ByVal data As Object, ByVal services As ServiceProvider) As Integer
        Dim DtVendimia As DataTable = New BdgVendimia().Filter("TOP 1 Vendimia", , "Vendimia DESC")
        If DtVendimia.Rows.Count > 0 Then
            Return DtVendimia.Rows(0)("Vendimia")
        End If
    End Function

    <Serializable()> _
    Public Class StCrearVendimia
        Public Vendimia As Integer
        Public VendimiaOrigen As Integer
        Public IDCartillista As String
        Public SoloCrearCupos As Boolean

        Public RdtoxHaT As Double
        Public RdtoxHaB As Double
        Public RdtoT As Double
        Public RdtoB As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal Vendimia As Integer, ByVal VendimiaOrigen As Integer, ByVal RdtoxHaT As Double, ByVal RdtoxHaB As Double, ByVal RdtoT As Double, ByVal RdtoB As Double)
            Me.Vendimia = Vendimia
            Me.VendimiaOrigen = VendimiaOrigen
            Me.RdtoxHaT = RdtoxHaT
            Me.RdtoxHaB = RdtoxHaB
            Me.RdtoT = RdtoT
            Me.RdtoB = RdtoB
        End Sub
    End Class

    Public Class ProcInfoVendimiaCupos
        Public mList As New List(Of DataTable)
    End Class
    <Task()> Public Shared Sub ValidarDatosVendimia(ByVal Vendimia As Integer, ByVal services As ServiceProvider)
        If Vendimia <= 0 Then
            ApplicationService.GenerateError("El valor asignado a la vendimia no es válido.")
        Else
            If Length(Vendimia) <> 4 Then ApplicationService.GenerateError("El valor asignado a la vendimia no es válido.")
        End If
    End Sub
    <Task()> Public Shared Sub CrearVendimia(ByVal data As StCrearVendimia, ByVal services As ServiceProvider)

        Dim ProcInfo As ProcInfoVendimiaCupos = services.GetService(Of ProcInfoVendimiaCupos)()

        ProcessServer.ExecuteTask(Of Integer)(AddressOf ValidarDatosVendimia, data.Vendimia, services)

        Dim uPck As New UpdatePackage

        '//Vendimia
        Dim dtVendimia As DataTable = New BdgVendimia().SelOnPrimaryKey(data.Vendimia)
        If dtVendimia.Rows.Count = 0 Then
            Dim drVendimia As DataRow = dtVendimia.NewRow
            drVendimia(_VDM.Vendimia) = data.Vendimia
            drVendimia(_VDM.RdtoxHaT) = data.RdtoxHaT
            drVendimia(_VDM.RdtoxHaB) = data.RdtoxHaB
            drVendimia(_VDM.RdtoT) = data.RdtoT
            drVendimia(_VDM.RdtoB) = data.RdtoB
            dtVendimia.Rows.Add(drVendimia)
        Else
            ApplicationService.GenerateError("La vendimia {0} ya existe en el sistema.", Quoted(data.Vendimia))
        End If
        ProcInfo.mList.Add(dtVendimia)

        '//Municipio
        Dim dtM As DataTable = New BdgMunicipio().Filter()
        For Each rwM As DataRow In dtM.Select
            rwM("RdtoT") = data.RdtoT
            rwM("RdtoB") = data.RdtoB
        Next
        ProcInfo.mList.Add(dtM)

        ProcessServer.ExecuteTask(Of StCrearVendimia)(AddressOf CrearCuposVendimia, data, services)
    End Sub
    <Task()> Public Shared Sub CrearCuposVendimia(ByVal data As StCrearVendimia, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Integer)(AddressOf ValidarDatosVendimia, data.Vendimia, services)

        If data.SoloCrearCupos Then
            Dim dtVendimia As DataTable = New BdgVendimia().SelOnPrimaryKey(data.Vendimia)
            If dtVendimia.Rows.Count > 0 Then
                data.RdtoxHaT = dtVendimia.Rows(0)(_VDM.RdtoxHaT)
                data.RdtoxHaB = dtVendimia.Rows(0)(_VDM.RdtoxHaB)
                data.RdtoT = dtVendimia.Rows(0)(_VDM.RdtoT)
                data.RdtoB = dtVendimia.Rows(0)(_VDM.RdtoB)
            End If
        End If

        Dim ProcInfo As ProcInfoVendimiaCupos = services.GetService(Of ProcInfoVendimiaCupos)()

        If data.VendimiaOrigen > 0 Then
            Dim fVendimiaOrigen As New Filter()
            fVendimiaOrigen.Add(New NumberFilterItem("Vendimia", data.VendimiaOrigen))

            Dim fVendimiaDestino As New Filter()
            fVendimiaDestino.Add(New NumberFilterItem("Vendimia", data.Vendimia))


            Dim fCartillista As New Filter()
            If Length(data.IDCartillista) > 0 Then fCartillista.Add(New StringFilterItem("IDCartillista", data.IDCartillista))
            Dim dtCartillistas As DataTable = New BdgCartillista().Filter(fCartillista, , "IDCartillista, CupoCartillaPor")
            If dtCartillistas.Rows.Count = 0 Then Exit Sub

            Dim CV As New BdgCartillistaVendimia
            Dim dtNewCV As DataTable = CV.AddNew

            Dim CM As New BdgCartillistaMunicipio
            Dim dtNewCM As DataTable = CM.AddNew

            Dim CF As New BdgCartillistaFinca
            Dim dtNewCF As DataTable = CF.AddNew

            Dim CVar As New BdgCartillistaVariedad
            Dim dtNewCVar As DataTable = CVar.AddNew

            Dim Cartillistas As List(Of DataRow) = (From c In dtCartillistas Order By c("IDCartillista") Select c).ToList
            For Each drCartillista As DataRow In Cartillistas
                Dim MaxTAcum As Double = 0
                Dim MaxBAcum As Double = 0
                Dim TotHaTAcum As Double = 0
                Dim TotHaBAcum As Double = 0

                '//CARTILLISTA MUNICIPIO
                Dim fCM As New Filter
                fCM.Add(fVendimiaOrigen)
                fCM.Add(New StringFilterItem("IDCartillista", drCartillista("IDCartillista")))
                Dim dtCM As DataTable = CM.Filter(fCM, "IDCartillista")
                If dtCM.Rows.Count > 0 Then
                    For Each drCM As DataRow In dtCM.Rows
                        Dim drNewCM As DataRow = dtNewCM.NewRow
                        For Each oCol As DataColumn In dtCM.Columns
                            drNewCM(oCol.ColumnName) = drCM(oCol)
                        Next
                        drNewCM("Vendimia") = data.Vendimia

                        drNewCM("RdtoT") = data.RdtoT
                        drNewCM("RdtoB") = data.RdtoB
                        drNewCM("MaxT") = drNewCM("HaT") * data.RdtoxHaT * data.RdtoT / 100
                        drNewCM("MaxB") = drNewCM("HaB") * data.RdtoxHaB * data.RdtoB / 100
                         dtNewCM.Rows.Add(drNewCM)
                    Next

                    If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Municipio Then
                        TotHaTAcum = (Aggregate c In dtNewCM Where c("IDCartillista") = drCartillista("IDCartillista") Into Sum(CDbl(c("HaT"))))
                        MaxTAcum = (Aggregate c In dtNewCM Where c("IDCartillista") = drCartillista("IDCartillista") Into Sum(CDbl(c("MaxT"))))
                        TotHaBAcum = (Aggregate c In dtNewCM Where c("IDCartillista") = drCartillista("IDCartillista") Into Sum(CDbl(c("HaB"))))
                        MaxBAcum = (Aggregate c In dtNewCM Where c("IDCartillista") = drCartillista("IDCartillista") Into Sum(CDbl(c("MaxB"))))
                    End If
                End If
                If Not dtNewCM Is Nothing AndAlso dtNewCM.Rows.Count > 0 Then ProcInfo.mList.Add(dtNewCM)

                '//CARTILLISTA FINCA
                Dim BEDataEngine As New BE.DataEngine()
                Dim fCF As New Filter
                fCF.Add(fVendimiaOrigen)
                fCF.Add(New StringFilterItem("IDCartillista", drCartillista("IDCartillista")))
                Dim dtCF As DataTable = BEDataEngine.Filter("frmBdgCartillistaFinca", fCF, , "IDCartillista")
                If dtCF.Rows.Count > 0 Then
                    For Each drCF As DataRow In dtCF.Rows
                        Dim drNewCF As DataRow = dtNewCF.NewRow
                        For Each oCol As DataColumn In dtNewCF.Columns
                            If oCol.ColumnName <> "Vendimia" AndAlso oCol.ColumnName <> "Maximo" Then
                                If dtCF.Columns.Contains(oCol.ColumnName) Then drNewCF(oCol.ColumnName) = drCF(oCol.ColumnName)
                            End If
                        Next
                        drNewCF("Vendimia") = data.Vendimia

                        ' drNewCF("Rdto") = drCF("Rdto")
                        Select Case drCF("TipoVariedad")
                            Case BdgTipoVariedad.Tinta
                                drNewCF("Maximo") = drCF("Superficie") * data.RdtoxHaT * drNewCF("Rdto") / 100
                                If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Finca Then
                                    TotHaTAcum += drCF("SuperficieViñedo")
                                    MaxTAcum += drCF("Maximo")
                                End If
                            Case BdgTipoVariedad.Blanca
                                drNewCF("Maximo") = drCF("Superficie") * data.RdtoxHaB * drNewCF("Rdto") / 100
                                If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Finca Then
                                    TotHaBAcum += drCF("SuperficieViñedo")
                                    MaxBAcum += drCF("Maximo")
                                End If
                        End Select
                        ''''''''''''''''
                        dtNewCF.Rows.Add(drNewCF)
                    Next
                    If Not dtNewCF Is Nothing AndAlso dtNewCF.Rows.Count > 0 Then ProcInfo.mList.Add(dtNewCF)

                End If

                '//CARTILLISTA VARIEDAD
                Dim fCVar As New Filter
                fCVar.Add(fVendimiaOrigen)
                fCVar.Add(New StringFilterItem("IDCartillista", drCartillista("IDCartillista")))
                Dim dtCVar As DataTable = BEDataEngine.Filter("frmBdgCartillistaVariedad", fCVar, , "IDCartillista")
                If dtCVar.Rows.Count > 0 Then
                    For Each drCVar As DataRow In dtCVar.Rows
                        Dim drNewCVar As DataRow = dtNewCVar.NewRow
                        For Each oCol As DataColumn In dtNewCVar.Columns
                            If oCol.ColumnName <> "Vendimia" Then
                                If dtCVar.Columns.Contains(oCol.ColumnName) Then drNewCVar(oCol.ColumnName) = drCVar(oCol.ColumnName)
                            End If
                        Next
                        drNewCVar("Vendimia") = data.Vendimia

                        Select Case drCVar("TipoVariedad")
                            Case BdgTipoVariedad.Tinta
                                If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Variedad Then
                                    MaxTAcum += drCVar("Maximo")
                                End If
                            Case BdgTipoVariedad.Blanca
                                If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Variedad Then
                                    MaxBAcum += drCVar("Maximo")
                                End If
                        End Select

                        dtNewCVar.Rows.Add(drNewCVar)
                    Next

                    If Not dtNewCVar Is Nothing AndAlso dtNewCVar.Rows.Count > 0 Then ProcInfo.mList.Add(dtNewCVar)
                End If

                '//CARTILLISTA VENDIMIA
                Dim fCV As New Filter
                fCV.Add(fVendimiaOrigen)
                fCV.Add(New StringFilterItem("IDCartillista", drCartillista("IDCartillista")))
                Dim dtCV As DataTable = CV.Filter(fCV, "IDCartillista")
                If dtCV.Rows.Count > 0 Then
                    For Each drCV As DataRow In dtCV.Rows
                        Dim drNewCV As DataRow = dtNewCV.NewRow
                        For Each oCol As DataColumn In dtCV.Columns
                            If oCol.ColumnName <> "Vendimia" AndAlso oCol.ColumnName <> "HaT" AndAlso oCol.ColumnName <> "MaxT" AndAlso oCol.ColumnName <> "HaB" AndAlso oCol.ColumnName <> "MaxB" Then
                                If dtCV.Columns.Contains(oCol.ColumnName) Then drNewCV(oCol.ColumnName) = drCV(oCol.ColumnName)
                            End If
                        Next
                        If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Manual Then
                            drNewCV("HaT") = drCV("HaT")
                        Else
                            drNewCV("HaT") = TotHaTAcum
                        End If
                        If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Manual Then
                            drNewCV("MaxT") = drCV("HaT") * data.RdtoxHaT * data.RdtoT / 100
                        Else
                            drNewCV("MaxT") = MaxTAcum
                        End If
                        If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Manual Then
                            drNewCV("HaB") = drCV("HaB")
                        Else
                            drNewCV("HaB") = TotHaBAcum
                        End If
                        If Nz(drCartillista("CupoCartillaPor"), BdgCupoCartillaPor.Manual) = BdgCupoCartillaPor.Manual Then
                            drNewCV("MaxB") = drCV("HaB") * data.RdtoxHaB * data.RdtoB / 100
                        Else
                            drNewCV("MaxB") = MaxBAcum
                        End If
                        drNewCV("Vendimia") = data.Vendimia
                        dtNewCV.Rows.Add(drNewCV)
                    Next

                    If Not dtNewCV Is Nothing AndAlso dtNewCV.Rows.Count > 0 Then ProcInfo.mList.Add(dtNewCV)
                End If

            Next
        End If

        'Actualizar todos los recordsets transaccionalmente
        AdminData.BeginTx()
        For Each dt As DataTable In ProcInfo.mList
            BusinessHelper.UpdateTable(dt)
        Next
        ProcInfo.mList.Clear()
        AdminData.CommitTx(True)
    End Sub


    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal Vendimia As Integer, ByVal services As ServiceProvider)
        If New BdgVendimia().SelOnPrimaryKey(Vendimia).Rows.Count = 0 Then
            ApplicationService.GenerateError("La Vendimia: | existe en la tabla: |.", Vendimia, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal Vendimia As Integer)
        If New BdgVendimia().SelOnPrimaryKey(Vendimia).Rows.Count > 0 Then
            ApplicationService.GenerateError("La Vendimia introducida ya existe en la tabla: |", cnEntidad)
        End If
    End Sub

    <Serializable()> _
    Public Class StArticuloUva
        Public Vendimia As Integer
        Public TipoUva As BdgTipoVariedad

        Public Sub New()
        End Sub

        Public Sub New(ByVal Vendimia As Integer, ByVal TipoUva As BdgTipoVariedad)
            Me.Vendimia = Vendimia
            Me.TipoUva = TipoUva
        End Sub
    End Class

    <Task()> Public Shared Function ArticuloUva(ByVal data As StArticuloUva, ByVal services As ServiceProvider) As String
        Dim Dr As DataRow = New BdgVendimia().GetItemRow(data.Vendimia)
        Dim strField As String
        If data.TipoUva = BdgTipoVariedad.Tinta Then
            strField = _VDM.IDArticuloT
        Else
            strField = _VDM.IDArticuloB
        End If
        If Length(Dr(strField)) = 0 Then
            ApplicationService.GenerateError("No se ha especificado un artículo para la uva en la vendimia")
        Else : Return Dr(strField)
        End If
    End Function

    <Serializable()> _
    Public Class StSetPrecioUva
        Public Vendimia As Integer
        Public TV As BdgTipoVariedad
        Public Precio As Double
        Public IDArticulo As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Vendimia As Integer, ByVal TV As BdgTipoVariedad, ByVal Precio As Double, ByVal IDArticulo As String)
            Me.Vendimia = Vendimia
            Me.TV = TV
            Me.Precio = Precio
            Me.IDArticulo = IDArticulo
        End Sub
    End Class

    <Task()> Public Shared Sub SetPrecioUva(ByVal data As StSetPrecioUva, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(_E.Vendimia, data.Vendimia)
        f.Add(_E.TipoVariedad, data.TV)

        Dim strIDAlmacen As String = New Parametro().AlmacenPredeterminado

        Dim dtEnts As DataTable = New BE.DataEngine().Filter("NegSetPrecioUva", f, , _ED.IDVino)
        Dim IDVino As Guid
        Dim dblQ As Double
        For Each oRw As DataRow In dtEnts.Rows
            If Not IDVino.Equals(oRw(_ED.IDVino)) Then
                If Not IDVino.Equals(Guid.Empty) Then
                    Dim StPrecio As New StSetPrecioUvaAlmacen(IDVino, data.IDArticulo, data.Precio, dblQ, strIDAlmacen)
                    ProcessServer.ExecuteTask(Of StSetPrecioUvaAlmacen)(AddressOf SetPrecioUvaAlmacen, StPrecio, services)
                End If
                IDVino = oRw(_ED.IDVino)
                dblQ = 0
            End If
            dblQ += oRw(_ED.Cantidad)
        Next
        If Not IDVino.Equals(Guid.Empty) Then
            Dim StPrecioAlm As New StSetPrecioUvaAlmacen(IDVino, data.IDArticulo, data.Precio, dblQ, strIDAlmacen)
            ProcessServer.ExecuteTask(Of StSetPrecioUvaAlmacen)(AddressOf SetPrecioUvaAlmacen, StPrecioAlm, services)
        End If
    End Sub

    <Serializable()> _
    Public Class StSetPrecioUvaAlmacen
        Public IDVino As Guid
        Public IDArticulo As String
        Public Precio As Double
        Public Q As Double
        Public IDAlmacen As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal IDArticulo As String, ByVal Precio As Double, ByVal Q As Double, ByVal IDAlmacen As String)
            Me.IDVino = IDVino
            Me.IDArticulo = IDArticulo
            Me.Precio = Precio
            Me.Q = Q
            Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Sub SetPrecioUvaAlmacen(ByVal data As StSetPrecioUvaAlmacen, ByVal services As ServiceProvider)
        Dim oVM As New BdgVinoMaterial
        Dim f As New Filter
        f.Add(_VM.IDVino, data.IDVino)
        f.Add(_VM.IDArticulo, data.IDArticulo)
        Dim dtVM As DataTable = oVM.Filter(f)
        Dim rwVM As DataRow
        If dtVM.Rows.Count > 0 Then
            rwVM = dtVM.Rows(0)
        Else
            rwVM = dtVM.NewRow
            rwVM(_VM.IDVino) = data.IDVino
            rwVM(_VM.IDArticulo) = data.IDArticulo
            rwVM(_VM.IDAlmacen) = data.IDAlmacen
            dtVM.Rows.Add(rwVM)
        End If
        If data.Precio > 0 Then
            rwVM(_VM.Cantidad) = data.Q
            rwVM(_VM.Fecha) = Date.Today
            rwVM(_VM.Precio) = data.Precio
        Else
            rwVM.Delete()
        End If
        BusinessHelper.UpdateTable(dtVM)
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _VDM
    Public Const Vendimia As String = "Vendimia"
    Public Const RdtoxHaT As String = "RdtoxHaT"
    Public Const RdtoxHaB As String = "RdtoxHaB"
    Public Const RdtoT As String = "RdtoT"
    Public Const RdtoB As String = "RdtoB"
    Public Const IDArticuloT As String = "IDArticuloT"
    Public Const IDArticuloB As String = "IDArticuloB"
    Public Const IDTarifaT As String = "IDTarifaT"
    Public Const IDTarifaB As String = "IDTarifaB"
    Public Const PrecioOrigenT As String = "PrecioOrigenT"
    Public Const PrecioOrigenB As String = "PrecioOrigenB"
    Public Const PrecioUvaT As String = "PrecioUvaT"
    Public Const PrecioUvaB As String = "PrecioUvaB"
    Public Const PrecioExcedenteT As String = "PrecioExcedenteT"
    Public Const PrecioExcedenteB As String = "PrecioExcedenteB"
End Class