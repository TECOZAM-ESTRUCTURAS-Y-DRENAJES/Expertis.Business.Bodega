

<Serializable()> _
Public Class DataExplosionVino
    Public IDVino As Guid
    Public IDTipoOperacion As String
    ' Public TipoOrigen As Origenes = -1

    Public VerTratamientos As Boolean
    Public VerCosteElaboracion As Boolean
    Public VerCosteEstanciaNave As Boolean
    Public VerCosteInicial As Boolean

    Public VerDetalleOperaciones As Boolean
    Public VerDetalleAnalisis As Boolean
    Public VerDetalleEntradasUva As Boolean
    Public VerDetalleEntradasVino As Boolean

    Public Fecha As Date
    Public ConsiderarNoTrazar As Boolean
    Public ConsiderarCosteInicial As Boolean
    Public IDQueryCoste As Integer

    Public Vinos As List(Of Vino)
    Public VinosEstructura As List(Of VinoEstructura)

    'Private Const FECHA_MIN As Date = "19000101"
    'Private Const FECHA_MAX As Date = New Date(3000, 1, 1)


    Public Sub New(ByVal IDVino As Guid, ByVal VerTratamientos As Boolean, ByVal VerCosteElaboracion As Boolean, ByVal VerCosteEstanciaNave As Boolean, ByVal VerCosteInicial As Boolean, _
                   ByVal VerDetalleOperaciones As Boolean, ByVal VerDetalleAnalisis As Boolean, ByVal VerDetalleEntradasUva As Boolean, ByVal VerDetalleEntradasVino As Boolean, _
                   Optional ByVal IDTipoOperacion As String = "", Optional ByVal ConsiderarNoTrazar As Boolean = False, _
                   Optional ByVal ConsiderarCosteInicial As Boolean = False, Optional ByVal Fecha As Date = cnMinDate)
        Me.IDVino = IDVino
        ' If TipoOrigen <> -1 Then Me.TipoOrigen = TipoOrigen
        If Length(IDTipoOperacion) > 0 Then Me.IDTipoOperacion = IDTipoOperacion

        Me.VerTratamientos = VerTratamientos
        Me.VerCosteElaboracion = VerCosteElaboracion
        Me.VerCosteEstanciaNave = VerCosteEstanciaNave
        Me.VerCosteInicial = VerCosteInicial

        Me.VerDetalleOperaciones = VerDetalleOperaciones
        Me.VerDetalleAnalisis = VerDetalleAnalisis
        Me.VerDetalleEntradasUva = VerDetalleEntradasUva
        Me.VerDetalleEntradasVino = VerDetalleEntradasVino

        If Fecha = cnMinDate Then
            Me.Fecha = New Date(3000, 1, 1)
        Else
            Me.Fecha = Fecha
        End If
    End Sub

    'Public Sub New(ByVal IDVino As Guid, ByVal TipoOrigen As Origenes)
    '    Me.IDVino = IDVino
    '    '   If TipoOrigen <> -1 Then Me.TipoOrigen = TipoOrigen
    '    If Me.Fecha = cnMinDate Then
    '        Me.Fecha = New Date(3000, 1, 1)
    '    End If
    'End Sub

    Public Sub New(ByVal IDVino As Guid, Optional ByVal Fecha As Date = cnMinDate, Optional ByVal ConsiderarNoTrazar As Boolean = False, Optional ByVal ConsiderarCosteInicial As Boolean = False)
        Me.IDVino = IDVino
        '  If TipoOrigen <> -1 Then Me.TipoOrigen = TipoOrigen
        If Length(IDTipoOperacion) > 0 Then Me.IDTipoOperacion = IDTipoOperacion

        Me.ConsiderarNoTrazar = ConsiderarNoTrazar
        Me.ConsiderarCosteInicial = ConsiderarCosteInicial

        If Fecha = cnMinDate Then
            Me.Fecha = New Date(3000, 1, 1)
        Else
            Me.Fecha = Fecha
        End If
    End Sub

    Public Sub New(ByVal IDVino As Guid, ByVal Vinos As List(Of Vino), ByVal VinosEstructura As List(Of VinoEstructura))
        Me.IDVino = IDVino
        Me.Vinos = Vinos
        If Not VinosEstructura Is Nothing AndAlso VinosEstructura.Count > 0 Then
            Me.VinosEstructura = VinosEstructura
        End If

        Me.Fecha = New Date(3000, 1, 1)
    End Sub


End Class
<Serializable()> _
Public Class DataExplosionVinoResult
    Public IDVino As Guid

    Public Vinos As List(Of Vino)
    Public VinosEstructura As List(Of VinoEstructura)
    Public Materiales As List(Of VinoTratamientos)
    Public VinosDetalle As List(Of VinoDetalle)

    Public Variedades As List(Of OrigenVariedades)
    Public Compras As List(Of OrigenCompras)
    Public Añadas As List(Of OrigenAniadas)
    Public Fincas As List(Of OrigenFincas)
    Public EntradasUVA As List(Of OrigenEntradasUVA)

    Public Sub New(ByVal IDVino As Guid)
        Me.IDVino = IDVino
    End Sub
End Class


<Serializable()> _
Public Class DataImplosionVino
    Public IDVino As Guid
    Public Fecha As Date
    Public VerEstructura As Boolean
    Public VerSalidas As Boolean
    Public VerStock As Boolean
    Public VerLotes As Boolean
    Public IDQueryCoste As Integer

    Public Vinos As List(Of Vino)
    Public VinosEstructura As List(Of VinoEstructura)

    Public Sub New(ByVal IDVino As Guid, Optional ByVal Fecha As Date = cnMinDate, Optional ByVal VerEstructura As Boolean = True, Optional ByVal VerSalidas As Boolean = True, Optional ByVal VerStock As Boolean = True, Optional ByVal VerLotes As Boolean = True)
        Me.IDVino = IDVino
        Me.VerEstructura = VerEstructura
        Me.VerSalidas = VerSalidas
        Me.VerStock = VerStock
        Me.VerLotes = VerLotes

        If Fecha = cnMinDate Then
            Me.Fecha = New Date(3000, 1, 1)
        Else
            Me.Fecha = Fecha
        End If
    End Sub

    Public Sub New(ByVal IDVino As Guid, ByVal Vinos As List(Of Vino), ByVal VinosEstructura As List(Of VinoEstructura), Optional ByVal VerEstructura As Boolean = True, Optional ByVal VerSalidas As Boolean = True, Optional ByVal VerStock As Boolean = True, Optional ByVal VerLotes As Boolean = True)
        Me.IDVino = IDVino
        Me.Vinos = Vinos
        If Not VinosEstructura Is Nothing AndAlso VinosEstructura.Count > 0 Then
            Me.VinosEstructura = VinosEstructura
        End If
        Me.VerEstructura = VerEstructura
        Me.VerSalidas = VerSalidas
        Me.VerStock = VerStock
        Me.VerLotes = VerLotes
        Me.Fecha = New Date(3000, 1, 1)
    End Sub


End Class


<Serializable()> _
Public Class DataImplosionVinoResult
    Public IDVino As Guid

    Public Vinos As List(Of Vino)
    Public VinosEstructura As List(Of VinoEstructura)

    Public TrazabilidadSalidas As List(Of VinoTrazabilidadSalidas)
    Public TrazabilidadStock As List(Of VinoTrazabilidadStock)
    Public TrazabilidadLotes As List(Of VinoTrazabilidadLotes)

    Public Sub New(ByVal IDVino As Guid)
        Me.IDVino = IDVino
    End Sub
End Class


<Serializable()> _
Public Class DataEstructuraVino
    Public Vinos As IList(Of Vino)
    Public VinosEstructura As IList(Of VinoEstructura)
    Public FechaFinCoste As Date

    Public Sub New(ByVal Vinos As IList(Of Vino), ByVal VinosEstructura As IList(Of VinoEstructura))
        Me.Vinos = Vinos
        Me.VinosEstructura = VinosEstructura
    End Sub
End Class


Public Class BdgExplosionVino

    Protected Shared TABLA_TMP_VINO As String = "tmpBdgVino"
    Protected Shared TABLA_TMP_VINO_COSTE_ESTANCIA As String = "tmpBdgVinoCosteEstancia"

    <Task()> Public Shared Function Explosion(ByVal data As DataExplosionVino, ByVal services As ServiceProvider) As DataExplosionVinoResult
        Dim dataContext As New WineDataContext(AdminData.GetConnectionString)
        Dim wineStruct As IMultipleResults = dataContext.GetWineStructure(data.IDVino, data.Fecha, data.VerTratamientos, data.VerCosteElaboracion, data.VerCosteEstanciaNave, data.VerCosteInicial, data.VerDetalleOperaciones, data.VerDetalleAnalisis, data.VerDetalleEntradasUva, data.VerDetalleEntradasVino, data.ConsiderarNoTrazar, data.ConsiderarCosteInicial, data.IDQueryCoste)
        If Not wineStruct Is Nothing Then
            Dim lstVinos As IList(Of Vino) = wineStruct.GetResult(Of Vino).ToList()
            Dim lstVinosEstructura As IList(Of VinoEstructura) = wineStruct.GetResult(Of VinoEstructura).ToList()
            Dim lstVinoMateriales As IList(Of VinoTratamientos) = wineStruct.GetResult(Of VinoTratamientos).ToList()
            Dim lstVinoDetalle As IList(Of VinoDetalle) = wineStruct.GetResult(Of VinoDetalle).ToList()

            Dim datCalPor As New DataEstructuraVino(lstVinos, lstVinosEstructura)
            datCalPor.FechaFinCoste = data.Fecha
            datCalPor = ProcessServer.ExecuteTask(Of DataEstructuraVino, DataEstructuraVino)(AddressOf DefinirVinculosEntreVinos, datCalPor, services)
            lstVinos = datCalPor.Vinos
            lstVinosEstructura = datCalPor.VinosEstructura

            Dim rslt As New DataExplosionVinoResult(data.IDVino)
            rslt.Vinos = lstVinos
            rslt.VinosEstructura = lstVinosEstructura
            rslt.Materiales = lstVinoMateriales
            rslt.VinosDetalle = lstVinoDetalle
            Return rslt
        End If

    End Function


    <Task()> Public Shared Function Implosion(ByVal data As DataImplosionVino, ByVal services As ServiceProvider) As DataImplosionVinoResult
        Dim dataContext As New WineDataContext(AdminData.GetConnectionString)
        Dim wineStruct As IMultipleResults = dataContext.GetWineStructureImplosion(data.IDVino, data.Fecha, data.VerEstructura, data.VerSalidas, data.VerStock, data.VerLotes, data.IDQueryCoste)
        If Not wineStruct Is Nothing Then
            Dim lstVinos As IList(Of Vino) = wineStruct.GetResult(Of Vino).ToList()
            Dim lstVinosEstructura As IList(Of VinoEstructura) = wineStruct.GetResult(Of VinoEstructura).ToList()

            Dim datCalPor As New DataEstructuraVino(lstVinos, lstVinosEstructura)
            datCalPor = ProcessServer.ExecuteTask(Of DataEstructuraVino, DataEstructuraVino)(AddressOf DefinirVinculosEntreVinos, datCalPor, services)
            lstVinos = datCalPor.Vinos
            lstVinosEstructura = datCalPor.VinosEstructura

            Dim rslt As New DataImplosionVinoResult(data.IDVino)
            rslt.Vinos = lstVinos
            rslt.VinosEstructura = lstVinosEstructura
            rslt.TrazabilidadSalidas = wineStruct.GetResult(Of VinoTrazabilidadSalidas).ToList()
            rslt.TrazabilidadStock = wineStruct.GetResult(Of VinoTrazabilidadStock).ToList()
            rslt.TrazabilidadLotes = wineStruct.GetResult(Of VinoTrazabilidadLotes).ToList()

            Return rslt
        End If

    End Function



    <Task()> Public Shared Function OrigenesVino(ByVal data As DataExplosionVino, ByVal services As ServiceProvider) As DataExplosionVinoResult
        Dim rslt As DataExplosionVinoResult
        If data.Vinos Is Nothing OrElse data.VinosEstructura Is Nothing OrElse data.Vinos.Count = 0 OrElse data.VinosEstructura.Count = 0 Then
            rslt = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf Explosion, data, services)
        Else
            If rslt Is Nothing Then rslt = New DataExplosionVinoResult(data.IDVino)
            rslt.Vinos = data.Vinos
            rslt.VinosEstructura = data.VinosEstructura
        End If

        If Not rslt Is Nothing AndAlso Not rslt.Vinos Is Nothing AndAlso Not rslt.VinosEstructura Is Nothing Then
            Dim lstVinos As IList(Of Vino) = rslt.Vinos
            Dim lstVinosEstructura As IList(Of VinoEstructura) = rslt.VinosEstructura

            Dim QDepositoVino As Double = Nz(lstVinos(0).QDepositoVino, 0)
            If QDepositoVino = 0 Then QDepositoVino = Nz(lstVinos(0).QTotal, 0)

            Dim datPorcTmp As New DataRellenarTablaPorcentajes(lstVinos, lstVinosEstructura, True, False)
            Dim queryId As Integer = ProcessServer.ExecuteTask(Of DataRellenarTablaPorcentajes, Integer)(AddressOf RellenarTablaPorcentajes, datPorcTmp, services)

            '//Recuperamos los distintos origenes a la vez (Variedad, Fincas, Compras, Añadas) y vaciamos la tabla temporal
            Dim dataContext As New WineDataContext(AdminData.GetConnectionString)
            Dim origen As IMultipleResults = dataContext.VinoOrigen(queryId, QDepositoVino)
            If Not origen Is Nothing Then
                ' Dim rslt As New DataExplosionVinoResult
                rslt.Variedades = origen.GetResult(Of OrigenVariedades).ToList
                rslt.Fincas = origen.GetResult(Of OrigenFincas).ToList
                rslt.Compras = origen.GetResult(Of OrigenCompras).ToList
                rslt.Añadas = origen.GetResult(Of OrigenAniadas).ToList
                rslt.EntradasUVA = origen.GetResult(Of OrigenEntradasUVA).ToList
            End If
            ProcessServer.ExecuteTask(Of Integer)(AddressOf DeleteTablaPorcentajes, queryId, services)
            Return rslt

            '///
        End If
    End Function

    <Serializable()> _
    Public Class DataRellenarTablaPorcentajes
        Public lstVinos As IList(Of Vino)
        Public lstVinosEstructura As IList(Of VinoEstructura)
        Public PorcentajeOrigenes As Boolean
        Public PorcentajeCostes As Boolean

        Public Sub New(ByVal lstVinos As IList(Of Vino), ByVal lstVinosEstructura As IList(Of VinoEstructura), Optional ByVal PorcentajeOrigenes As Boolean = False, Optional ByVal PorcentajeCostes As Boolean = False)
            Me.lstVinos = lstVinos
            Me.lstVinosEstructura = lstVinosEstructura
            Me.PorcentajeOrigenes = PorcentajeOrigenes
            Me.PorcentajeCostes = PorcentajeCostes
        End Sub
    End Class
    <Task()> Public Shared Function RellenarTablaPorcentajes(ByVal data As DataRellenarTablaPorcentajes, ByVal services As ServiceProvider) As Integer
        Dim QDepositoVino As Double = Nz(data.lstVinos(0).QDepositoVino, 0)

        '//Rellenamos tabla temporal
        Dim tmpTable As DataTable = New BE.DataEngine().Filter(TABLA_TMP_VINO, New NoRowsFilterItem)
        Dim queryId As Integer = New Random().[Next](Integer.MaxValue)
        Dim tmpTableVE As DataTable = New BE.DataEngine().Filter(TABLA_TMP_VINO_COSTE_ESTANCIA, New NoRowsFilterItem)

        For Each v As Vino In data.lstVinos
            Dim PorcentajeAux As Double = 0
            Dim Porcentaje As Decimal
            Try
                If data.PorcentajeOrigenes Then
                    PorcentajeAux = v.PorcentajeOrigen
                ElseIf data.PorcentajeCostes Then
                    PorcentajeAux = v.PorcentajeCoste
                End If
                Porcentaje = CDec(PorcentajeAux / 100)

                '//PENDIENTE REVISAR
                Dim NumDecimales As Integer
                Dim partedecimal As Double = Porcentaje - CInt(Porcentaje)
                If Length(partedecimal.ToString) > 23 Then
                    Porcentaje = 0
                End If

            Catch ex As OverflowException
                Porcentaje = 0
            End Try
            tmpTable.Rows.Add(queryId, v.IDVino, Porcentaje, v.QTotal)

            '//Rellenamos tabla temporal Vino Estructura 
            For Each cst As VinoCoste In v.CosteVinoFecha.Values
                tmpTableVE.Rows.Add(queryId, v.IDVino, cst.CantidadCoste, cst.FechaFinCoste)
            Next
        Next

        If tmpTable.Rows.Count > 0 Then
            Dim cnn As Common.DbConnection = AdminData.GetSessionConnection.Connection
            Dim bulk As New System.Data.SqlClient.SqlBulkCopy(cnn)
            bulk.DestinationTableName = TABLA_TMP_VINO
            bulk.BulkCopyTimeout = 0
            bulk.WriteToServer(tmpTable)
        End If

        If tmpTableVE.Rows.Count > 0 Then
            Dim cnn As Common.DbConnection = AdminData.GetSessionConnection.Connection
            Dim bulk As New System.Data.SqlClient.SqlBulkCopy(cnn)
            bulk.DestinationTableName = TABLA_TMP_VINO_COSTE_ESTANCIA
            bulk.BulkCopyTimeout = 0
            bulk.WriteToServer(tmpTableVE)
        End If

        Return queryId
    End Function

    <Task()> Public Shared Sub DeleteTablaPorcentajes(ByVal queryId As Integer, ByVal services As ServiceProvider)
        AdminData.Execute("DELETE FROM tmpBdgVino WHERE ID = " & queryId)
        AdminData.Execute("DELETE FROM tmpBdgVinoCosteEstancia WHERE ID = " & queryId)
    End Sub

#Region " Establecemos vinculos entre un vino y sus componentes y calculamos porcentajes "

    <Task()> Public Shared Function DefinirVinculosEntreVinos(ByVal data As DataEstructuraVino, ByVal services As ServiceProvider) As DataEstructuraVino
        '//En esta función, datos los vinos del articulo y sus correspondientes enlaces, se establecen los vinculos entre ellos y se definen los niveles y porcentajes.
        Dim kk As Boolean
        Dim Orden As Integer = 0
        For Each v As Vino In data.Vinos
            v.Orden = Orden
            Orden = Orden + 1
        Next
        For Each v As Vino In data.Vinos
            '//VER CON MAP
            ' If v.QTotal = 0 AndAlso v.QDepositoVino <> 0 Then v.QTotal = v.QDepositoVino
            Dim links As List(Of VinoEstructura) = (From link In data.VinosEstructura Where link.IDVino = v.IDVino Select link).ToList
            If Not links Is Nothing AndAlso links.Count <> 0 Then
                For Each link As VinoEstructura In links
                    link.Parent = v
                    Dim VinoComponente As List(Of Vino) = (From vAux In data.Vinos Where vAux.IDVino = link.IDVinoComponente).ToList
                    If Not VinoComponente Is Nothing AndAlso VinoComponente.Count > 0 Then
                        link.Child = VinoComponente(0)
                    End If
                    v.Links.Add(link)
                Next
            Else
                v.IsLeaf = True
            End If
        Next

        data = ProcessServer.ExecuteTask(Of DataEstructuraVino, DataEstructuraVino)(AddressOf CalcularPorcentajesYCantidadCoste, data, services)
        Return data
    End Function

    Public Shared TestStack As New Stack(Of Vino)
    Public Shared ActivarTest As Boolean = False    '//Poner a True para poder ver si hay vinos buclados
    <Task()> Public Shared Function CalcularPorcentajesYCantidadCoste(ByVal data As DataEstructuraVino, ByVal services As ServiceProvider) As DataEstructuraVino

        '//Se calculan los costes de los acumulados
        Dim v As Vino = data.Vinos(0)
        If ActivarTest Then TestStack = New Stack(Of Vino)
        v.Level = 0
        v.PorcentajeOrigen = 100
        v.PorcentajeCoste = 100

        Dim CosteEnFechaPadreRaiz As New VinoCoste
        CosteEnFechaPadreRaiz.CantidadCoste = v.QTotal
        If data.FechaFinCoste <> cnMinDate Then
            CosteEnFechaPadreRaiz.FechaFinCoste = data.FechaFinCoste
        Else
            CosteEnFechaPadreRaiz.FechaFinCoste = Today
        End If
        If v.CosteVinoFecha Is Nothing Then v.CosteVinoFecha = New Dictionary(Of Date, VinoCoste)
        v.CosteVinoFecha.Add(CosteEnFechaPadreRaiz.FechaFinCoste, CosteEnFechaPadreRaiz)

        Dim FactoresConversion As New Dictionary(Of String, Double)
        Dim LinksOrdenados As List(Of VinoEstructura) = (From link In data.VinosEstructura Order By link.Child.Level Select link).ToList
        For Each link As VinoEstructura In LinksOrdenados
            ''//
            Dim prctgOrigen As Double = 0
            If (link.Parent.QTotal) <> 0 Then
                Dim Factor As Double = 1
                If link.Parent.IDUdMedida <> link.Child.IDUdMedida Then
                    Dim keyFactor As String = UCase(link.Parent.IDUdMedida) & "/" & UCase(link.Child.IDUdMedida)
                    If FactoresConversion.ContainsKey(keyFactor) Then
                        Factor = FactoresConversion(keyFactor)
                    Else
                        Dim datFactor As New UnidadAB.UnidadMedidaInfo
                        datFactor.IDUdMedidaA = link.Parent.IDUdMedida
                        datFactor.IDUdMedidaB = link.Child.IDUdMedida
                        datFactor.UnoSiNoExiste = True
                        Factor = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, datFactor, services)
                        FactoresConversion(keyFactor) = Factor
                    End If
                End If
                prctgOrigen = ((link.Cantidad + If(link.Merma < 0, -link.Merma, 0)) / (link.Parent.QTotal * Factor)) * (link.Parent.PorcentajeOrigen / 100)
            End If

            ''//
            Dim prctgCoste As Double = 0
            If (link.Child.QTotal - link.Child.QTotMermaCoste) <> 0 Then
                Dim dblCreces As Double = 0
                If link.Merma < 0 AndAlso link.CrecesAcumulanCoste Then
                    dblCreces = If(link.Merma < 0, -link.Merma, 0)
                End If
                If Not link.UsarRepartoTipoOperacion Then
                    prctgCoste = ((link.Cantidad + dblCreces) / (link.Child.QTotal - link.Child.QTotMermaCoste)) * (link.Parent.PorcentajeCoste / 100)
                Else
                    prctgCoste = ((link.CantidadRepartoTipoOperacion * link.PorcentajeRepartoTipoOperacion / 100) / (link.Child.QTotal - link.Child.QTotMermaCoste)) * (link.Parent.PorcentajeCoste / 100)
                End If
            End If
            link.Child.PorcentajeOrigen += prctgOrigen * 100
            link.Child.PorcentajeCoste += prctgCoste * 100

            ''//
            Dim FechaPadre As Date = link.Parent.Fecha
            Dim CosteEnFechaPadre As New VinoCoste
            If link.Parent.CosteVinoFecha Is Nothing Then link.Parent.CosteVinoFecha = New Dictionary(Of Date, VinoCoste)
            For Each cst As VinoCoste In link.Parent.CosteVinoFecha.Values
                CosteEnFechaPadre.CantidadCoste += cst.CantidadCoste
            Next

            Dim CantidadCoste As Double = 0
            Dim NumeroHermanos As Double = (Aggregate lnk In data.VinosEstructura Where lnk.IDVino = link.Parent.IDVino Into Count())
            If NumeroHermanos = 1 Then
                If CosteEnFechaPadre.CantidadCoste / link.Factor < link.Cantidad Then
                    CantidadCoste = CosteEnFechaPadre.CantidadCoste / link.Factor
                Else
                    CantidadCoste = link.Cantidad
                End If
            Else
                Dim CantidadComponentes As Double = (Aggregate lnk In data.VinosEstructura Where lnk.IDVino = link.Parent.IDVino Into Sum(lnk.Cantidad))
                If CantidadComponentes <> 0 Then
                    CantidadCoste = (CosteEnFechaPadre.CantidadCoste / link.Factor) * (link.Cantidad / CantidadComponentes)
                Else
                    CantidadCoste = 0
                End If
            End If

            Dim CosteEnFecha As VinoCoste
            If link.Child.CosteVinoFecha Is Nothing Then link.Child.CosteVinoFecha = New Dictionary(Of Date, VinoCoste)
            If link.Child.CosteVinoFecha.Keys.Contains(FechaPadre) Then
                CosteEnFecha = link.Child.CosteVinoFecha(FechaPadre)
                CosteEnFecha.CantidadCoste += CantidadCoste
            Else
                CosteEnFecha = New VinoCoste
                CosteEnFecha.FechaFinCoste = FechaPadre
                CosteEnFecha.CantidadCoste = CantidadCoste
                link.Child.CosteVinoFecha.Add(CosteEnFecha.FechaFinCoste, CosteEnFecha)
            End If
        Next
        Return data
    End Function
#End Region

End Class

<Serializable()> _
Public Class DataOpcionesTraza
    Public VerTratamientos As Boolean
    Public VerCosteElaboracion As Boolean
    Public VerCosteEstanciaNave As Boolean
    Public VerCosteInicial As Boolean
    Public NivelMaximo As Integer
    Public NivelMaximoTraza As Integer
    'Public Trazabilidad As Boolean
    Public VerDetalleOperaciones As Boolean
    Public VerDetalleAnalisis As Boolean
    Public VerDetalleEntradasUva As Boolean
    Public VerDetalleEntradasVino As Boolean
    Public VerDetalleSalidasAV As Boolean
    Public VerDetalleOtrasEntradas As Boolean
    Public VerDetalleAjustes As Boolean
    Public VerDetalleOfs As Boolean
    Public VerDetalleTransferencias As Boolean
    Public VerDetalleInventarios As Boolean
    Public VerDetalleOtrasSalidas As Boolean
    Public VerPorcentajes As Boolean
    Public SoloVinosEnExistencias As Boolean

    Public Sub New()
        Me.VerTratamientos = False
        Me.VerCosteElaboracion = False
        Me.VerCosteEstanciaNave = False
        Me.VerCosteInicial = False
        Me.NivelMaximo = 100
        Me.NivelMaximoTraza = Me.NivelMaximo
        'Me.Trazabilidad = False
        Me.VerDetalleOperaciones = True
        Me.VerDetalleAnalisis = True
        Me.VerDetalleEntradasUva = True
        Me.VerDetalleEntradasVino = True
        Me.VerDetalleSalidasAV = True
        Me.VerDetalleOtrasEntradas = True
        Me.VerDetalleAjustes = True
        Me.VerDetalleOfs = True
        Me.VerDetalleTransferencias = True
        Me.VerDetalleInventarios = True
        Me.VerDetalleOtrasSalidas = True
        Me.SoloVinosEnExistencias = False
    End Sub

    Public Sub New(ByVal VerTratamientos As Boolean, ByVal VerCosteElaboracion As Boolean, ByVal VerCosteEstanciaNave As Boolean, _
                   ByVal VerCosteInicial As Boolean, ByVal NivelMaximo As Integer, ByVal NivelMaximoTraza As Integer, _
                   ByVal Trazabilidad As Boolean, ByVal VerDetalleOperaciones As Boolean, ByVal VerDetalleAnalisis As Boolean, _
                   ByVal VerDetalleEntradasUva As Boolean, ByVal VerDetalleEntradasVino As Boolean, ByVal VerDetalleSalidasAV As Boolean, _
                   ByVal VerDetalleOtrasEntradas As Boolean, ByVal VerDetalleAjustes As Boolean, ByVal VerDetalleOfs As Boolean, _
                   ByVal VerDetalleTransferencias As Boolean, ByVal VerDetalleInventarios As Boolean, ByVal VerDetalleOtrasSalidas As Boolean, _
                   ByVal SoloVinosEnExistencias As Boolean)
        Me.VerTratamientos = VerTratamientos
        Me.VerCosteElaboracion = VerCosteElaboracion
        Me.VerCosteEstanciaNave = VerCosteEstanciaNave
        Me.VerCosteInicial = VerCosteInicial
        Me.NivelMaximo = NivelMaximo
        Me.NivelMaximoTraza = NivelMaximoTraza
        'Me.Trazabilidad = Trazabilidad
        Me.VerDetalleOperaciones = VerDetalleOperaciones
        Me.VerDetalleAnalisis = VerDetalleAnalisis
        Me.VerDetalleEntradasUva = VerDetalleEntradasUva
        Me.VerDetalleEntradasVino = VerDetalleEntradasVino
        Me.VerDetalleSalidasAV = VerDetalleSalidasAV
        Me.VerDetalleOtrasEntradas = VerDetalleOtrasEntradas
        Me.VerDetalleAjustes = VerDetalleAjustes
        Me.VerDetalleOfs = VerDetalleOfs
        Me.VerDetalleTransferencias = VerDetalleTransferencias
        Me.VerDetalleInventarios = VerDetalleInventarios
        Me.VerDetalleOtrasSalidas = VerDetalleOtrasSalidas
        Me.SoloVinosEnExistencias = SoloVinosEnExistencias
    End Sub
End Class
