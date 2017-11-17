'///////////////////////////////////////////////////////////////////////
'//                                                                   //
'// WEB SERVICE GENERIC DB V.1.5                                      //
'//                                                                   //
'// AUTORES:    IVAN SALDIVAR RODRIGUEZ (R)2010-2012                  //                   
'//             JOSE CORTÉS ARANEDA     (R)2010-2012                  //
'//                                                                   //
'// TODOS LOS DERECHOS DE COPIA ESTÁN RESERVADOS A:                   //
'//                                                                   //
'//           * IVAN SALDIVAR RODRIGUEZ                               //
'//           * JOSE CORTÉS ARANEDA                                   //
'//                                                                   //
'// Y SU USO NO AUTORIZADO SERÁ SANCIONADO DE ACUERDO A LAS LEYES DE  //
'// DERECHO  DE AUTOR VIGENTES EN EL PAÍS DONDE SE HAYA COMETIDO LA   //
'// FALTA.                                                            //
'// SE PROHIBE LA COPIA, REPRODUCCIÓN Y DISTRIBUCIÓN CON FINES        //
'// COMERCIALES  SIN PREVIA AUTORIZACIÓN. LOS ALGORITMO QUE           //
'// COMPONEN ESTE SOFTWARE PROPIEDAD DEL AUTOR, POR LO TANTO, ESTOS   //
'// FORMA PARTE INTEGRA DEL CÓDIGO EN EL CUAL SE HAN IMPLEMENTADO.    //
'//                                                                   //
'// SANTIAGO - CHILE, 2012-01-01                                      //       
'//                                                                   //
'///////////////////////////////////////////////////////////////////////

Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System

Public Class WS_Generic_DB

#Region "METODOS GENERICOS"

    Public Function CONSULTA_GENERICA_SP(ByVal NOMBRESP As String, _
                                         ByVal XMLParams As String, _
                                         ByVal vConnectionString As String) As Data.DataSet
        Dim Cmd As New SqlCommand
        Dim Resultado As Data.DataSet = New Data.DataSet()
        Dim Serv As New Class_Servidor

        'ISR(20110517)
        'OBJETIVO:  CONSULTA PERMITE EJECUTAR EN FORMA GENÉRICA CUALQUIER SP.
        '           PASANDO SU NOMBRE Y PARÁMETROS
        Try
            Serv.Conectar(vConnectionString)
            With Cmd
                .CommandText = NOMBRESP
                .CommandType = Data.CommandType.StoredProcedure
                .Connection = Serv.Conec
            End With
            If XMLParams.Trim <> "" Then
                CargaParametrosSP(Cmd, XMLParams)
            End If

            Dim SqlAdapter As New SqlDataAdapter
            SqlAdapter.SelectCommand = Cmd

            SqlAdapter.Fill(Resultado, "CONSULTA_GENERICA_SP")

            Serv.Desconectar()

        Catch ex As Exception
            Serv.Desconectar()
            Resultado = Nothing
        End Try


        Return Resultado


    End Function
    Public Function CONSULTA_GENERICA_SP2(ByVal NOMBRESP As String, _
                                          ByVal XMLParams As String, _
                                          ByVal vConnectionString As String) As String
        Dim Cmd As New SqlCommand
        Dim Resultado As Data.DataSet = New Data.DataSet()
        Dim salida As String
        Dim Serv As New Class_Servidor
        salida = ""

        'ISR(20110517)
        'OBJETIVO:  CONSULTA PERMITE EJECUTAR EN FORMA GENÉRICA CUALQUIER SP.
        '           PASANDO SU NOMBRE Y PARÁMETROS. ESTÁ DISEÑADA ESPECIALMENTE
        '           PARA EJECUTAR PROCEDIMIENTOS QUE REALIZAN OPERACIONES DE 
        '           ACTUALIZACIÓN EN LA BASE DE DATOS.
        Try
            Serv.Conectar(vConnectionString)
            With Cmd
                .CommandText = NOMBRESP
                .CommandType = Data.CommandType.StoredProcedure
                .Connection = Serv.Conec
            End With
            If XMLParams.Trim <> "" Then
                CargaParametrosSP(Cmd, XMLParams)
            End If

            Dim SqlAdapter As New SqlDataAdapter
            SqlAdapter.SelectCommand = Cmd

            SqlAdapter.Fill(Resultado, "CONSULTA_GENERICA_SP2")

            Serv.Desconectar()

            If Resultado.Tables(0).Rows.Count > 0 Then
                If Resultado.Tables(0).Rows(0).Item(0).ToString <> "-1" Then
                    salida = "OK" & " - " & Resultado.Tables(0).Rows(0).Item(0).ToString
                Else
                    For Each FILA As DataRow In Resultado.Tables(0).Rows
                        salida = "Procedimiento: " & FILA("ERROR_PROCEDURE_") & Chr(13) & _
                                 "Código error: " & FILA("ERROR_NUMBER_") & Chr(13) & _
                                 "Descripción: " & FILA("ERROR_MESSAGE_") & Chr(13) & _
                                 "Número línea: " & FILA("ERROR_LINE_")
                    Next
                End If
            End If

        Catch ex As Exception
            Serv.Desconectar()
            Resultado = Nothing
            salida = ex.Message
        End Try

        Return salida

    End Function
    Public Function CONSULTA_GENERICA_SP3(ByVal NOMBRESP As String, _
                                             ByVal XMLParams As String, _
                                             ByVal XMLDetalle As String, _
                                             ByVal vConnectionString As String) As String
        Dim Cmd As New SqlCommand
        Dim Resultado As Data.DataSet = New Data.DataSet()
        Dim salida As String
        Dim Serv As New Class_Servidor
        Dim oDOM As New System.Xml.XmlDocument

        salida = ""

        'ISR(20110629)
        'OBJETIVO:  CONSULTA PERMITE EJECUTAR EN FORMA GENÉRICA CUALQUIER SP.
        '           PASANDO SU NOMBRE Y PARÁMETROS. ESTÁ DISEÑADA ESPECIALMENTE
        '           PARA EJECUTAR PROCEDIMIENTOS QUE REALIZAN OPERACIONES DE 
        '           ACTUALIZACIÓN EN LA BASE DE DATOS.
        '           A DIFERENCIA DE CONSULTA_GENERICA_SP2, ESTE WEB METHOD CONSIDERA
        '           UN NUEVO PARÁMETRO (XMLDetalles) EN EL CUAL SE INCLUYEN LOS VALORES 
        '           PARA GRABAR EN BASE DE DATOS VARIAS INSTANCIAS REPETIDAS DE UN TIPO
        '           DE REGISTRO "DETALLE", POR EJEMPLO: UNA RECETA Y SU DETALLE, EN 
        '           ESTE CASO EL DETALLE DE LA RECETA SE FORMATEA A UN XML QUE
        '           SERÁ PROCESADO INTERNAMENTE EN EL PROCEDIMIENTO INVOCADO, DE ESTE
        '           MODO SE ENVÍA DE UNA SOLA VEZ A LA CAPA DE DATOS LAS INSTANCIAS 
        '           REPETIDAS ASOCIADAS A UNA CABECERA
        Try
            Serv.Conectar(vConnectionString)
            With Cmd
                .CommandText = NOMBRESP
                .CommandType = Data.CommandType.StoredProcedure
                .Connection = Serv.Conec
            End With

            'AQUÍ AGREGAMOS LOS PARAMETROS ESCALARES BASICOS
            If XMLParams.Trim <> "" Then
                CargaParametrosSP(Cmd, XMLParams)
            End If
            oDOM.LoadXml(XMLDetalle)

            'AQUÍ AGREGAMOS EL PARÁMETRO ASOCIADO AL DETALLE DE LA INSTANCIA PRIMARIA
            Cmd.Parameters.AddWithValue("@DETALLE_XML", oDOM.InnerXml)

            Dim SqlAdapter As New SqlDataAdapter
            SqlAdapter.SelectCommand = Cmd

            SqlAdapter.Fill(Resultado, "CONSULTA_GENERICA_SP3")

            Serv.Desconectar()

            If Resultado.Tables(0).Rows.Count > 0 Then
                If Resultado.Tables(0).Rows(0).Item(0).ToString <> "-1" Then
                    salida = "OK" & " - " & Resultado.Tables(0).Rows(0).Item(0).ToString
                Else
                    For Each FILA As DataRow In Resultado.Tables(0).Rows
                        salida = "Procedimiento: " & FILA("ERROR_PROCEDURE_") & Chr(13) & _
                                 "Código error: " & FILA("ERROR_NUMBER_") & Chr(13) & _
                                 "Descripción: " & FILA("ERROR_MESSAGE_") & Chr(13) & _
                                 "Número línea: " & FILA("ERROR_LINE_")
                    Next
                End If
            End If

        Catch ex As Exception
            Serv.Desconectar()
            Resultado = Nothing
            salida = ex.Message
        End Try

        Return salida

    End Function

    Public Function CONSULTA_GENERICA_SP_SERIALIZADA(ByVal NOMBRESP As Object, _
                                                     ByVal XMLParams As Object, _
                                                     ByVal vConnectionString As String) As Byte()
        Dim Cmd As New SqlCommand
        Dim Resultado As Data.DataSet = New Data.DataSet()
        Dim Serv As New Class_Servidor
        Dim XMLParams_str As String = ""
        Dim msgExcepcion As String = ""
        Dim objSerializado As Byte() = Nothing

        'ISR(20110517)
        'OBJETIVO:  CONSULTA PERMITE EJECUTAR EN FORMA GENÉRICA CUALQUIER SP.
        '           PASANDO SU NOMBRE Y PARÁMETROS
        Try
            Serv.Conectar(vConnectionString)
            With Cmd
                .CommandText = Deserialize(NOMBRESP)
                .CommandType = Data.CommandType.StoredProcedure
                .Connection = Serv.Conec
            End With

            XMLParams_str = Deserialize(XMLParams)
            If XMLParams_str.Trim <> "" Then
                CargaParametrosSP(Cmd, XMLParams_str)
            End If

            Dim SqlAdapter As New SqlDataAdapter
            SqlAdapter.SelectCommand = Cmd

            SqlAdapter.Fill(Resultado, "CONSULTA_GENERICA_SP_SERIALIZADA")

            Serv.Desconectar()

            If Resultado.Tables(0).Rows.Count > 0 Then
                If Resultado.Tables(0).Rows(0).Item(0).ToString <> "-1" Then

                    'SALIDA CON RETORNO DE DATOS
                    objSerializado = Serialize(Resultado, True)

                Else
                    For Each FILA As DataRow In Resultado.Tables(0).Rows
                        msgExcepcion = "Procedimiento: " & FILA("ERROR_PROCEDURE_") & Chr(13) & _
                                       "Código error: " & FILA("ERROR_NUMBER_") & Chr(13) & _
                                       "Descripción: " & FILA("ERROR_MESSAGE_") & Chr(13) & _
                                       "Número línea: " & FILA("ERROR_LINE_")
                    Next

                    '(1) SALIDA CON ERROR
                    objSerializado = Serialize(msgExcepcion, True)
                End If
            End If



        Catch ex As Exception
            Serv.Desconectar()
            Resultado = Nothing

            msgExcepcion = "Procedimiento: " & Cmd.CommandText() & Chr(13) & _
                           "Código error:" & Chr(13) & _
                           "Descripción:" & ex.Message & Chr(13) & _
                           "Número línea: "

            '(2) SALIDA CON ERROR
            objSerializado = Serialize(msgExcepcion, True)
        End Try

        Return objSerializado

    End Function
    Public Function CONSULTA_GENERICA_SP2_SERIALIZADA(ByVal NOMBRESP As Object, _
                                                      ByVal XMLParams As Object, _
                                                      ByVal vConnectionString As String) As Byte()
        Dim Cmd As New SqlCommand
        Dim Resultado As Data.DataSet = New Data.DataSet()
        Dim salida As String = ""
        Dim Serv As New Class_Servidor
        Dim XMLParams_str As String = ""


        'ISR(20110517)
        'OBJETIVO:  CONSULTA PERMITE EJECUTAR EN FORMA GENÉRICA CUALQUIER SP.
        '           PASANDO SU NOMBRE Y PARÁMETROS. ESTÁ DISEÑADA ESPECIALMENTE
        '           PARA EJECUTAR PROCEDIMIENTOS QUE REALIZAN OPERACIONES DE 
        '           ACTUALIZACIÓN EN LA BASE DE DATOS.
        Try
            Serv.Conectar(vConnectionString)
            With Cmd
                .CommandText = Deserialize(NOMBRESP)
                .CommandType = Data.CommandType.StoredProcedure
                .Connection = Serv.Conec
            End With

            XMLParams_str = Deserialize(XMLParams)
            If XMLParams_str.Trim <> "" Then
                CargaParametrosSP(Cmd, XMLParams_str)
            End If

            Dim SqlAdapter As New SqlDataAdapter
            SqlAdapter.SelectCommand = Cmd

            SqlAdapter.Fill(Resultado, "CONSULTA_GENERICA_SP2_SERIALIZADA")

            Serv.Desconectar()

            If Resultado.Tables(0).Rows.Count > 0 Then
                If Resultado.Tables(0).Rows(0).Item(0).ToString <> "-1" Then
                    salida = "OK" & " - " & Resultado.Tables(0).Rows(0).Item(0).ToString
                Else
                    For Each FILA As DataRow In Resultado.Tables(0).Rows
                        salida = "Procedimiento: " & FILA("ERROR_PROCEDURE_") & Chr(13) & _
                                 "Código error: " & FILA("ERROR_NUMBER_") & Chr(13) & _
                                 "Descripción: " & FILA("ERROR_MESSAGE_") & Chr(13) & _
                                 "Número línea: " & FILA("ERROR_LINE_")
                    Next
                End If
            End If

        Catch ex As Exception
            Serv.Desconectar()
            Resultado = Nothing
            salida = ex.Message
        End Try

        Return Serialize(salida, True)

    End Function
    Public Function CONSULTA_GENERICA_SP3_SERIALIZADA(ByVal NOMBRESP As Object, _
                                                     ByVal XMLParams As Object, _
                                                     ByVal XMLDetalle As Object, _
                                                     ByVal vConnectionString As String) As Byte()
        Dim Cmd As New SqlCommand
        Dim Resultado As Data.DataSet = New Data.DataSet()
        Dim salida As String
        Dim Serv As New Class_Servidor
        Dim oDOM As New System.Xml.XmlDocument
        Dim XMLParam_str As String = ""
        Dim XMLDetalle_str As String = ""

        salida = ""

        'ISR(20110629)
        'OBJETIVO:  CONSULTA PERMITE EJECUTAR EN FORMA GENÉRICA CUALQUIER SP.
        '           PASANDO SU NOMBRE Y PARÁMETROS. ESTÁ DISEÑADA ESPECIALMENTE
        '           PARA EJECUTAR PROCEDIMIENTOS QUE REALIZAN OPERACIONES DE 
        '           ACTUALIZACIÓN EN LA BASE DE DATOS.
        '           A DIFERENCIA DE CONSULTA_GENERICA_SP2, ESTE WEB METHOD CONSIDERA
        '           UN NUEVO PARÁMETRO (XMLDetalles) EN EL CUAL SE INCLUYEN LOS VALORES 
        '           PARA GRABAR EN BASE DE DATOS VARIAS INSTANCIAS REPETIDAS DE UN TIPO
        '           DE REGISTRO "DETALLE", POR EJEMPLO: UNA RECETA Y SU DETALLE, EN 
        '           ESTE CASO EL DETALLE DE LA RECETA SE FORMATEA A UN XML QUE
        '           SERÁ PROCESADO INTERNAMENTE EN EL PROCEDIMIENTO INVOCADO, DE ESTE
        '           MODO SE ENVÍA DE UNA SOLA VEZ A LA CAPA DE DATOS LAS INSTANCIAS 
        '           REPETIDAS ASOCIADAS A UNA CABECERA
        Try
            Serv.Conectar(vConnectionString)
            With Cmd
                .CommandText = Deserialize(NOMBRESP)
                .CommandType = Data.CommandType.StoredProcedure
                .Connection = Serv.Conec
            End With

            'AQUÍ AGREGAMOS LOS PARAMETROS ESCALARES BASICOS
            XMLParam_str = Deserialize(XMLParams)
            If XMLParam_str.Trim <> "" Then
                CargaParametrosSP(Cmd, XMLParam_str)
            End If

            XMLDetalle_str = Deserialize(XMLDetalle)
            If XMLDetalle_str.Trim <> "" Then oDOM.LoadXml(XMLDetalle_str)

            'AQUÍ AGREGAMOS EL PARÁMETRO ASOCIADO AL DETALLE DE LA INSTANCIA PRIMARIA
            If XMLDetalle_str.Trim <> "" Then Cmd.Parameters.AddWithValue("@DETALLE_XML", oDOM.InnerXml)

            Dim SqlAdapter As New SqlDataAdapter
            SqlAdapter.SelectCommand = Cmd

            SqlAdapter.Fill(Resultado, "CONSULTA_GENERICA_SP3_SERIALIZADA")

            Serv.Desconectar()

            If Resultado.Tables(0).Rows.Count > 0 Then
                If Resultado.Tables(0).Rows(0).Item(0).ToString <> "-1" Then
                    salida = "OK" & " - " & Resultado.Tables(0).Rows(0).Item(0).ToString
                Else
                    For Each FILA As DataRow In Resultado.Tables(0).Rows
                        salida = "Procedimiento: " & FILA("ERROR_PROCEDURE_") & Chr(13) & _
                                 "Código error: " & FILA("ERROR_NUMBER_") & Chr(13) & _
                                 "Descripción: " & FILA("ERROR_MESSAGE_") & Chr(13) & _
                                 "Número línea: " & FILA("ERROR_LINE_")
                    Next
                End If
            End If

        Catch ex As Exception
            Serv.Desconectar()
            Resultado = Nothing
            salida = ex.Message
        End Try

        Return Serialize(salida, True)

    End Function

    Public Function CONSULTA_GENERICA_SP4_SERIALIZADA(ByVal NOMBRESP As Object, _
                                                      ByVal XMLParams As Object, _
                                                      ByVal ObjetoBinario As Object, _
                                                      ByVal vConnectionString As String) As Byte()
        Dim Cmd As New SqlCommand
        Dim Resultado As Data.DataSet = New Data.DataSet()
        Dim salida As String = ""
        Dim Serv As New Class_Servidor
        Dim XMLParams_str As String = ""


        'JMCA(20120809)
        'OBJETIVO:  CONSULTA PERMITE EJECUTAR EN FORMA GENÉRICA CUALQUIER SP.
        '           PASANDO SU NOMBRE, PARÁMETROS Y OBJETO SERIALIZADO.
        '           ESTÁ DISEÑADO ESPECIALMENTE
        '           PARA EJECUTAR PROCEDIMIENTOS QUE REALIZAN GRABAN IMAGENES
        Try
            Serv.Conectar(vConnectionString)
            With Cmd
                .CommandText = Deserialize(NOMBRESP)
                .CommandType = Data.CommandType.StoredProcedure
                .Connection = Serv.Conec
            End With

            XMLParams_str = Deserialize(XMLParams)
            If XMLParams_str.Trim <> "" Then
                CargaParametrosSP(Cmd, XMLParams_str)
            End If

            Cmd.Parameters.AddWithValue("@OBJETOBINARIO", ObjetoBinario)

            Dim SqlAdapter As New SqlDataAdapter
            SqlAdapter.SelectCommand = Cmd

            SqlAdapter.Fill(Resultado, "CONSULTA_GENERICA_SP4_SERIALIZADA")

            Serv.Desconectar()

            If Resultado.Tables(0).Rows.Count > 0 Then
                If Resultado.Tables(0).Rows(0).Item(0).ToString <> "-1" Then
                    salida = "OK" & " - " & Resultado.Tables(0).Rows(0).Item(0).ToString
                Else
                    For Each FILA As DataRow In Resultado.Tables(0).Rows
                        salida = "Procedimiento: " & FILA("ERROR_PROCEDURE_") & Chr(13) & _
                                 "Código error: " & FILA("ERROR_NUMBER_") & Chr(13) & _
                                 "Descripción: " & FILA("ERROR_MESSAGE_") & Chr(13) & _
                                 "Número línea: " & FILA("ERROR_LINE_")
                    Next
                End If
            End If

        Catch ex As Exception
            Serv.Desconectar()
            Resultado = Nothing
            salida = ex.Message
        End Try

        Return Serialize(salida, True)

    End Function

    Public Sub CargaParametrosSP(ByRef cmd As SqlCommand, ByVal XMLParams As String)
        '----------------------------------------------------------------
        'LOS PARÁMETROS DEL PROCEDIMIENTO ALMACENADO VIENEN EN EL FORMATO
        'DEL SIGUIENTE XML
        '<PARAMS><PARAM nombre="@CODIGO" valor="100-001"/></PARAMS>

        Dim oDOM As New System.Xml.XmlDocument
        Dim listaParametros As System.Xml.XmlNodeList
        Dim itemn As System.Xml.XmlNode
        Dim VALOR As String
        Dim NOMBRE As String

        oDOM.LoadXml(XMLParams)
        listaParametros = oDOM.SelectNodes(".//PARAM")

        For Each itemn In listaParametros
            VALOR = itemn.Attributes.GetNamedItem("valor").Value
            NOMBRE = itemn.Attributes.GetNamedItem("nombre").Value

            cmd.Parameters.AddWithValue(NOMBRE, VALOR)
        Next
        '----------------------------------------------------------------


    End Sub
    Private Function Serialize(ByVal Obj As Object, ByVal AsByte As Boolean) As Byte()
        Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim ms As New IO.MemoryStream
        If Obj IsNot Nothing Then
            bf.Serialize(ms, Obj)
            Return ms.ToArray
        Else
            bf.Serialize(ms, "")
            Return ms.ToArray
        End If
    End Function
    Private Function Deserialize(ByVal Obj As Byte()) As Object

        If Obj IsNot Nothing Then
            Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
            Dim ms As New IO.MemoryStream(Obj)
            Return bf.Deserialize(ms)
        Else
            Return Nothing
        End If
    End Function

#End Region
End Class
