'///////////////////////////////////////////////////////////////////////
'//                                                                   //
'// Data Serialized Transmission V.2.0 (DTSTX)                        //
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
Public Class DataSerializedTX
    Public Function Deserialize(ByVal Obj As Byte(), Optional ByVal TIPO As String = "") As Object
        Dim objDeserialize As Object

        If Obj IsNot Nothing Then
            Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
            Dim ms As New IO.MemoryStream(Obj)
            objDeserialize = bf.Deserialize(ms)

            If objDeserialize.GetType.Name.Trim.ToUpper = TIPO Or TIPO = "" Then
                Return objDeserialize
            Else
                If objDeserialize.GetType.Name.Trim.ToUpper = "STRING" Then
                    MsgBox("Ha sucedido una excepción : " & Chr(13) & Chr(13) & objDeserialize.ToString, MsgBoxStyle.Exclamation, "")
                End If

                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Function Serialize(ByVal Obj As Object, ByVal AsByte As Boolean) As Byte()
        Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim ms As New IO.MemoryStream
        bf.Serialize(ms, Obj)
        Return ms.ToArray
    End Function
End Class
