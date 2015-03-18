Attribute VB_Name = "modDHTML"
'El siguiente c�digo le permitir� utilizar el conjunto de propiedades
'de su navegador Web para mantener informaci�n a trav�s de distintas
'p�ginas DHTML.
Public objWebBrowser As WebBrowser

'PutProperty: Almacena informaci�n en Property llamando a esta funci�n.
'             La informaci�n necesaria abarca el nombre Property y el
'             valor de la propiedad que desee almacenar.
'
Public Sub PutProperty(strNombre As String, vntValor As Variant)
    
    'Verificar si se ha ejecutado una instancia del navegador.
    If objWebBrowser Is Nothing Then Set objWebBrowser = New WebBrowser
    
    'Llamar al m�todo PutProperty del navegador para almacenar el valor.
    objWebBrowser.PutProperty strNombre, vntValor

End Sub

'GetProperty: Extrae informaci�n de Property llamando a esta funci�n.
'             La informaci�n necesaria abarca el nombre Property y el
'             valor devuelto por la funci�n es el valor actual de la
'             propiedad.
'
Public Function GetProperty(strNombre As String) As Variant
    
    'Verificar si se ha ejecutado una instancia del navegador.
    If objWebBrowser Is Nothing Then Set objWebBrowser = New WebBrowser
    
    'Llamar al m�todo GetProperty para obtener el valor.
    GetProperty = objWebBrowser.GetProperty(strNombre)

End Function
