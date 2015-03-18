Attribute VB_Name = "modDHTML"
'El siguiente código le permitirá utilizar el conjunto de propiedades
'de su navegador Web para mantener información a través de distintas
'páginas DHTML.
Public objWebBrowser As WebBrowser

'PutProperty: Almacena información en Property llamando a esta función.
'             La información necesaria abarca el nombre Property y el
'             valor de la propiedad que desee almacenar.
'
Public Sub PutProperty(strNombre As String, vntValor As Variant)
    
    'Verificar si se ha ejecutado una instancia del navegador.
    If objWebBrowser Is Nothing Then Set objWebBrowser = New WebBrowser
    
    'Llamar al método PutProperty del navegador para almacenar el valor.
    objWebBrowser.PutProperty strNombre, vntValor

End Sub

'GetProperty: Extrae información de Property llamando a esta función.
'             La información necesaria abarca el nombre Property y el
'             valor devuelto por la función es el valor actual de la
'             propiedad.
'
Public Function GetProperty(strNombre As String) As Variant
    
    'Verificar si se ha ejecutado una instancia del navegador.
    If objWebBrowser Is Nothing Then Set objWebBrowser = New WebBrowser
    
    'Llamar al método GetProperty para obtener el valor.
    GetProperty = objWebBrowser.GetProperty(strNombre)

End Function
