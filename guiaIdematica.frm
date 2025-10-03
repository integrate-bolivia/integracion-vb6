VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Login"
      Height          =   855
      Left            =   1080
      TabIndex        =   3
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "IDEMATICA"
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SANDBOX"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "listado"
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SendGraphQLRequest
End Sub
Private Sub TestMSXML()
    On Error GoTo ErrorHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    MsgBox "MSXML2.XMLHTTP está instalado correctamente."
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub
Private Sub SendGraphQLRequest()
    Dim http As Object
    Dim url As String
    Dim query As String
    Dim requestBody As String
    Dim response As String

    ' URL del endpoint GraphQL
    url = "https://sandbox.isipass.net/api"
    
     Dim filePath As String
    Dim fileContent As String
    Dim fileNumber As Integer
    
    filePath = "C:\Users\vboxuser\Downloads\test1.txt"
    fileNumber = FreeFile
    
    ' Abrir el archivo en modo lectura
    Open filePath For Input As #fileNumber
    
    ' Leer todo el contenido del archivo
    fileContent = Input$(LOF(fileNumber), fileNumber)
    
    ' Cerrar el archivo
    Close #fileNumber
    
    ' Mostrar el contenido
    MsgBox fileContent

    ' Consulta GraphQL
     query = fileContent

    ' Crear el objeto XMLHTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Abrir la conexión
    http.Open "POST", url, False

    ' Establecer los encabezados de la solicitud
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Authorization", "Bearer [TOKEN_CLIENTE]"
    

    ' Enviar la solicitud
    http.send query

    ' Verificar el estado de la respuesta
    If http.Status = 200 Then
        ' Leer la respuesta
        response = http.responseText
        Set jsonObject = JsonConverter.ParseJson(response)
        MsgBox "Respuesta del servidor: " & jsonObject("data")("licencia")("tienda")
    Else
        MsgBox "Error en la solicitud. Código de estado: " & http.Status
    End If

    ' Liberar el objeto
    Set http = Nothing
End Sub

Private Sub Command2_Click()
     Call SendGraphQLRequest2
End Sub

Private Sub SendGraphQLRequest2()
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim mutation As String

    ' URL del endpoint GraphQL
    url = "https://sandbox.isipass.net/api"
    Dim filePath As String
    Dim fileContent As String
    Dim fileNumber As Integer
    
    filePath = "C:\Users\vboxuser\Downloads\test2.txt"
    fileNumber = FreeFile
    
    ' Abrir el archivo en modo lectura
    Open filePath For Input As #fileNumber
    
    ' Leer todo el contenido del archivo
    fileContent = Input$(LOF(fileNumber), fileNumber)
    
    ' Cerrar el archivo
    Close #fileNumber
    
    ' Mostrar el contenido
    MsgBox fileContent

    ' Cuerpo de la mutación GraphQL en formato JSON
    mutation = fileContent
    requestBody = mutation
    MsgBox requestBody
    
    ' Crear el objeto XMLHTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Abrir la conexión
    http.Open "POST", url, False

    ' Establecer los encabezados de la solicitud
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Authorization", "Bearer [TOKEN_CLIENTE]"

    ' Enviar la solicitud
    http.send requestBody


    ' Verificar el estado de la respuesta
    If http.Status = 200 Then
        ' Leer la respuesta
        response = http.responseText
        ' Parsear la respuesta JSON
        Dim jsonObject As Object
        Set jsonObject = JsonConverter.ParseJson(response)

        ' Extraer los valores de la respuesta
        Dim numeroFactura As String
        Dim state As String

        numeroFactura = jsonObject("data")("facturaCompraVentaCreate")("representacionGrafica")("pdf")
       

        ' Mostrar los valores en un mensaje
        MsgBox "Número de Factura: " & numeroFactura

    Else
        MsgBox "Error en la solicitud. Código de estado: " & http.Status & http.responseText
    End If

    ' Liberar el objeto
    Set http = Nothing
End Sub

Private Sub Command3_Click()
    SendGraphQLRequest3
End Sub
Private Sub SendGraphQLRequest3()
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim mutation As String

    ' URL del endpoint GraphQL
    url = "https://api.idematica.net/api"
    Dim filePath As String
    Dim fileContent As String
    Dim fileNumber As Integer
    
    filePath = "C:\Users\vboxuser\Downloads\test3.txt"
    
    fileNumber = FreeFile
    
    ' Abrir el archivo en modo lectura
    Open filePath For Input As #fileNumber
    
    ' Leer todo el contenido del archivo
    fileContent = Input$(LOF(fileNumber), fileNumber)
    
    ' Cerrar el archivo
    Close #fileNumber
    
    ' Mostrar el contenido
    MsgBox fileContent

    ' Cuerpo de la mutación GraphQL en formato JSON
    mutation = fileContent
    requestBody = mutation
    MsgBox requestBody
    
    ' Crear el objeto XMLHTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Abrir la conexión
    http.Open "POST", url, False

    ' Establecer los encabezados de la solicitud
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Authorization", "Bearer [TOKEN_CLIENTE]"

    ' Enviar la solicitud
    http.send requestBody


    ' Verificar el estado de la respuesta
    If http.Status = 200 Then
        ' Leer la respuesta
        response = http.responseText
        ' Parsear la respuesta JSON
        Dim jsonObject As Object
        Set jsonObject = JsonConverter.ParseJson(response)

        ' Extraer los valores de la respuesta
        Dim numeroFactura As String
        Dim state As String

        numeroFactura = jsonObject("data")("facturaCompraVentaCreate")("representacionGrafica")("pdf")
       

        ' Mostrar los valores en un mensaje
        MsgBox "Número de Factura: " & numeroFactura

    Else
        MsgBox "Error en la solicitud. Código de estado: " & http.Status & http.responseText
    End If

    ' Liberar el objeto
    Set http = Nothing
End Sub

Private Sub Command4_Click()
Dim shop As String
    Dim email As String
    Dim password As String
    
    ' Valores dinámicos
    shop = "URL Comercio "
    email = "usuario@gmail.com"
    password = "password"
    
    SendGraphQLRequestFromFile "C:\Users\vboxuser\Downloads\test5.txt", shop, email, password
End Sub
Private Sub SendGraphQLRequestFromFile(filePath As String, shop As String, email As String, password As String)
    Dim http As Object
    Dim fileContent As String
    Dim fileNumber As Integer
    Dim requestBody As String
    Dim response As String
    
    ' Leer archivo
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileContent = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber
    
    ' Reemplazar variables
    requestBody = Replace(fileContent, "{{shop}}", shop)
    requestBody = Replace(requestBody, "{{email}}", email)
    requestBody = Replace(requestBody, "{{password}}", password)
    
    ' Crear objeto HTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", "https://api.idematica.net/compra-venta", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    
    ' Enviar
    http.send requestBody
    
    ' Manejar respuesta
    If http.Status = 200 Then
        response = http.responseText
        MsgBox "Respuesta: " & vbCrLf & response
    Else
        MsgBox "Error en la solicitud. Código: " & http.Status & vbCrLf & http.responseText
    End If
    
    Set http = Nothing
End Sub
