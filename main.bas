

   Sub API_Req(JsonString As String)
   
    'Creating xml obj
    Dim xml_obj As MSXML2.XMLHTTP60
    Set xml_obj = New MSXML2.XMLHTTP60
    
    'Define URL Components
    base_url = "http://localhost:3333/"
    endpoint = "product-upload"
    
    'Combinding components
    apiurl = base_url + endpoint
    
    'Open a new request using the url
    xml_obj.Open bstrMethod:="POST", bstrURL:=base_url + endpoint
    
    'Setting headers
    xml_obj.setRequestHeader "Content-Type", "application/json"
    xml_obj.setRequestHeader "Accept", "application/json"
    
    'Send request with JSON body
    xml_obj.Send JsonString
    
    'Print the status
    Debug.Print "The request was " + xml_obj.statusText
    'Displays a succeful message in dialog box
    MsgBox "Your request was SUCCESSFUL!"
    
End Sub


Public Sub SOFinal()
    'Declare variables
    Dim outlook     As outlook.Application
    Dim item        As outlook.MailItem
    Dim html        As Object: Set html = CreateObject("htmlfile")
    Dim tables      As Object
    Dim table4       As Object
    Dim LineItemCollection As Collection
    Dim JSONCode As String
    
    'grabs the active window
    Set item = GetCurrentItem()
        ' Make sure it's a Mail Item
        If item.Class = olMail Then
            ' set the body of the email equal to the html from outlook email
            html.Body.innerHTML = item.HTMLBody
            ' Get all the table elements
            Set tables = html.getElementsByTagName("table")(5)
            Set table4 = html.getElementsByTagName("table")(4)
            'Gets an object of line items
            Set LineItemCollection = GetLineItems(tables)
            'adds order info object
                Set Dict = New Dictionary
                Dict.Add "endUser", table4.Rows(0).cells(1).innerText
                Dict.Add "customerPO", table4.Rows(1).cells(1).innerText
                Dict.Add "orderNum", table4.Rows(2).cells(1).innerText
                Dict.Add "lineItems", LineItemCollection
            'converts collection/dictionary object to JSON
            JSONCode = JsonConverter.ConvertToJson(Dict, Whitespace:=4)
            JSONCode = Replace(JSONCode, " \r\n", "")
            Debug.Print JSONCode
            API_Req JSONCode
        End If
   End Sub
Function GetCurrentItem() As Object
    Dim objApp As outlook.Application
           
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
       
    Set objApp = Nothing
End Function


