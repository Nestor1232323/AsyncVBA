# AsyncVBA

AsyncVBA - **Asynchronous** client for Visual Basic For Applications, preferably used alongside VBA-JSON.

## Functions:
* GET/POST/PUT/DELETE responds
* Can use with JSON
* UTF-8 build-in support
* Custom headers
* Excel doesn't freeze

## How to usage
Example of **code**:
``` vba
Option Explicit

Dim req As AsyncVBA 

' get
Sub TestAsyncGET()
    Set req = New AsyncVBA
    
    req.SendRequest "https://jsonplaceholder.typicode.com/todos/1", "GET"
    
    Debug.Print "GET request sent asynchronously..."
End Sub

' post
Sub TestAsyncPOST()
    Set req = New AsyncVBA
    
    Dim json As String
    json = "{""model"":""gpt-3.5-turbo"",""messages"":[{""role"":""user"",""content"":""Hello""}]}"
    
    Dim headers As Variant
    headers = Array( _
        "Authorization: Bearer PUT_YOUR_KEY_HERE", _
        "Content-Type: application/json" _
    )
    
    req.SendRequest "https://openrouter.ai/api/v1/chat/completions", "POST", headers, json
    
    Debug.Print "POST request sent asynchronously..."
End Sub

' put
Sub TestAsyncPUT()
    Set req = New AsyncVBA
    
    Dim data As String
    data = "{""name"":""VBA Test"",""status"":""active""}"
    
    Dim headers As Variant
    headers = Array("Content-Type: application/json")
    
    req.SendRequest "https://jsonplaceholder.typicode.com/posts/1", "PUT", headers, data
    
    Debug.Print "PUT request sent asynchronously..."
End Sub

```
