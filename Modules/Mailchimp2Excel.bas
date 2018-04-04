Attribute VB_Name = "Mailchimp2Excel"
''
' Mailchimp2excel v1
' (c) Joseph Banegas - https://github.com/bonez0607/mailchimp-to-excel
'
' Mailchimp2Excel
'
' @author Joseph Banegas
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Sub get_list(apiKey As String, listId As String, displayCount As Integer, sheetName As String)
    
    Dim objHTTP As Object
    Dim Parsed As Object
        
    Dim strUrl As String
    
    strUrl = "https://us9.api.mailchimp.com/3.0/lists/" & listId & "/members?count=" & displayCount

    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Open "GET", strUrl, False
    objHTTP.SetRequestHeader "Content-Type", "application/json"
    objHTTP.SetRequestHeader "Authorization", "Basic " & Base64Encode("user:" & apiKey)
    objHTTP.Send
    
    Set Parsed = JsonConverter.ParseJson(objHTTP.ResponseText)

    Call get_list_info(Parsed, sheetName)
End Sub
'Returns info for all subscribers and merge tags
Private Sub get_list_info(oJSON As Object, sheetName As String)
    
    Dim i As Integer
    Dim totalSubscribers As Long
    Dim fieldKeys As Variant

    totalSubscribers = oJSON("members").Count
    fieldKeys = oJSON("members")(1)("merge_fields").Keys
   
    Call import_colHeaders(fieldKeys, sheetName)
    Call import_subscriberData(oJSON, fieldKeys, totalSubscribers, sheetName)
    
End Sub

'Writes header row to first column from merge_field names
Private Sub import_colHeaders(headers As Variant, sheetName As String)
    Dim col As Integer
    Dim headerRow As Integer
    
    col = 1
    headerRow = 1
    For Each header In headers
        Sheets(sheetName).Cells(1, col) = header
        col = col + 1
    Next header
End Sub
'Writes subscriber information to specified sheet leaving space for header column
Private Sub import_subscriberData(oJSON As Object, headers As Variant, totalSubscribers As Long, sheetName As String)
    Dim col As Integer
    Dim row As Integer
    Dim i As Long
    
    row = 2
    col = 1
    
    For i = 1 To totalSubscribers
           For Each header In headers
                 Sheets(sheetName).Cells(row, col) = oJSON("members")(i)("merge_fields")(header)
               col = col + 1
           
           Next header
           col = 1
           row = row + 1
    Next i

End Sub

