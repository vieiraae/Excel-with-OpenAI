Attribute VB_Name = "OpenAIModule"
''
' EXCEL with OPENAI
' https://github.com/vieiraae/Excel-with-OpenAI
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' @class OpenAIModule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
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
Option Explicit

Public Function GPT_GENERATE_INSIGHTS(TEXT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_GENERATE_INSIGHTS = "WIP"
End Function

Public Function GPT_ITEMIZE(TEXT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_ITEMIZE = "WIP"
End Function

Public Function GPT_ANSWER_QUESTION(QUESTION As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_ANSWER_QUESTION = GPT_COMPLETION("Provide an anwer for this question: " & QUESTION, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
End Function


Public Function GPT_GENERATE_NEW_ITEM(EXISTING_ITEMS As Range, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    Dim existingITems As String
    Dim row As Range
    existingITems = ""
    For Each row In EXISTING_ITEMS.Rows
        existingITems = existingITems & ", " & row.Value
    Next
    GPT_GENERATE_NEW_ITEM = GPT_COMPLETION("Generate a new item to complement the following ones: " & existingITems, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
End Function


Public Function GPT_EXTRACT_ENTITIES(TEXT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_EXTRACT_ENTITIES = "WIP"
End Function


Public Function GPT_EXTRACT_TABLE(TEXT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_EXTRACT_TABLE = GPT_COMPLETION("Summarize the following TEXT into a table in json format. ONLY JSON IS ALLOWED as an answer. TEXT: " & TEXT, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
End Function


Public Function GPT_CLASSIFY(CATEGORIES As Range, TEXT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    Dim categoriesText As String
    Dim row As Range
    categoriesText = ""
    For Each row In CATEGORIES.Rows
        categoriesText = categoriesText & ", " & row.Value
    Next
    GPT_CLASSIFY = GPT_COMPLETION("Classify the TEXT into 1 of the following categories: " & categoriesText & vbCrLf & "TEXT: " & TEXT, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
End Function


Public Function GPT_SENTIMENT_ANALYSIS(TEXT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_SENTIMENT_ANALYSIS = GPT_COMPLETION("Do a sentiment analysis over the following text and provide the result with ONE AND just ONE emoji in unicode. TEXT: " & TEXT, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
End Function

Public Function GPT_TRANSLATE_TO(LANGUAGE As String, TEXT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_TRANSLATE_TO = GPT_COMPLETION("Translate the following text to " & LANGUAGE & ": " & TEXT, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
End Function

Public Function GPT_SUMMARIZE(TEXT As String, Optional NUMBER_OF_WORDS As Integer = 0, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    If NUMBER_OF_WORDS > 0 Then
        GPT_SUMMARIZE = GPT_COMPLETION("Summarize the following text to " & NUMBER_OF_WORDS & " words: " & vbCrLf & TEXT, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
    Else
        GPT_SUMMARIZE = GPT_COMPLETION("Summarize the following text: " & vbCrLf & TEXT, SECONDS_TO_WAIT, MAX_TOKENS, TEMPERATURE, TOP_P)
    End If
End Function

Public Function GPT_GEN_PYTHON(PROMPT As String, Optional NUMBER_OF_WORDS As Integer = 0, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    GPT_GEN_PYTHON = ""
End Function

Public Function GPT_COMPLETION(PROMPT As String, Optional SECONDS_TO_WAIT As Integer = 0, Optional MAX_TOKENS As Integer = 100, Optional TEMPERATURE As Double = 1, Optional TOP_P As Double = 1) As String
    If SECONDS_TO_WAIT > 0 Then Wait SECONDS_TO_WAIT
    Dim FSO As New FileSystemObject
    Dim TS As TextStream
    Set TS = FSO.OpenTextFile(ActiveWorkbook.Path & "\" & "OpenAIModule.json", ForReading)
    Dim OpenAISettings
    Set OpenAISettings = JsonConverter.ParseJson(TS.ReadAll)
    TS.Close
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "POST", OpenAISettings("AZURE_OPENAI_ENDPOINT") & "openai/deployments/" & OpenAISettings("AZURE_OPENAI_DEPLOYMENT_MODEL") & "/completions?" & OpenAISettings("AZURE_OPENAI_API_VERSION"), False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.setRequestHeader "api-key", OpenAISettings("AZURE_OPENAI_KEY")
    Dim request As New Scripting.Dictionary
    request.Add "max_tokens", MAX_TOKENS
    If TEMPERATURE <> 1 Then request.Add "temperature", TEMPERATURE
    If TOP_P <> 1 Then request.Add "top_p", TOP_P
    request.Add "prompt", PROMPT
    Dim Payload
    Payload = JsonConverter.ConvertToJson(request, 2)
    Debug.Print Payload
    xhr.Send Payload
    Debug.Print xhr.responseText
    Dim response
    Set response = JsonConverter.ParseJson(xhr.responseText)
    If response.Exists("choices") Then
        GPT_COMPLETION = response("choices")(1)("text")
    Else
        GPT_COMPLETION = "ERROR: " & response("error")("message")
    End If
End Function


Sub Wait(secondsToWait As Integer)
    Dim endTime As Double
    endTime = Timer + secondsToWait
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

