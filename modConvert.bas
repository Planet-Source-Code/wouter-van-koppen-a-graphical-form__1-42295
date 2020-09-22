Attribute VB_Name = "modConvert"
Option Explicit

Public Function ToUpper(Text As String) As String
    Dim Length As Integer
    Dim counter As Integer
    Dim CurrentChar As String
    Length = Len(Text)
    counter = 0
    Do While counter < Length
        counter = counter + 1
        ToUpper = ToUpper & Chr(Asc(Mid(Text, counter, 1)) - 32)
    Loop
End Function

