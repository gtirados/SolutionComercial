Attribute VB_Name = "modMain"
Public Function FormatoFecha_yyyyMMdd(xFecha As String) As String

 Dim dayPart As Integer
    Dim monthPart As Integer
    Dim yearPart As Integer
    Dim validDate As Date
    
    dayPart = CInt(Mid(xFecha, 1, 2))
    monthPart = CInt(Mid(xFecha, 4, 2))
    yearPart = CInt(Mid(xFecha, 7, 4))
    
    validDate = DateSerial(yearPart, monthPart, dayPart)
    
    FormatoFecha_yyyyMMdd = Format(validDate, "yyyymmdd")

End Function
