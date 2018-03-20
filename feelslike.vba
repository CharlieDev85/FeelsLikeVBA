'The feelslike temperature is calculated as follows
'This code assumes units of Fahrenheit, MPH, and Relative Humidity by percentage.

Sub feelslike()
    Dim vTemperature As Single
    Dim vWindSpeed As Single
    Dim vRelativeHumidity As Single
    Dim vFeelsLike As Single
    Dim lastRow As Integer
    Dim i As Integer
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    On Error GoTo Oops
    
    For i = 2 To lastRow
        vTemperature = Cells(i, 5)
        vWindSpeed = Cells(i, 7)
        vRelativeHumidity = Cells(i, 6)
        
'Try Wind Chill first
        If vTemperature <= 50 And vWindSpeed >= 3 Then
            vFeelsLike = 35.74 + (0.6215 * vTemperature) - 35.75 * (vWindSpeed ^ 0.16) + ((0.4275 * vTemperature) * (vWindSpeed ^ 0.16))
        Else
            vFeelsLike = vTemperature
        End If
        
'Replace it with the Heat Index, if necessary
        If vFeelsLike = vTemperature And vTemperature >= 80 Then
            vFeelsLike = 0.5 * (vTemperature + 61 + ((vTemperature - 68) * 1.2) + (vRelativeHumidity * 0.094))
            
            If vFeelsLike >= 80 Then
                vFeelsLike = -42.379 + 2.04901523 * vTemperature + 10.14333127 * vRelativeHumidity - 0.22475541 * vTemperature * vRelativeHumidity - 0.00683783 * vTemperature * vTemperature - 0.05481717 * vRelativeHumidity * vRelativeHumidity + 0.00122874 * vTemperature * vTemperature * vRelativeHumidity + 0.00085282 * vTemperature * vRelativeHumidity * vRelativeHumidity - 0.00000199 * vTemperature * vTemperature * vRelativeHumidity * vRelativeHumidity
                
                If vRelativeHumidity < 13 And vTemperature >= 80 And vTemperature <= 112 Then
                    vFeelsLike = vFeelsLike - ((13 - vRelativeHumidity) / 4) * VBA.Sqr((17 - Abs(vTemperature - 95)) / 17)
                End If
                
                If vRelativeHumidity > 85 And vTemperature >= 80 And vTemperature <= 87 Then
                    vFeelsLike = vFeelsLike + ((vRelativeHumidity - 85) / 10) * ((87 - vTemperature) / 5)
                End If
                
            End If
            
        End If
        
'MsgBox ("Feels like: " & Round(vFeelsLike, 1) & "°F")
        Cells(i, 8).Value = vFeelsLike
        Next i
        MsgBox ("Done!  :)")
        Exit Sub
        
Oops:
        Dim text As String
        text = "Check row number: " & i
        MsgBox (text)
        
    End Sub
    
    
    

