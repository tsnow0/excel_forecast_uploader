Attribute VB_Name = "DataUpload"
Sub UploadData()
Application.ScreenUpdating = False

    'Declare variables
    Dim frow As Integer, lrow As Integer, fcol As Integer, lcol As Integer
    Dim ochannel_id As Integer
    
    'Define variables
    frow = 30
    If Cells(frow, 2).Offset(1, 0).Value = "" Then
        lrow = frow
    Else
        lrow = Cells(frow, 2).End(xlDown).Row
    End If
    fcol = 25
    lcol = 36
    
    'Ensure that there are SKU's to be uploaded
    If Cells(frow, 2) = "" Then
        MsgBox ("No SKU's to be uploaded")
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        End
    Else
        Call BuildUpload(frow, lrow, fcol, lcol, ochannel_id)
    End If

Application.ScreenUpdating = True
End Sub
Sub BuildUpload(ByRef frow As Integer, ByRef lrow As Integer, ByRef fcol As Integer, ByRef lcol As Integer, ByRef ochannel_id As Integer)

    Dim oproduct_id As Long, month_ As Integer, year_ As Integer, acf_qty As Integer
    Dim OCDT As String
    Dim strSQL As String, strSQL2 As String
    
    
    OCDT = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    strSQL = ""
    
    'Build a SKU list uploader for the forecasts
    oRow = 30
    For rCounter = frow To lrow
        oCol = 60
        For cCounter = fcol To lcol
            new_forecast = Cells(rCounter, cCounter).Value
            original_forecast = Cells(oRow, oCol).Value
            
            If new_forecast <> original_forecast Then
                
                sku = Cells(oRow, 2).Value
                oproduct_id = Cells(oRow, 37).Value
                month_ = Cells(6, cCounter).Value
                year_ = Cells(5, cCounter).Value
                forecast_qty = Cells(rCounter, cCounter).Value
                channel = Cells(1, 4).Value
                entity = Cells(3, 2).Value
                
                'Build String
                strSQL = strSQL & "('" & oproduct_id & "'," & channel & "," & entity & "," & month_ & "," & year_ & "," & forecast_qty & "),"
                NewRows = NewRows + 1
            End If
            oCol = oCol + 1
        Next cCounter
        oRow = oRow + 1
    Next rCounter
    
    
    strSQL2 = "INSERT INTO analytics.achannel_forecasts (oproduct_id, ochannel_id, aentity_id, acf_month, acf_year, acf_qty) " _
            & "VALUES " & Left(strSQL, Len(strSQL) - 1) & " ON DUPLICATE KEY UPDATE acf_qty = VALUES(acf_qty)"
    
    Call UploadToHQ(strSQL2)

MsgBox "Upload is complete", vbInformation
    
    
End Sub

Sub UploadToHQ(ByRef strSQL As String)

'DB info
    Set rs = CreateObject("ADODB.Recordset")
    database_name = "test" ' Name of database
    user_id = "test_uploader" 'id user or username
    Password = "test123" 'Password
    server_name = "testserver.com"
    
'ADODB connection
    Set cn = CreateObject("ADODB.Connection") 'NEW STATEMENT
    
    Range("BT1") = strSQL
    
    cn.Open "Driver={MySQL ODBC 9.0 ANSI Driver};Server=" & server_name & ";Database=" & database_name & _
    ";Uid=" & user_id & ";Pwd=" & Password & ";"
    rs.Open strSQL, cn
    cn.Close

End Sub

