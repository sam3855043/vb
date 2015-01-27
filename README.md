# vb
'
''''''''''''''''''''''''''''''SQL
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
''''''''''''''''''''''''''''''SQL
Dim i As Integer
Dim j As Integer
Dim adressgit As Integer
Dim rs2 As ADODB.Recordset






Private Sub Close_Click()
    rs.Close
End Sub









Private Sub ComPort_button_Click()
If (ComPort_button.Caption = "Connect") Then
    On Error GoTo Err:
    MSComm1.CommPort = Val(Comm_port.Text)
    MSComm1.Settings = "9600,n,8,1"
    'MSComm1.RThreshold = 1'設定當接收到1個byte就觸發comEvReceive事件
    MSComm1.PortOpen = True
   
    ''TimerReadTemp.Enabled = True
    ''TimerReadHum.Enabled = True
    TimerReadHumTemp.Enabled = True
    OpenFlag = ture
    ComPort_button.Caption = "Disconnect"
ElseIf (ComPort_button.Caption = "Disconnect") Then
    ComPort_button.Caption = "Connect"
    'TimerReadTemp.Enabled = False
    MSComm1.PortOpen = False
    'TimerReadHum.Enabled = False
    TimerReadHumTemp.Enabled = False



End If
Exit Sub
Err:


End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Err
conn.Close
Err:
End Sub

Private Sub MoveNext_Click()
    rs.Movenext              '將欄位移到下一個
End Sub

Private Sub Form_Load()
    On Error GoTo Err
    Me.Top = (Screen.Height - Me.Height) / 2    '窗口居中
    Me.Left = (Screen.Width - Me.Width) / 2
    i = 0
    adressgit = 1
    'label set
    'Label1 Title cption font size
    
    ' Button enable
    'TimerReadTemp.Enabled = False
    'TimerReadHum.Enabled = False
    TimerReadHumTemp.Enabled = False

    ' variable Flag

    
    'mysql
    'Set conn = New ADODB.Connection
    'Set rs = New ADODB.Recordset

    'conn.CursorLocation = adUseClient
    '連線字串                                      'MySQL server IP            '資料庫       '帳號    '密碼
    'conn.ConnectionString = "DRIVER=MySQL ODBC 5.1 DRIVER;SERVER=192.168.0.50;DATABASE=test;UID=root;password=root"
    
   ' conn.Open
    
    '資料表test
    'rs.Open "test1", conn, 2, 3
    
    
    'Set DataGrid1.DataSource = rs
    
    
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset

    conn.CursorLocation = adUseClient
    '連線字串                                      'MySQL server IP            '資料庫       '帳號    '密碼
    conn.ConnectionString = "DRIVER=MySQL ODBC 5.1 DRIVER;SERVER=localhost;DATABASE=test;UID=root;password=root"
    
    conn.Open
    
    '資料表test
    rs.Open "SwitchData", conn, 2, 3
    
    Set DataGrid1.DataSource = rs
    

Err:
End Sub



Private Sub Timer1_Timer()
If (Text4 = "T") Then
    Text1(0) = (Val(GETT(adressgit)) - 4000) / 100
ElseIf (Text4 = "H") Then
    Text1(1) = (Val(GETH(adressgit))) / 100
ElseIf (Text4 = "HT") Then
  
 
  
  
  If (Check1.Value) Then
    Call GETTH(adressgit)
    adressgit = adressgit + 1
    If (adressgit > 9) Then
     adressgit = 0
    End If
    Text9.Text = adressgit
  Else
    Call GETTH(Val(Text9.Text))
  End If
 

  
'
'     Set conn = New ADODB.Connection
'     Set rs = New ADODB.Recordset
'     Set rs2 = New ADODB.Recordset 'SQL
'     conn.CursorLocation = adUseClient
'    '連線字串                                      'MySQL server IP            '資料庫       '帳號    '密碼
'     conn.ConnectionString = "DRIVER=MySQL ODBC 5.1 DRIVER;SERVER=127.0.0.1;DATABASE=test;UID=root;password=s186354"
'     conn.Open
'     rs.Open "TEMPHUM", conn, 2, 3


  If (rs.Fields("Humidity") = Val(Text1(3).Text)) Then
  Else
    rs.Fields("Humidity") = Val(Text1(3).Text)
    rs.Update
  End If
  If (rs.Fields("Tempture") = Val(Text1(2).Text)) Then
  Else
    rs.Fields("Tempture") = Text1(2).Text
    rs.Update
  End If

  If (rs.Fields("Date") = Now()) Then
  Else
    rs.Fields("Date") = Now()
    rs.Update
  End If
     
    ' rs2.Open SQL, conn, 2, 3
    ' Set DataGrkid1.DataSource = rs
 

End If

End Sub

Private Sub TimerReadHum_Click()

HumTemp_Flag = 1
ReadStatus = "H"
Text4 = ReadStatus
If TimerReadHum.Caption = "讀取轉轍器濕度" Then
   TimerReadHum.Caption = "停止"
   
   
    Timer1.Enabled = True
    ' Button enable
    ComPort_button.Enabled = False
    'TimerReadTemp.Enabled = False
    TimerReadHumTemp.Enabled = False

Else
    'TimerReadHum.Caption = "讀取轉轍器濕度"
   
    
    ' Button enable
    ComPort_button.Enabled = True
    '''TimerReadTemp.Enabled = True
    TimerReadHumTemp.Enabled = True

    Timer1.Enabled = False
End If
End Sub

Private Sub TimerReadHumTemp_Click()
HumTemp_Flag = 1
ReadStatus = "HT"
Text4 = ReadStatus
If TimerReadHumTemp.Caption = "讀取轉轍器溫濕度" Then
   TimerReadHumTemp.Caption = "停止"
   
  '
    Timer1.Enabled = True
   
    ' Button enable
     ComPort_button.Enabled = False 'connect button disable
    'TimerReadTemp.Enabled = False
    'TimerReadHum.Enabled = False

Else
    TimerReadHumTemp.Caption = "讀取轉轍器溫濕度"
   
    
    ' Button enable
    ComPort_button.Enabled = True 'connect button disable
    '''TimerReadTemp.Enabled = True
    'TimerReadHum.Enabled = True

    Timer1.Enabled = False
   
End If

End Sub



Private Sub TimerReadTemp_Click()


HumTemp_Flag = 1
ReadStatus = "T"
Text4 = ReadStatus
If TimerReadTemp.Caption = "讀取轉轍器溫度" Then
   TimerReadTemp.Caption = "停止"
   Timer1.Enabled = True
   
  
    ' Button enable
    
     ComPort_button.Enabled = False 'connect button disable
    'TimerReadHum.Enabled = False
    TimerReadHumTemp.Enabled = False

Else
    TimerReadTemp.Caption = "讀取轉轍器溫度"
   
    Timer1.Enabled = False
    ' Button enable
    
    ComPort_button.Enabled = True 'connect button disable
    'TimerReadHum.Enabled = True
    TimerReadHumTemp.Enabled = True


End If

End Sub



Sub GETTH(ByVal TT As Byte)
    Dim strCommand As String
    Dim strTemp As String
    Dim OverFlag As Boolean
    Dim strReturn As String
    Dim TimeOver As Single
    Dim OverCount As Integer
    Dim intType As Integer
    Dim ext2 As String
    Dim ext3 As String
    Dim arrayx(8) As Byte
    Dim lonCRC As Long
    Dim intCnt As Integer
    Dim intBit As Integer
    Dim intLeng As Integer
    Dim intTemp As Integer
    Dim bytTemp As Byte
    Dim bytRes() As Byte
    Dim getdata(20) As Byte
    Dim CRCCK(2) As Byte
    
    arrayx(0) = Hex(TT)
    arrayx(1) = "&H04"
    arrayx(2) = "&H00"
    arrayx(3) = "&H00"
    arrayx(4) = "&H00"
    arrayx(5) = "&H02"
     lonCRC = &HFFFF&
    For intCnt = 0 To 5
        lonCRC = lonCRC Xor arrayx(intCnt)
        For intBit = 0 To 7
            intTemp = lonCRC Mod 2
            lonCRC = lonCRC \ 2
            If intTemp = 1 Then
                lonCRC = lonCRC Xor &HA001&
            End If
        Next intBit
    Next intCnt

    arrayx(6) = lonCRC Mod 256
    arrayx(7) = lonCRC \ 256
    
    Debug.Print arrayx
    MSComm1.Output = arrayx

    TimeOver = Timer()
    OverFlag = False
     
    bytRes = MSComm1.Input
     
     ext3 = ""
    For intCnt = 0 To UBound(bytRes)
        If intCnt <> 0 Then
            ext3 = ext3 + " "
        End If
        getdata(intCnt) = Val(bytRes(intCnt))
        If bytRes(intCnt) < 16 Then
            ext3 = ext3 & "0" & Hex(bytRes(intCnt))
        Else
            ext3 = ext3 & Hex(bytRes(intCnt))
        End If
    Next intCnt
Text5 = ext3
    lonCRC = &HFFFF&
    For intCnt = 0 To 6
        lonCRC = lonCRC Xor getdata(intCnt)
        For intBit = 0 To 7
            intTemp = lonCRC Mod 2
            lonCRC = lonCRC \ 2
            If intTemp = 1 Then
                lonCRC = lonCRC Xor &HA001&
            End If
        Next intBit
    Next intCnt
 
    CRCCK(0) = lonCRC Mod 256
    CRCCK(1) = lonCRC \ 256
    
    Text6 = Hex(CRCCK(0))
    Text7 = Hex(CRCCK(1))
    
    
    If (getdata(7) = CRCCK(0)) Then
     If (getdata(8) = CRCCK(1)) Then
      Text1(2) = (getdata(3) * 256 + getdata(4) - 4000) / 100
      Text1(3) = (getdata(5) * 256 + getdata(6)) / 100
     End If
    End If
    

End Sub

Function GETT(ByVal TT As Byte) As String
    Dim strCommand As String
    Dim strTemp As String
    Dim OverFlag As Boolean
    Dim strReturn As String
    Dim TimeOver As Single
    Dim OverCount As Integer
    Dim intType As Integer
    Dim ext2 As String
    Dim ext3 As String
    Dim arrayx(8) As Byte
    Dim lonCRC As Long
    Dim intCnt As Integer
    Dim intBit As Integer
    Dim intLeng As Integer
    Dim intTemp As Integer
    Dim bytTemp As Byte
    Dim bytRes() As Byte
    Dim getdata(10) As Byte
    Dim CRCCK(2) As Byte
    
    arrayx(0) = Hex(TT)
    arrayx(1) = "&H04"
    arrayx(2) = "&H00"
    arrayx(3) = "&H00"
    arrayx(4) = "&H00"
    arrayx(5) = "&H01"
     lonCRC = &HFFFF&
    For intCnt = 0 To 5
        lonCRC = lonCRC Xor arrayx(intCnt)
        For intBit = 0 To 7
            intTemp = lonCRC Mod 2
            lonCRC = lonCRC \ 2
            If intTemp = 1 Then
                lonCRC = lonCRC Xor &HA001&
            End If
        Next intBit
    Next intCnt

    arrayx(6) = lonCRC Mod 256
    arrayx(7) = lonCRC \ 256
    
    Debug.Print arrayx
    MSComm1.Output = arrayx

    TimeOver = Timer()
    OverFlag = False
     
    bytRes = MSComm1.Input
     
     ext3 = ""
    For intCnt = 0 To UBound(bytRes)
        If intCnt <> 0 Then
            ext3 = ext3 + " "
        End If
        getdata(intCnt) = Val(bytRes(intCnt))
        If bytRes(intCnt) < 16 Then
            ext3 = ext3 & "0" & Hex(bytRes(intCnt))
        Else
            ext3 = ext3 & Hex(bytRes(intCnt))
        End If
    Next intCnt
Text5 = ext3
    lonCRC = &HFFFF&
    For intCnt = 0 To 4
        lonCRC = lonCRC Xor getdata(intCnt)
        For intBit = 0 To 7
            intTemp = lonCRC Mod 2
            lonCRC = lonCRC \ 2
            If intTemp = 1 Then
                lonCRC = lonCRC Xor &HA001&
            End If
        Next intBit
    Next intCnt
 
    CRCCK(0) = lonCRC Mod 256
    CRCCK(1) = lonCRC \ 256
    
    Text6 = Hex(CRCCK(0))
    Text7 = Hex(CRCCK(1))
    
    
    If (getdata(5) = CRCCK(0)) Then
     If (getdata(6) = CRCCK(1)) Then
      GETT = (getdata(3) * 256) + (getdata(4))
     End If
    End If
    

End Function

Function GETH(ByVal TT As Byte) As String

    Dim strCommand As String
    Dim strTemp As String
    Dim OverFlag As Boolean
    Dim strReturn As String
    Dim TimeOver As Single
    Dim OverCount As Integer
    Dim intType As Integer
    Dim ext2 As String
    Dim ext3 As String
    Dim arrayx(8) As Byte
    Dim lonCRC As Long
    Dim intCnt As Integer
    Dim intBit As Integer
    Dim intLeng As Integer
    Dim intTemp As Integer
    Dim bytTemp As Byte
    Dim bytRes() As Byte
    Dim getdata(10) As Byte
    Dim CRCCK(2) As Byte
    
    arrayx(0) = Hex(TT)
    arrayx(1) = "&H04"
    arrayx(2) = "&H00"
    arrayx(3) = "&H01"
    arrayx(4) = "&H00"
    arrayx(5) = "&H01"
     lonCRC = &HFFFF&
    For intCnt = 0 To 5
        lonCRC = lonCRC Xor arrayx(intCnt)
        For intBit = 0 To 7
            intTemp = lonCRC Mod 2
            lonCRC = lonCRC \ 2
            If intTemp = 1 Then
                lonCRC = lonCRC Xor &HA001&
            End If
        Next intBit
    Next intCnt

    arrayx(6) = lonCRC Mod 256
    arrayx(7) = lonCRC \ 256
    
    Debug.Print arrayx
    MSComm1.Output = arrayx

    TimeOver = Timer()
    OverFlag = False
     
    bytRes = MSComm1.Input
     
     ext3 = ""
    For intCnt = 0 To UBound(bytRes)
        If intCnt <> 0 Then
            ext3 = ext3 + " "
        End If
        getdata(intCnt) = Val(bytRes(intCnt))
        If bytRes(intCnt) < 16 Then
            ext3 = ext3 & "0" & Hex(bytRes(intCnt))
        Else
            ext3 = ext3 & Hex(bytRes(intCnt))
        End If
    Next intCnt
Text5 = ext3
    lonCRC = &HFFFF&
    For intCnt = 0 To 4
        lonCRC = lonCRC Xor getdata(intCnt)
        For intBit = 0 To 7
            intTemp = lonCRC Mod 2
            lonCRC = lonCRC \ 2
            If intTemp = 1 Then
                lonCRC = lonCRC Xor &HA001&
            End If
        Next intBit
    Next intCnt
 
    CRCCK(0) = lonCRC Mod 256
    CRCCK(1) = lonCRC \ 256
    
    Text6 = Hex(CRCCK(0))
    Text7 = Hex(CRCCK(1))
    
    If (getdata(5) = CRCCK(0)) Then
     If (getdata(6) = CRCCK(1)) Then
       GETH = (getdata(3) * 256) + (getdata(4))
     End If
    End If
End Function

