Attribute VB_Name = "RNDAA"
'------------------------------------------------------'
'|      Random User-Agent/Accept-Language Module      |'
'|                                                    |'
'| Reference UA() and AL() in the winsock header:     |'
'|                                                    |'
'| strSData = strSData & AL()                         |'
'| strSData = strSData & UA()                         |'
'| strSData = strSData & "Host: www.yahoo.com"        |'
'|                                                    |'
'| <Kurt.K> <adidos@adidos.cjb.net>                   |'
'|                                                    |'
'|      Random User-Agent/Accept-Language Module      |'
'------------------------------------------------------'

Global MSIEr(0 To 9) As String
Global WINr(0 To 6) As String
Global MOZr(0 To 8) As String
Global ALf(0 To 11) As String

Public Function UA()
UA = "User-Agent: Mozilla/" & MOZv & " (compatible; MSIE " & MSIEv & "; Windows " & WINv & ")"
End Function

Public Function AL()
AL = "Accept-Language: " & ALc
End Function


Public Function MSIEv()
MSIEr(0) = "5.0"
MSIEr(1) = "5.5"
MSIEr(2) = "5.0"
MSIEr(3) = "5.01"
MSIEr(4) = "4.0"
MSIEr(5) = "4.01"
MSIEr(6) = "5.5"
MSIEr(7) = "5.5"
MSIEr(8) = "5.0"
MSIEr(9) = "5.0"
Randomize
rannum1 = Int(Rnd * 10)
MSIEv = MSIEr(rannum1)
End Function

Public Function WINv()
WINr(0) = "98"
WINr(1) = "98"
WINr(2) = "95"
WINr(3) = "NT 5.0"
WINr(4) = "NT 4.0"
WINr(5) = "NT"
WINr(6) = "98"
Randomize
rannum2 = Int(Rnd * 7)
WINv = WINr(rannum2)
End Function

Public Function MOZv()
MOZr(0) = "4.0"
MOZr(1) = "4.0"
MOZr(2) = "4.0"
MOZr(3) = "4.0"
MOZr(4) = "4.0"
MOZr(5) = "4.0"
MOZr(6) = "3.01"
MOZr(7) = "3.0"
MOZr(8) = "2.0"
Randomize
rannum3 = Int(Rnd * 9)
MOZv = MOZr(rannum3)
End Function

Public Function ALc()
ALf(0) = "en"
ALf(1) = "en-us"
ALf(2) = "en-us"
ALf(3) = "en-us"
ALf(4) = "en-ca"
ALf(5) = "en-au"
ALf(6) = "de"
ALf(7) = "fr"
ALf(8) = "ru"
ALf(8) = "jp"
ALf(9) = "en-us"
ALf(10) = "en-us"
ALf(11) = "en-us"
Randomize
rannum4 = Int(Rnd * 12)
ALc = ALf(rannum4)
End Function


