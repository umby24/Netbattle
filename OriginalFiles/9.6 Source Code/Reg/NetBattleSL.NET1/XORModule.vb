Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module XORMod
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Sub CopyMemory Lib "KERNEL32"  Alias "RtlMoveMemory"(ByRef Dest As Any, ByRef Source As Any, ByVal Length As Integer)
	Private Const MaxLen As Short = 252
	
	Public Function FormatPacket(ByVal Data As String, ByVal Encrypt As Boolean) As String
		Dim X As Short
		If Encrypt Then Data = XOREncrypt(Data)
		Data = Left(Data, 256)
		Data = Chr(Len(Data) - 1) & Data
		FormatPacket = Data
	End Function
	Public Function GetPacket(ByRef Socket As AxMSWinsockLib.AxWinsock, ByVal BytesTotal As Integer, ByRef PacketBuffer() As String) As Boolean
		Dim Temp As String
		Dim Packet() As String
		Dim X As Integer
		Dim Y As Integer
		Dim R As Integer
		On Error GoTo Failed
		ReDim Packet(0)
		Y = 0
		Socket.GetData(Temp,  , BytesTotal)
		Temp = Socket.Tag & Temp
		Socket.Tag = ""
		''>>> Call WriteDebugLog("Packet Rcv: " & Str2Hex(Temp))
		R = Len(Temp)
		Do Until R = 0
			R = R - 1
			X = Asc(Left(Temp, 1)) + 1
			If R < X Then
				'>>> Call WriteDebugLog("Missing Data, Saving")
				Socket.Tag = Temp
				Exit Do
			End If
			Temp = Right(Temp, R)
			R = R - X
			Y = Y + 1
			ReDim Preserve Packet(Y)
			Packet(Y) = Left(Temp, X)
			Temp = Right(Temp, R)
		Loop 
		PacketBuffer = VB6.CopyArray(Packet)
		GetPacket = True
		Exit Function
Failed: 
		'>>> Call WriteDebugLog("Error " & Err.Number & " - " & Err.Description)
		GetPacket = False
	End Function
	Public Sub RawInterpret(ByRef Raw As String)
		Dim Temp As String
		Dim Packet() As String
		Dim X As Short
		Dim Y As Short
		Dim R As Short
		On Error GoTo Failed
		ReDim Packet(0)
		Y = -1
		Temp = Raw
		'>>> Call WriteDebugLog("Packet Rcv: " & Str2Hex(Temp))
		R = Len(Temp)
		While R <> 0
			R = R - 1
			X = Asc(Left(Temp, 1)) + 1
			Temp = Right(Temp, R)
			If R < X Then
				'>>> Call WriteDebugLog("Invalid Len Marker")
				GoTo Failed
			End If
			R = R - X
			Y = Y + 1
			ReDim Preserve Packet(Y)
			Packet(Y) = Left(Temp, X)
			Temp = Right(Temp, R)
		End While
		Exit Sub
Failed: 
		'>>> Call WriteDebugLog("Error " & Err.Number & " - " & Err.Description)
	End Sub
	Public Function XOREncrypt(ByVal Orig As String) As String
		Dim MainKey As Byte
		Dim iLen As Short
		Dim pLen As Short
		Dim iByte() As Byte
		Dim oByte() As Byte
		Dim tByte() As Byte
		Dim Checksum1 As Integer
		Dim Checksum2 As Integer
		Dim Checksum3 As Integer
		Dim Working As Integer
		Dim V As Short
		Dim W As Short
		Dim X As Short
		Dim Y As Short
		Dim Z As Short
		Dim Build As String
		On Error GoTo ETCTrap
		
		Orig = Left(Orig, MaxLen)
		iLen = Len(Orig)
		If iLen Mod 2 = 1 Then
			Orig = vbNullChar & Orig
			iLen = iLen + 1
		End If
		pLen = iLen + 4
		'UPGRADE_WARNING: Lower bound of array iByte was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iByte(pLen)
		'UPGRADE_WARNING: Lower bound of array oByte was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim oByte(pLen)
		'UPGRADE_WARNING: Lower bound of array tByte was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim tByte(pLen)
		'Alright.  The first thing we do is convert the string
		'to a byte array for speedy manipulation.  The first Chr
		'is a length marker, and the rest of the array is filled
		'with random characters.
		CopyMemory(iByte(1), Orig, Len(Orig))
		'    For X = 1 To iLen
		'        iByte(X) = Asc(Mid$(Orig, X, 1))
		'    Next X
		oByte = VB6.CopyArray(iByte)
		
		'This generates the main key.
		
		'--------------------
		MainKey = CByte(Int(Rnd() * 256))
		'--------------------
		
		'This is where the encryption takes place
		oByte(pLen) = MainKey Xor 12
		For X = 1 To iLen
			oByte(X) = oByte(X) Xor oByte(iLen - X + 1) Xor MainKey
		Next X
		
		'And this part generates a few checksums
		Checksum1 = 0
		Checksum2 = 0
		Checksum3 = 0
		
		'The first is simply an addition of the original characters XORed
		'against the character in front of it.
		For X = 1 To iLen - 1
			Checksum1 = Checksum1 + CShort(iByte(X) Xor iByte(X + 1))
		Next X
		Checksum1 = Checksum1 Mod 256
		oByte(pLen - 3) = CByte(Checksum1)
		
		'The second is an addition of the each character in the encrypted
		'string XORed against a its position in the string.
		For X = 1 To pLen - 3
			Checksum2 = Checksum2 + CShort(oByte(X) Xor (256 - X))
		Next X
		Checksum2 = (Checksum2 + MainKey) Mod 256
		oByte(pLen - 2) = CByte(Checksum2)
		
		'And the third is too complicated to explain =p
		V = pLen \ 4
		For X = 1 To V
			Working = X
			For Y = 0 To 3
				Z = Y * V + X
				If Z <> pLen - 1 Then
					Working = Working + oByte(Z) * (CShort(Y Mod 2 = 1) * 2 + 1) + 7
				End If
			Next Y
			Checksum3 = Checksum3 + Working
		Next X
		Checksum3 = System.Math.Abs(Int(Checksum3 + (oByte((MainKey Mod iLen) + 1)))) Mod 256
		oByte(pLen - 1) = CByte(Checksum3)
		
		'And the finishing touch, the string is scrambled
		For Z = 1 To 2
			tByte = VB6.CopyArray(oByte)
			Y = 0
			For X = 1 To pLen - 1 Step 2
				Y = Y + 1
				oByte(X) = tByte(Y)
			Next X
			Y = pLen + 1
			For X = 2 To pLen Step 2
				Y = Y - 1
				oByte(X) = tByte(Y)
			Next X
		Next Z
		
		'Finally, the string is reformed from the array, with one final ajustment
		Build = New String(vbNullChar, pLen)
		'    For X = 1 To pLen
		'        Mid$(Build, X, 1) = Chr$(oByte(X))
		'    Next X
		CopyMemory(Build, oByte(1), pLen)
		Build = StrReverse(Left(Build, pLen \ 2)) & StrReverse(Right(Build, pLen \ 2))
		Select Case Left(Build, 5)
			Case "REQN:", "RPWD:", "NAME:", "SPWD:", "BANU:", "NOIP:" 'Yes, yes, it's a billion to one chance I know, but just in case...
				XOREncrypt = XOREncrypt(Orig)
			Case Else
				XOREncrypt = Build
		End Select
		Exit Function
ETCTrap: 
		If Err.Number = 16 Then Resume  Else Err.Raise(Err.Number)
	End Function
	
	Public Function XORDecrypt(ByVal EncryptedString As String, Optional ByRef SkipChecksums As Boolean = False) As String
		Dim MainKey As Byte
		Dim iByte() As Byte
		Dim oByte() As Byte
		Dim tByte() As Byte
		Dim Build As String
		Dim iLen As Short
		Dim pLen As Short
		Dim Checksum1 As Integer
		Dim Checksum2 As Integer
		Dim Checksum3 As Integer
		Dim Working As Integer
		Dim V As Short
		Dim W As Short
		Dim X As Short
		Dim Y As Short
		Dim Z As Short
		Build = EncryptedString
		pLen = Len(Build)
		If pLen Mod 2 = 1 Then
			'>>> Call WriteDebugLog("DECRYPT ERROR: Odd length packet rcvd")
			'>>> Call WriteDebugLog("HEXED ESTRING: " & Str2Hex(Build))
			XORDecrypt = "XXXXX"
			Exit Function
		End If
		If Len(Build) < 5 Then
			'>>> Call WriteDebugLog("DECRYPT ERROR: Len(Build) < 5")
			'>>> Call WriteDebugLog("HEXED ESTRING: " & Str2Hex(Build))
			XORDecrypt = "XXXXX"
			Exit Function
		End If
		iLen = pLen - 4
		'UPGRADE_WARNING: Lower bound of array iByte was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iByte(pLen)
		'UPGRADE_WARNING: Lower bound of array oByte was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim oByte(pLen)
		Build = StrReverse(Left(Build, pLen \ 2)) & StrReverse(Right(Build, pLen \ 2))
		CopyMemory(iByte(1), Build, Len(Build))
		
		For Z = 1 To 2
			Y = 0
			For X = 1 To pLen - 1 Step 2
				Y = Y + 1
				oByte(Y) = iByte(X)
			Next X
			Y = pLen + 1
			For X = 2 To pLen Step 2
				Y = Y - 1
				oByte(Y) = iByte(X)
			Next X
			iByte = VB6.CopyArray(oByte)
		Next Z
		
		MainKey = oByte(pLen) Xor 12
		For X = iLen To 1 Step -1
			oByte(X) = oByte(X) Xor oByte(iLen - X + 1) Xor MainKey
		Next X
		
		Checksum1 = 0
		Checksum2 = 0
		Checksum3 = 0
		
		For X = 1 To iLen - 1
			Checksum1 = Checksum1 + CShort(oByte(X) Xor oByte(X + 1))
		Next X
		Checksum1 = Checksum1 Mod 256
		
		For X = 1 To pLen - 3
			Checksum2 = Checksum2 + CShort(iByte(X) Xor (256 - X))
		Next X
		Checksum2 = (Checksum2 + MainKey) Mod 256
		
		V = pLen \ 4
		For X = 1 To V
			Working = X
			For Y = 0 To 3
				Z = Y * V + X
				If Z <> pLen - 1 Then
					Working = Working + iByte(Z) * (CShort(Y Mod 2 = 1) * 2 + 1) + 7
				End If
			Next Y
			Checksum3 = Checksum3 + Working
		Next X
		Checksum3 = System.Math.Abs(Int(Checksum3 + (iByte((MainKey Mod iLen) + 1)))) Mod 256
		
		X = 0
		If iByte(pLen - 3) <> Checksum1 Then X = 1
		If iByte(pLen - 2) <> Checksum2 Then X = 1
		If iByte(pLen - 1) <> Checksum3 Then X = 1
		If X = 0 Or SkipChecksums Then
			Build = New String(vbNullChar, iLen)
			For X = 1 To iLen
				Mid(Build, X, 1) = Chr(oByte(X))
			Next X
			If Mid(Build, 1, 1) = vbNullChar Then Build = Right(Build, Len(Build) - 1)
		Else
			'>>> Call WriteDebugLog("DECRYPT ERROR: Checksum mismatch.")
			'>>> Call WriteDebugLog("HEXED ESTRING: " & Str2Hex(EncryptedString))
			'        If iByte(pLen - 3) <> Checksum1 Then '>>> Call WriteDebugLog("CHECK1 FAIL")
			'        If iByte(pLen - 2) <> Checksum2 Then '>>> Call WriteDebugLog("CHECK2 FAIL")
			'        If iByte(pLen - 1) <> Checksum3 Then '>>> Call WriteDebugLog("CHECK3 FAIL")
			Build = "XXXXX"
		End If
		XORDecrypt = Build
	End Function
	Public Sub Benchmark(ByRef Seconds As Short)
		Dim X As Short
		Dim Y As Short
		Dim Z As Integer
		Dim T As Single
		Dim E As String
		T = VB.Timer() + Seconds
		While VB.Timer() < T
			Y = Int(Rnd() * 250) + 2
			E = RndStr(Y)
			If E <> XORDecrypt(XOREncrypt(E)) Then Stop
			Z = Z + Y
		End While
		MsgBox(CStr(System.Math.Round(Z / 1024 / Seconds, 2)) & " kilobytes per second.", MsgBoxStyle.Information, "Result")
	End Sub
	Public Function RndStr(ByRef StrLen As Short) As String
		Dim X As Short
		Dim B As Byte
		Dim Build As String
		Build = New String(vbNullChar, StrLen)
		For X = 1 To StrLen
			Do 
				B = Int(Rnd() * 256)
			Loop Until X <> 1 Or B <> 0
			Mid(Build, X) = Chr(B)
		Next X
		RndStr = Build
	End Function
	'Public Function Str2Hex(Orig As String) As String
	'    Dim X As Integer
	'    Dim Temp As String
	'    Temp = String$(Len(Orig) * 3, vbNullChar)
	'    For X = 1 To Len(Orig)
	'        Mid(Temp, X * 3 - 2) = FixedHex(Asc(Mid(Orig, X, 1)), 2)
	'    Next X
	'    Str2Hex = Temp
	'End Function
	'Public Function Hex2Str(ByVal Hexed As String) As String
	'    Dim H() As String
	'    Dim X As Integer
	'    Dim Temp As String
	'    If Len(Hexed) > 2 And Mid(Hexed, 3, 1) <> " " Then
	'        For X = 1 To Len(Hexed) Step 2
	'            Temp = Temp & Mid(Hexed, X, 2) & " "
	'        Next X
	'        Hexed = Temp
	'    End If
	'    H = Split(Trim(Hexed), " ")
	'    For X = 0 To UBound(H)
	'        H(X) = Chr$(Dec(H(X)))
	'    Next X
	'    Hex2Str = Join(H, vbNullString)
	'End Function
	Public Function Rebuild(ByRef ByteArray() As Byte) As String
		Dim X As Short
		Dim Temp As String
		For X = LBound(ByteArray) To UBound(ByteArray)
			Temp = Temp & Chr(ByteArray(X))
		Next X
		Rebuild = Temp
		
	End Function
	
	Public Function Deconstruct(ByRef TheString As String) As Byte()
		Dim X As Integer
		Dim Temp As String
		Dim B() As Byte
		ReDim B(Len(TheString) - 1)
		For X = 0 To Len(TheString) - 1
			B(X) = Asc(Mid(TheString, X + 1, 1))
		Next X
		Deconstruct = VB6.CopyArray(B)
	End Function
End Module