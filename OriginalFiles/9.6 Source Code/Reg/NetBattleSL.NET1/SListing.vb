Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class MSListing
	Inherits System.Windows.Forms.Form
	Const VERSION As String = "0.9.4"
	Const MAXSERVERS As Short = 50
	Const MAXCLIENTS As Short = 25
	Private Structure PingType
		Dim Chances As Short
		Dim Pongs As Short
		Dim SentPing As Boolean
	End Structure
	Private Structure QueueType2
		Dim iData() As String
	End Structure
	Private Structure QueueType1
		Dim Num() As QueueType2
	End Structure
	Private Structure ServerType
		Dim Address As String
		Dim Admin As String
		Dim ServerName As String
		Dim Users As Short
		Dim MaxUsers As Short
		Dim Active As Boolean
		Dim Description As String
		Dim SentPing As Boolean
		Dim Pongs As Short
		Dim Info As String
		Dim IPChangeable As Boolean
		Dim ClientsKnow As Boolean
		Dim InfoChanges As Short
		Dim UserChanges As Short
		Dim DisconnectMe As Short
		Dim SID As String
		Dim Regged As Boolean
	End Structure
	Private Structure ServerDBEntry
		Dim ServerName As String
		Dim SID As String
		Dim Pass As String
	End Structure
	
	'UPGRADE_WARNING: Lower bound of array Queue was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private Queue(4) As QueueType1
	'Usage:
	'1=SvrSnd, 2=SvrRcv, 3=ClntSnd, 4=ClntRcv
	'Queue(1).Num(3).iData(1) stores the
	'next thing to be sent to Server #3.
	'Queue(4).Num(1).iData(2) is the second
	'thing in line to be recieved from Client #1.
	Private SDiscon(MAXSERVERS) As Boolean
	Private CDiscon(MAXCLIENTS) As Boolean
	Private Server(MAXSERVERS) As ServerType
	Private CPing(MAXCLIENTS) As PingType
	Private IPBan() As String
	Private AttemptLog() As String
	Private SlashPath As String
	Private Entry() As ServerDBEntry
	
	
	Private Sub CSendAll(ByVal SendMe As String)
		Dim X As Short
		For X = 1 To ClientSocket.UBound
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ClientSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected Then
				Call AddToQueue(3, X, SendMe)
			End If
		Next X
	End Sub
	
	Private Sub ChannelScanner_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChannelScanner.Tick
		Dim X As Short
		Dim Temp As String
		On Error Resume Next
		For X = 1 To MAXSERVERS
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ServerSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected Then
				If Server(X).SentPing Then Call DisconnectPlayer(True, X)
				Call AddToQueue(1, X, "PING:")
				Server(X).SentPing = True
				Server(X).Pongs = 0
				Server(X).InfoChanges = 0
				Server(X).UserChanges = 0
			Else
				If Server(X).Active Then Call DisconnectPlayer(True, X)
			End If
		Next X
		For X = 1 To MAXCLIENTS
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ClientSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected Then
				If CPing(X).SentPing Then CPing(X).Chances = CPing(X).Chances + 1
				If CPing(X).Chances = 3 Then Call DisconnectPlayer(False, X)
				Call AddToQueue(3, X, "PING:")
				CPing(X).SentPing = True
				CPing(X).Pongs = 0
			End If
		Next X
		ReDim AttemptLog(0)
		For X = 1 To UBound(IPBan)
			Temp = Mid(IPBan(X), 1, 1)
			If Temp = "F" Then
				IPBan(X) = ""
			Else
				Mid(IPBan(X), 1, 1) = Hex(Val("&H" & Temp) + 1)
			End If
		Next X
		For X = UBound(IPBan) To 1 Step -1
			If IPBan(X) = "" Then ReDim Preserve IPBan(X - 1) Else Exit For
		Next X
		
		On Error GoTo FileOpen_Renamed
		Kill(SlashPath & "servers.csv")
		FileOpen(1, SlashPath & "servers.csv", OpenMode.Output)
		For X = 1 To UBound(Entry)
			With Entry(X)
				WriteLine(1, .ServerName, .SID, .Pass)
			End With
		Next X
		FileClose(1)
FileOpen_Renamed: 
	End Sub
	
	Private Sub ClientSocket_CloseEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ClientSocket.CloseEvent
		Dim Index As Short = ClientSocket.GetIndex(eventSender)
		Call DisconnectPlayer(False, Index)
	End Sub
	
	Private Sub ClientSocket_ConnectionRequest(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ConnectionRequestEvent) Handles ClientSocket.ConnectionRequest
		Dim Index As Short = ClientSocket.GetIndex(eventSender)
		Dim X As Short
		Dim Y As Short
		Dim Temp As String
		On Error Resume Next
		For X = 1 To ClientSocket.UBound
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ClientSocket(0).RemoteHostIP = ClientSocket(X).RemoteHostIP And ClientSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected Then Temp = "Repeat"
		Next X
		If Temp = "" Then
			Temp = ShortenIP(ClientSocket(0).RemoteHostIP)
			For X = 1 To UBound(IPBan)
				If Mid(IPBan(X), 2, 8) = Temp Then Exit Sub
			Next X
			For X = 1 To UBound(AttemptLog)
				If Mid(AttemptLog(X), 2, 8) = Temp Then
					If Mid(AttemptLog(X), 1, 1) = "F" Then
						Call TempBan(Temp)
						Temp = "Ban"
					Else
						Mid(AttemptLog(X), 1, 1) = Hex(Val("&H" & Mid(AttemptLog(X), 1, 1)) + 1)
					End If
					Exit For
				End If
			Next X
		End If
		If X = UBound(AttemptLog) + 1 Then
			ReDim Preserve AttemptLog(X)
			AttemptLog(X) = "0" & Temp
		End If
		For X = 1 To MAXCLIENTS
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ClientSocket(X).CtlState <> MSWinsockLib.StateConstants.sckConnected Then Exit For
		Next X
		If X = MAXCLIENTS + 1 Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object CPing(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CPing(X) = CPing(0)
		CDiscon(X) = False
		ReDim Queue(3).Num(X).iData(0)
		ReDim Queue(4).Num(X).iData(0)
		If Temp = "Ban" Then
			Call TempBan(ShortenIP(ClientSocket(X).RemoteHostIP))
		ElseIf Temp = "Repeat" Then 
			Exit Sub
		Else
			ClientSocket(X).Close()
			ClientSocket(X).Accept(eventArgs.requestID)
			For Y = 1 To MAXSERVERS
				If Server(Y).Active Then Call AddToQueue(3, X, "SERV:" & Server(Y).Info)
			Next Y
		End If
		Call UpdateClientCount()
	End Sub
	
	
	Private Sub ClientSocket_Error(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent) Handles ClientSocket.Error
		Dim Index As Short = ClientSocket.GetIndex(eventSender)
		Call DisconnectPlayer(False, Index)
	End Sub
	
	Private Sub cmdDisconnect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDisconnect.Click
		Dim Temp As String
		If ListDisplay.Items.Count = 0 Then Exit Sub
		Temp = ListDisplay.FocusedItem.Name
		Call DisconnectPlayer(True, Val(VB.Right(Temp, Len(Temp) - 5)))
	End Sub
	
	Private Sub cmdMassMsg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMassMsg.Click
		Dim Temp As String
		Dim X As Short
		Temp = InputBox("Please enter a message.  This message will be sent to all players on all connected servers.", "Mass Message")
		If Temp = "" Then Exit Sub
		For X = 1 To ServerSocket.UBound
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ServerSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected Then
				Call AddToQueue(1, X, "MASS:" & Temp)
			End If
		Next X
	End Sub
	
	Private Sub KickTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles KickTimer.Tick
		Dim X As Short
		For X = 1 To MAXSERVERS
			If Server(X).DisconnectMe <> 0 Then
				Server(X).DisconnectMe = Server(X).DisconnectMe - 1
				If Server(X).DisconnectMe = 0 Then Call DisconnectPlayer(True, X)
			End If
		Next X
	End Sub
	
	Private Sub ListDisplay_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ListDisplay.Click
		Dim X As Short
		On Error Resume Next
		X = Val(VB.Right(ListDisplay.FocusedItem.Name, Len(ListDisplay.FocusedItem.Name) - 5))
		Label1.Text = Server(X).Description
	End Sub
	
	'UPGRADE_ISSUE: MSComctlLib.ListView event ListDisplay.ItemClick was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub ListDisplay_ItemClick(ByVal Item As System.Windows.Forms.ListViewItem)
		Label1.Text = Server(Val(VB.Right(Item.Name, Len(Item.Name) - 5))).Description
	End Sub
	
	Private Sub ServerSocket_CloseEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ServerSocket.CloseEvent
		Dim Index As Short = ServerSocket.GetIndex(eventSender)
		Call DisconnectPlayer(True, Index)
	End Sub
	
	Private Sub ServerSocket_ConnectionRequest(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ConnectionRequestEvent) Handles ServerSocket.ConnectionRequest
		Dim Index As Short = ServerSocket.GetIndex(eventSender)
		Dim X As Short
		Dim Y As Short
		Dim Temp As String
		On Error Resume Next
		For X = 1 To MAXSERVERS
			If ServerSocket(0).RemoteHostIP = Server(X).Address Then Temp = "Repeat"
		Next X
		If Temp = "" Then
			Temp = ShortenIP(ServerSocket(0).RemoteHostIP)
			For X = 1 To UBound(IPBan)
				If IPBan(X) = Temp Then Exit Sub
			Next X
			For X = 1 To UBound(AttemptLog)
				If Mid(AttemptLog(X), 2, 8) = Temp Then
					If Mid(AttemptLog(X), 1, 1) = "F" Then
						Call TempBan(Temp)
						Temp = "Ban"
					Else
						Mid(AttemptLog(X), 1, 1) = Hex(Val("&H" & Mid(AttemptLog(X), 1, 1)) + 1)
					End If
				End If
			Next X
		End If
		If X = UBound(AttemptLog) + 1 Then
			ReDim Preserve AttemptLog(X)
			AttemptLog(X) = "0" & Temp
		End If
		For X = 1 To MAXSERVERS
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ServerSocket(X).CtlState <> MSWinsockLib.StateConstants.sckConnected Then Exit For
		Next X
		If X = MAXSERVERS + 1 Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object Server(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Server(X) = Server(0)
		ReDim Queue(1).Num(X).iData(0)
		ReDim Queue(2).Num(X).iData(0)
		ServerSocket(X).Close()
		ServerSocket(X).Accept(eventArgs.requestID)
		If Temp = "Ban" Then
			Call TempBan(ShortenIP(ServerSocket(0).RemoteHostIP))
			Call AddToQueue(1, X, "TBAN:2")
		ElseIf Temp = "Repeat" Then 
			Call AddToQueue(1, X, "MULTI")
		Else
			Server(X).Active = True
			Call AddToQueue(1, X, "RINF:")
		End If
	End Sub
	
	Private Sub ServerSocket_DataArrival(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles ServerSocket.DataArrival
		Dim Index As Short = ServerSocket.GetIndex(eventSender)
		Dim Worked As Boolean
		Dim Packet() As String
		Dim X As Short
		Worked = GetPacket(ServerSocket(Index), eventArgs.BytesTotal, Packet)
		If Worked Then
			For X = 1 To UBound(Packet)
				Call AddToQueue(2, Index, Packet(X))
			Next X
		Else
			Call DisconnectPlayer(True, Index)
		End If
	End Sub
	Private Sub ClientSocket_DataArrival(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles ClientSocket.DataArrival
		Dim Index As Short = ClientSocket.GetIndex(eventSender)
		Dim Worked As Boolean
		Dim Packet() As String
		Dim X As Short
		Worked = GetPacket(ClientSocket(Index), eventArgs.BytesTotal, Packet)
		If Worked Then
			For X = 0 To UBound(Packet)
				Call AddToQueue(4, Index, Packet(X))
			Next X
		Else
			Call DisconnectPlayer(False, Index)
		End If
	End Sub
	
	
	Private Sub MSListing_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim TempVar As String
		Dim X As Short
		Dim Y As Short
		ReDim IPBan(0)
		ReDim AttemptLog(0)
		For X = 1 To 2
			ReDim Queue(X).Num(MAXSERVERS)
			For Y = 0 To MAXSERVERS
				ReDim Queue(X).Num(Y).iData(0)
			Next Y
		Next X
		For X = 3 To 4
			ReDim Queue(X).Num(MAXCLIENTS)
			For Y = 0 To MAXCLIENTS
				ReDim Queue(X).Num(Y).iData(0)
			Next Y
		Next X
		For X = 1 To MAXSERVERS
			ServerSocket.Load(X)
		Next X
		For X = 1 To MAXCLIENTS
			ClientSocket.Load(X)
		Next X
		ServerSocket(0).Listen()
		ClientSocket(0).Listen()
		
		SlashPath = My.Application.Info.DirectoryPath
		ReDim Entry(0)
		If VB.Right(SlashPath, 1) <> "\" Then SlashPath = SlashPath & "\"
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(SlashPath & "servers.csv")) = 0 Then
			FileOpen(1, SlashPath & "servers.csv", OpenMode.Output)
			FileClose(1)
		Else
			X = 0
			FileOpen(1, SlashPath & "servers.csv", OpenMode.Input)
			Do Until EOF(1)
				X = X + 1
				ReDim Preserve Entry(X)
				With Entry(X)
					Input(1, .ServerName)
					Input(1, .SID)
					Input(1, .Pass)
				End With
			Loop 
			FileClose(1)
		End If
	End Sub
	
	Private Sub AddToQueue(ByVal QNum As Short, ByVal Number As Short, ByVal QData As String)
		Dim X As Short
		On Error GoTo BadNum
		If QNum Mod 2 = 0 Then QData = XORDecrypt(QData)
		X = UBound(Queue(QNum).Num(Number).iData) + 1
		ReDim Preserve Queue(QNum).Num(Number).iData(X)
		Queue(QNum).Num(Number).iData(X) = QData
		Exit Sub
BadNum: 
		Call DisconnectPlayer(True, Number)
	End Sub
	
	Private Sub QueueTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles QueueTimer.Tick
		Dim X As Short
		Dim Y As Short
		Dim Z As Short
		Dim Temp As String
		For X = 1 To 4
			For Y = 1 To UBound(Queue(X).Num)
				If UBound(Queue(X).Num(Y).iData) <> 0 Then
					Temp = Queue(X).Num(Y).iData(1)
					For Z = 2 To UBound(Queue(X).Num(Y).iData)
						Queue(X).Num(Y).iData(Z - 1) = Queue(X).Num(Y).iData(Z)
					Next Z
					ReDim Preserve Queue(X).Num(Y).iData(Z - 2)
					Select Case X
						Case 1 : Call SendData(True, Y, Temp)
						Case 2 : Call DoIncoming(True, Y, Temp)
						Case 3 : Call SendData(False, Y, Temp)
						Case 4 : Call DoIncoming(False, Y, Temp)
					End Select
				End If
			Next Y
		Next X
	End Sub
	Private Sub SendData(ByVal ToServer As Boolean, ByVal Index As Short, ByVal SendMe As String)
		On Error GoTo ErrorTrap
		Dim XORSendMe As String
		XORSendMe = FormatPacket(SendMe, True)
		If ToServer Then
			If Index > ServerSocket.UBound Then Exit Sub
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ServerSocket(Index).CtlState <> MSWinsockLib.StateConstants.sckConnected Then Exit Sub
			ServerSocket(Index).SendData(XORSendMe)
			If SendMe = "TERR:" Or VB.Left(SendMe, 5) = "TBAN:" Or VB.Left(SendMe, 5) = "MULTI" Or VB.Left(SendMe, 5) = "OLDVR" Or SendMe = "WRONG" Then Server(Index).DisconnectMe = 5
		Else
			If Index > ClientSocket.UBound Then Exit Sub
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ClientSocket(Index).CtlState <> MSWinsockLib.StateConstants.sckConnected Then Exit Sub
			ClientSocket(Index).SendData(XORSendMe)
		End If
		Exit Sub
ErrorTrap: 
		Call DisconnectPlayer(ToServer, Index)
	End Sub
	Private Sub DoIncoming(ByVal FromServer As Boolean, ByVal Index As Short, ByVal Info As String)
		Dim X As Short
		Dim P1 As Short
		Dim P2 As Short
		Dim Temp As String
		Dim Temp2 As String
		Dim Temp3 As String
		Dim TempVar As Object
		Temp = VB.Right(Info, Len(Info) - 5)
		If FromServer Then
			Server(Index).Active = True
			Select Case VB.Left(Info, 5)
				Case "INFO:" 'new server INFO
					If Len(Temp) < 44 Or Server(Index).ClientsKnow Then Call DisconnectPlayer(True, Index)
					For X = 1 To ServerSocket.UBound
						'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						If ServerSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected And Server(X).ClientsKnow Then
							If Server(X).ServerName = Trim(Mid(Temp, 1, 20)) Then
								Call AddToQueue(1, Index, "NMIU:")
								Exit Sub
							End If
						End If
					Next X
					Server(Index).ServerName = Trim(Mid(Temp, 1, 20))
					Server(Index).Admin = Trim(Mid(Temp, 21, 20))
					Server(Index).Users = Asc(Mid(Temp, 41, 1))
					Server(Index).MaxUsers = Asc(Mid(Temp, 42, 1))
					Server(Index).SID = DecompressSID(Mid(Temp, 51, 13)) '& "T"
					Server(Index).Address = ServerSocket(Index).RemoteHostIP
					Server(Index).Description = VB.Right(Temp, Len(Temp) - 63)
					If Trim(Mid(Temp, 43, 8)) <> VERSION Then
						Call AddToQueue(1, Index, "OLDVR")
						Exit Sub
					End If
					For X = 1 To UBound(Entry)
						If Entry(X).ServerName = Server(Index).ServerName And Entry(X).SID <> Server(Index).SID Then
							Call AddToQueue(1, Index, "NAMPW")
							Exit Sub
						End If
					Next X
					For X = 1 To UBound(Entry)
						If Entry(X).SID = Server(Index).SID Then
							Entry(X).ServerName = Server(Index).ServerName
							Server(Index).Regged = True
						End If
					Next X
					
					Call SetInfo(Index)
					Call RefreshListing()
					'                If Mid(Server(Index).Address, 1, 7) = "192.168" Then
					'                    Server(Index).IPChangeable = True
					'                    Call AddToQueue(1, Index, "RQIP:")
					'                    Exit Sub
					'                Else
					Server(Index).IPChangeable = False
					'End If
					Call CSendAll("SERV:" & Server(Index).Info)
					Server(Index).ClientsKnow = True
					If Server(Index).Regged Then
						Call AddToQueue(1, Index, "OKAY!R")
					Else
						Call AddToQueue(1, Index, "OKAY!")
					End If
				Case "RLIP:"
					If Not Server(Index).IPChangeable Then
						Call DisconnectPlayer(True, Index)
						Exit Sub
					End If
					Call ChangeVar(Server(Index).Address, Index, Temp)
					Server(Index).ClientsKnow = True
					Call AddToQueue(1, Index, "OKAY!")
				Case "NAMC:" 'NAMe Change
					Call ChangeVar(Server(Index).ServerName, Index, Temp)
				Case "ADMC:" 'ADMin Change
					Call ChangeVar(Server(Index).Admin, Index, Temp)
				Case "DESC:" 'DEScription Change
					Call ChangeVar(Server(Index).Description, Index, Temp)
				Case "MAXC:" 'MAX user Change
					Call ChangeVar(Server(Index).MaxUsers, Index, Temp)
				Case "USRC:" 'USeR Change
					If Server(Index).Users = CDbl(Temp) Or Server(Index).ServerName = "" Then Exit Sub
					Server(Index).Users = CShort(Temp)
					Call SetInfo(Index)
					Call RefreshListing()
					Call CSendAll("SERV:" & Server(Index).Info)
					Server(Index).UserChanges = Server(Index).UserChanges + 1
					If Server(Index).UserChanges > 13 Then
						Call TempBan(ShortenIP(Server(Index).Address))
						Call AddToQueue(1, Index, "TBAN:0")
					End If
				Case "PASS:" 'Password
					For X = 1 To ServerSocket.UBound
						'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						If ServerSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected And Server(X).ClientsKnow Then
							If Server(X).ServerName = Server(Index).ServerName Then
								Call DisconnectPlayer(True, Index)
								Exit Sub
							End If
						End If
					Next X
					
					For X = 1 To UBound(Entry)
						If Entry(X).ServerName = Server(Index).ServerName Then
							If Temp = Entry(X).Pass Then
								Entry(X).SID = Server(Index).SID
								Call AddToQueue(1, Index, "RIGHT")
								Server(Index).Regged = True
								If Not Server(Index).ClientsKnow Then
									Call SetInfo(Index)
									Call RefreshListing()
									Server(Index).IPChangeable = False
									Call CSendAll("SERV:" & Server(Index).Info)
									Server(Index).ClientsKnow = True
									Call AddToQueue(1, Index, "OKAY!R")
								End If
							Else
								Call AddToQueue(1, Index, "WRONG")
							End If
							Exit For
						End If
					Next X
					If X > UBound(Entry) Then
						ReDim Preserve Entry(X)
						With Entry(X)
							.Pass = Temp
							.ServerName = Server(Index).ServerName
							.SID = Server(Index).SID
						End With
						Call AddToQueue(1, Index, "REGED")
					End If
				Case "PONG:"
					Server(Index).SentPing = False
					Server(Index).Pongs = Server(Index).Pongs + 1
					If Server(Index).Pongs > 4 Then
						Call TempBan(ShortenIP(Server(Index).Address))
						Call AddToQueue(1, Index, "TBAN:0")
					End If
				Case Else
					Call DisconnectPlayer(True, Index)
			End Select
		Else
			Select Case VB.Left(Info, 5)
				Case "PONG:"
					CPing(Index).Pongs = CPing(Index).Pongs + 1
					If CPing(Index).Pongs > 4 Then Call DisconnectPlayer(False, Index)
					CPing(Index).SentPing = False
					CPing(Index).Chances = 0
				Case Else
					Call DisconnectPlayer(False, Index)
			End Select
		End If
	End Sub
	Private Sub DisconnectPlayer(ByVal IsServer As Boolean, ByVal Number As Short)
		Dim X As Short
		'On Error Resume Next
		If IsServer Then
			If SDiscon(Number) Then Exit Sub
			SDiscon(Number) = True
			If Server(Number).ClientsKnow Then
				Call CSendAll("DISC:" & ShortenIP(Server(Number).Address))
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Server(Number). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Server(Number) = Server(0)
			ReDim Queue(1).Num(Number).iData(0)
			ReDim Queue(2).Num(Number).iData(0)
			ServerSocket(Number).Close()
			SDiscon(Number) = False
			Call RefreshListing()
		Else
			If ClientSocket.UBound < Number Then Exit Sub
			If CDiscon(Number) Then Exit Sub
			CDiscon(Number) = True
			'UPGRADE_WARNING: Couldn't resolve default property of object CPing(Number). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CPing(Number) = CPing(0)
			ReDim Queue(3).Num(Number).iData(0)
			ReDim Queue(4).Num(Number).iData(0)
			ClientSocket(Number).Close()
			CDiscon(Number) = False
		End If
		Call UpdateClientCount()
	End Sub
	Private Sub RefreshListing()
		Dim X As Short
		Dim TempItem As System.Windows.Forms.ListViewItem
		ListDisplay.Items.Clear()
		'UPGRADE_ISSUE: MSComctlLib.ListView property ListDisplay.Sorted was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		ListDisplay.Sorted = False
		For X = 1 To UBound(Server)
			If Server(X).Active Then
				TempItem = ListDisplay.Items.Add("SRVR:" & X, CStr(X), "")
				'UPGRADE_WARNING: Lower bound of collection TempItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				If TempItem.SubItems.Count > 1 Then
					TempItem.SubItems(1).Text = Server(X).ServerName
				Else
					TempItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Server(X).ServerName))
				End If
				'UPGRADE_WARNING: Lower bound of collection TempItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				If TempItem.SubItems.Count > 2 Then
					TempItem.SubItems(2).Text = Server(X).Address
				Else
					TempItem.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Server(X).Address))
				End If
				'UPGRADE_WARNING: Lower bound of collection TempItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				If TempItem.SubItems.Count > 3 Then
					TempItem.SubItems(3).Text = Server(X).Admin
				Else
					TempItem.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Server(X).Admin))
				End If
				'UPGRADE_WARNING: Lower bound of collection TempItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				If TempItem.SubItems.Count > 4 Then
					TempItem.SubItems(4).Text = Server(X).Users & "/" & Server(X).MaxUsers
				Else
					TempItem.SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Server(X).Users & "/" & Server(X).MaxUsers))
				End If
			End If
		Next X
		ListDisplay.Sort()
		Call ListDisplay_Click(ListDisplay, New System.EventArgs())
	End Sub
	Private Function ShortenIP(ByVal IP As String) As String
		Dim Temp As String
		Dim Temp2 As String
		Dim X As Short
		On Error GoTo NoIP
		Temp = ""
		Temp2 = ""
		For X = 1 To Len(IP)
			If Mid(IP, X, 1) = "." Then
				'UPGRADE_WARNING: Couldn't resolve default property of object FHex(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Temp2 = Temp2 & FHex(CShort(Temp))
				Temp = ""
			Else
				Temp = Temp & Mid(IP, X, 1)
			End If
		Next X
		'UPGRADE_WARNING: Couldn't resolve default property of object FHex(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Temp2 = Temp2 & FHex(CShort(Temp))
		If Len(Temp2) <> 8 Then GoTo NoIP
		ShortenIP = Temp2
		Exit Function
NoIP: 
		ShortenIP = "00000000"
	End Function
	Private Function FHex(ByVal Number As Short, Optional ByVal Digits As Short = 2) As Object
		Dim Temp As String
		Temp = Hex(Number)
		If Len(Temp) >= Digits Then
			'UPGRADE_WARNING: Couldn't resolve default property of object FHex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FHex = Temp
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object FHex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FHex = New String("0", Digits - Len(Temp)) & Temp
		End If
	End Function
	Private Sub SetInfo(ByVal Index As Short)
		With Server(Index)
			'UPGRADE_WARNING: Couldn't resolve default property of object FHex(Server(Index).MaxUsers). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FHex(Server(Index).Users). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FHex(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Info = FHex(Index) & .ServerName & New String(" ", 20 - Len(.ServerName)) & .Admin & New String(" ", 20 - Len(.Admin)) & FHex(.Users) & FHex(.MaxUsers) & ShortenIP(.Address) & .Description
		End With
	End Sub
	Private Sub ChangeVar(ByRef Var As Object, ByRef Index As Short, ByRef NewVal As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object NewVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Var. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Var = NewVal Or Server(Index).ServerName = "" Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object NewVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Var. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Var = NewVal
		Call SetInfo(Index)
		Call RefreshListing()
		Call CSendAll("SERV:" & Server(Index).Info)
		Server(Index).InfoChanges = Server(Index).InfoChanges + 1
		If Server(Index).InfoChanges = 10 Then Call AddToQueue(1, Index, "WARN:")
		If Server(Index).InfoChanges = 15 Then
			Call TempBan(ShortenIP(Server(Index).Address))
			Call AddToQueue(1, Index, "TBAN:2")
		End If
	End Sub
	Private Sub TempBan(ByVal iAddress As String)
		Dim X As Short
		X = UBound(IPBan) + 1
		ReDim Preserve IPBan(X)
		IPBan(X) = "0" & iAddress
	End Sub
	Private Sub UpdateClientCount()
		Dim X As Short
		Dim Y As Short
		For X = 0 To MAXCLIENTS
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If ClientSocket(X).CtlState = MSWinsockLib.StateConstants.sckConnected Then Y = Y + 1
		Next X
		'UPGRADE_WARNING: Lower bound of collection StatusBar1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		StatusBar1.Items.Item(3).Text = "Clients: " & Y
	End Sub
	
	Private Sub ServerSocket_Error(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent) Handles ServerSocket.Error
		Dim Index As Short = ServerSocket.GetIndex(eventSender)
		DisconnectPlayer(True, Index)
	End Sub
	
	Public Function ChopString(ByRef Source As String, ByVal Count As Short) As String
		Dim Temp As String
		If Source = "" Then Exit Function
		Temp = VB.Left(Source, Count)
		Source = VB.Right(Source, Len(Source) - Count)
		ChopString = Temp
	End Function
	
	Public Function DecompressSID(ByRef SID As String, Optional ByRef Fake As Boolean = False) As String
		Dim Build As String
		Dim Temp As String
		Dim X As Integer
		Dim Y As Integer
		Build = VB.Left(Chr2Bin(SID), 100)
		For X = 1 To 5
			Temp = Temp & Mid(Build, X * 20, 1)
		Next X
		Build = Temp & Build
		For X = 1 To 21
			Y = Bin2Dec(Mid(Build, X * 5 - 4, 5))
			Y = Y + IIf(Y > 8, 56, 49)
			Mid(Build, X, 1) = Chr(Y)
		Next X
		If Fake Then Mid(Build, 1, 1) = "Y"
		DecompressSID = VB.Left(Build, 21)
	End Function
	Public Function Chr2Bin(ByVal ChrString As String) As String
		Dim Build As String
		Dim X As Integer
		'Reverse of the above
		Build = New String(vbNullChar, Len(ChrString) * 8)
		For X = 1 To Len(ChrString)
			Mid(Build, X * 8 - 7) = Dec2Bin(Asc(Mid(ChrString, X, 1)), 8)
		Next X
		Chr2Bin = Build
	End Function
	Public Function Bin2Dec(ByVal BitString As String) As Integer
		Dim X As Integer
		Static T() As Short
		If BitString = vbNullString Then Exit Function
		ReDim T(Len(BitString) - 1)
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		CopyMemory(T(0), StrPtr(BitString), LenB(BitString))
		Bin2Dec = T(0) - System.Windows.Forms.Keys.D0
		For X = 1 To UBound(T)
			Bin2Dec = Bin2Dec + Bin2Dec + T(X) - System.Windows.Forms.Keys.D0
		Next X
	End Function
	Public Function Dec2Bin(ByVal X As Integer, ByVal Fixed As Short) As String
		Static lDone As Integer
		'UPGRADE_NOTE: sByte was upgraded to sByte_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Static sByte_Renamed(255) As String
		Dim sNibble(15) As String
		Dim Y As Integer
		'If Sgn(X) = -1 And InVBMode Then Stop
		If lDone = 0 Then
			sNibble(0) = "0000"
			sNibble(1) = "0001"
			sNibble(2) = "0010"
			sNibble(3) = "0011"
			sNibble(4) = "0100"
			sNibble(5) = "0101"
			sNibble(6) = "0110"
			sNibble(7) = "0111"
			sNibble(8) = "1000"
			sNibble(9) = "1001"
			sNibble(10) = "1010"
			sNibble(11) = "1011"
			sNibble(12) = "1100"
			sNibble(13) = "1101"
			sNibble(14) = "1110"
			sNibble(15) = "1111"
			For lDone = 0 To 255
				sByte_Renamed(lDone) = sNibble(lDone \ &H10) & sNibble(lDone And &HF)
			Next 
		End If
		
		If X < &H100 Then
			Dec2Bin = VB.Right(sByte_Renamed(X), Fixed)
		ElseIf X < &H10000 Then 
			Dec2Bin = VB.Right(sByte_Renamed(X \ &H100 And &HFF) & sByte_Renamed(X And &HFF), Fixed)
		ElseIf X < &H1000000 Then 
			Dec2Bin = VB.Right(sByte_Renamed(X \ &H10000 And &HFF) & sByte_Renamed(X \ &H100 And &HFF) & sByte_Renamed(X And &HFF), Fixed)
		Else
			Dec2Bin = VB.Right(sByte_Renamed(X \ &H1000000 And &HFF) & sByte_Renamed(X \ &H10000 And &HFF) & sByte_Renamed(X \ &H100 And &HFF) & sByte_Renamed(X And &HFF), Fixed)
		End If
		Y = Len(Dec2Bin)
		If Y < Fixed Then Dec2Bin = New String("0", Fixed - Y) & Dec2Bin
		
		'    If InVBMode Then
		'        If X <> Bin2Dec(Dec2Bin) Then Err.Raise 6
		'    End If
		
	End Function
End Class