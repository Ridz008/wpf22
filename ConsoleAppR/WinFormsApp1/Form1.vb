
Option Explicit On

Private hszDDEServer As Long ' String handle for the server name.
Private hszDDETopic As Long ' String handle for the topic name.
Private hszDDEItemPoke As Long  ' String handle for the Poke item name.
Private bRunning As Boolean ' Server running flag.
Public Class Form1


    Private Sub Command2_Click()
        ' Clear the text box.
        Text1.Text = ""

    End Sub

    Private Sub Command1_Click()
        ' We have to initiate a DDEPostAdvise() in order to let all interested clients
        ' know that something has changed.
        If (DdePostAdvise(lInstID, 0, 0)) Then
            Debug.Print "DdePostAdvise() Success."
      Else
            Debug.Print "DdePostAdvise() Failed."
      End If
    End Sub

Private Sub Form_Load()
    Command1.Caption = "Advice"
    Command2.Caption = "Clear"
    Me.Caption = "DDE Server"

    ' Initialize the DDE subsystem. We need to let the DDEML know what callback
    ' function we intend to use so we pass it address using the AddressOf operator.
    ' If we can't initialize the DDEML subsystem we exit.
    If DdeInitialize(lInstID, AddressOf DDECallback, APPCMD_FILTERINITS, 0) Then
        Exit Sub
    End If

    ' Now that the DDEML subsystem is initialized we create string handles for our
    ' server/topic name.
    hszDDEServer = DdeCreateStringHandle(lInstID, DDE_SERVER, CP_WINANSI)
    hszDDETopic = DdeCreateStringHandle(lInstID, DDE_TOPIC, CP_WINANSI)
    hszDDEItemPoke = DdeCreateStringHandle(lInstID, DDE_POKE, CP_WINANSI)
    hszDDEItemAdvise = DdeCreateStringHandle(lInstID, DDE_ADVISE, CP_WINANSI)

    ' Lets check to see if another DDE server has already registered with identical
    ' server/topic names. If so we'll exit. If we were to continue the DDE subsystem
    ' could become unstable when a client tried to converse with the server/topic.
    If (DdeConnect(lInstID, hszDDEServer, hszDDETopic, ByVal 0&)) Then
        MsgBox "A DDE server named " & Chr(34) & DDE_SERVER & Chr(34) & " with topic " &
        Chr(34) & DDE_TOPIC & Chr(34) & " is already running.", vbExclamation, App.Title
    Unload Me
    Exit Sub
    End If

    ' We need to register the server with the DDE subsystem.
    If (DdeNameService(lInstID, hszDDEServer, 0, DNS_REGISTER)) Then
        ' Set the server running flag.
        bRunning = True
    End If

    Me.WindowState = vbMinimized
End Sub

    Private Sub Form_Terminate()
        ' We need to release our string handles.
        DdeFreeStringHandle lInstID, hszDDEServer
  DdeFreeStringHandle lInstID, hszDDETopic
  DdeFreeStringHandle lInstID, hszDDEItemPoke
  DdeFreeStringHandle lInstID, hszDDEItemAdvise

  ' Unregister the DDE server.
        If bRunning Then
            DdeNameService lInstID, hszDDEServer, 0, DNS_UNREGISTER
  End If

        ' Break down the link with the DDE subsystem.
        DdeUninitialize lInstID
End Sub
End Class