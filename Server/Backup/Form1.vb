Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents h As System.Windows.Forms.Button
    Friend WithEvents c As System.Windows.Forms.Button
    Friend WithEvents w1 As AxMSWinsockLib.AxWinsock
    Friend WithEvents s As System.Windows.Forms.Button
    Friend WithEvents txt As System.Windows.Forms.TextBox
    Friend WithEvents cs As System.Windows.Forms.Button
    Friend WithEvents rt As System.Windows.Forms.ListBox
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.h = New System.Windows.Forms.Button()
        Me.c = New System.Windows.Forms.Button()
        Me.w1 = New AxMSWinsockLib.AxWinsock()
        Me.s = New System.Windows.Forms.Button()
        Me.txt = New System.Windows.Forms.TextBox()
        Me.cs = New System.Windows.Forms.Button()
        Me.rt = New System.Windows.Forms.ListBox()
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.w1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'h
        '
        Me.h.Location = New System.Drawing.Point(224, 16)
        Me.h.Name = "h"
        Me.h.Size = New System.Drawing.Size(56, 24)
        Me.h.TabIndex = 0
        Me.h.Text = "Host"
        '
        'c
        '
        Me.c.Location = New System.Drawing.Point(224, 48)
        Me.c.Name = "c"
        Me.c.Size = New System.Drawing.Size(56, 24)
        Me.c.TabIndex = 1
        Me.c.Text = "Connect"
        '
        'w1
        '
        Me.w1.Enabled = True
        Me.w1.Location = New System.Drawing.Point(240, 104)
        Me.w1.Name = "w1"
        Me.w1.OcxState = CType(resources.GetObject("w1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.w1.Size = New System.Drawing.Size(28, 28)
        Me.w1.TabIndex = 2
        '
        's
        '
        Me.s.Location = New System.Drawing.Point(224, 160)
        Me.s.Name = "s"
        Me.s.Size = New System.Drawing.Size(56, 24)
        Me.s.TabIndex = 3
        Me.s.Text = "Send"
        '
        'txt
        '
        Me.txt.Location = New System.Drawing.Point(8, 160)
        Me.txt.Name = "txt"
        Me.txt.Size = New System.Drawing.Size(200, 20)
        Me.txt.TabIndex = 5
        Me.txt.Text = ""
        '
        'cs
        '
        Me.cs.Location = New System.Drawing.Point(224, 160)
        Me.cs.Name = "cs"
        Me.cs.Size = New System.Drawing.Size(56, 24)
        Me.cs.TabIndex = 6
        Me.cs.Text = "&Send"
        '
        'rt
        '
        Me.rt.Location = New System.Drawing.Point(8, 16)
        Me.rt.Name = "rt"
        Me.rt.Size = New System.Drawing.Size(200, 134)
        Me.rt.TabIndex = 7
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4, Me.MenuItem1, Me.MenuItem2, Me.MenuItem3, Me.MenuItem5})
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.Text = "Clear Chat Window"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.Text = "Change Nick"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 3
        Me.MenuItem3.Text = "Disconnect"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 4
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem7})
        Me.MenuItem5.Text = "Select Mode"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Host"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "Connect"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Info
        Me.Label1.Location = New System.Drawing.Point(216, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 48)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Right Click For Advance Menu"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(296, 189)
        Me.ContextMenu = Me.ContextMenu1
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1, Me.rt, Me.cs, Me.txt, Me.s, Me.w1, Me.c, Me.h})
        Me.Name = "Form1"
        Me.Text = "1 On 1 Chat By Jeremiah Ng"
        CType(Me.w1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim nick As Object
    ' This chat Program has 2 send buttons overlapping, dont blame me as im only a beginner
    Private Sub c_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles c.Click
        'For Connect(Client)
        Dim port
        Dim host
        s.Visible = False 'Host send button is not visible when connecting 
        h.Visible = False 'Host button is not visible when connecting or connected
        c.Enabled = False 'Client Connect button is not Enabled again when connecting  or connected
        'disable and enable Right Click Menu if used
        MenuItem3.Enabled = True
        MenuItem6.Enabled = False
        MenuItem7.Enabled = False
        Me.Text = "Client Mode:  I'm " & nick 'Shows the mode and the Nick only if Connect is clicked
        port = 80 'host port is set to 80
        'An inputbox asking for the Host Ip Address
        host = InputBox("Enter IP of the Host", , "LocalHost")
        w1.Connect(host, port) 'Connects to the host ip and port
    End Sub

    Private Sub h_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles h.Click
        'For Host(Server)
        MsgBox("You are now Hosting") 'Let the Host know the Hosting is Valid
        cs.Visible = False
        h.Enabled = False
        c.Visible = False
        MenuItem3.Enabled = True
        MenuItem6.Enabled = False
        MenuItem7.Enabled = False
        Me.Text = "Server Mode:  I'm " & nick 'Shows the mode and the Nick 
        w1.LocalPort = 80 'set port to 80
        w1.Listen() 'waiting for connection
    End Sub

    Private Sub w1_ConnectionRequest(ByVal sender As Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_ConnectionRequestEvent) Handles w1.ConnectionRequest
        'Used In Host(Server)
        w1.Close() 'stop connection
        w1.Accept(e.requestID) 'get the connection made
        rt.Items.Add("Client Connected") 'shows the Client has connected on the Chat Window
    End Sub

    Private Sub w1_DataArrival(ByVal sender As Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles w1.DataArrival
        'Used in Host(Server) and Connect(Client)
        Dim dat As Object
        w1.GetData(dat) 'To recive the incomming data(TEXT)
        rt.Items.Add(temstring(dat)) 'shows the data(TEXT) recived on the Chat Window
    End Sub

    Private Sub s_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles s.Click
        'For Host(Server)
        rt.Items.Add(nick + " : " & txt.Text) 'shows the Host Nick and data(Text) sent on the Chat Window
        w1.SendData(nick + " : " & txt.Text) 'To Send the Text from the host textbox to the Client's Screen
        txt.Text = "" 'clear the text in txt box
    End Sub

    Private Sub cs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cs.Click
        'For Connect(Client)
        rt.Items.Add(nick + " : " & txt.Text)
        w1.SendData(nick + " : " & txt.Text)
        txt.Text = ""
    End Sub

    Private Function temstring(ByVal strData As Array) As String
        'Used in Host(Server) and Connect(Client)
        'need to bind the data as object or somthing like that
        Dim temp As Long, mString As String
        mString = ""
        For temp = 0 To UBound(strData)
            mString &= Chr(strData(temp))
        Next
        Return mString
    End Function

    Private Sub w1_ConnectEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles w1.ConnectEvent
        'For Connect(Client)
        rt.Items.Add("Connected to Server")
    End Sub

    Private Sub w1_CloseEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles w1.CloseEvent
        'Used in Host(Server) and Connect(Client)
        rt.Items.Add("You have Been Disconnected")
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        nick = InputBox("Enter your Nick", , "SomeOne")
        MenuItem3.Enabled = False
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        'Used in Host(Server) and Connect(Client)
        nick = InputBox("Enter Your New Nick") 'set Nick = Nick
        rt.Items.Add("Nick Change to:" & nick)
        w1.SendData("Nick Change to:" & nick)
    End Sub


    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        'the disconnection button
        s.Visible = True
        cs.Visible = True
        s.Enabled = False
        cs.Enabled = False
        MenuItem3.Enabled = False
        c.Visible = True
        h.Visible = True
        c.Enabled = True
        h.Enabled = True
        txt.Text = ""
        rt.Items.Add("****Disconnnected****")
        w1.Close()
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        'Used in Host(Server) and Connect(Client)
        rt.Items.Clear()
    End Sub
    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        'For Connect
        Dim port
        Dim host
        s.Visible = False
        h.Visible = False
        c.Enabled = False
        MenuItem3.Enabled = True
        MenuItem6.Enabled = False
        MenuItem7.Enabled = False
        Me.Text = "Client Mode:  I'm " & nick 'Shows the mode and the Nick 
        port = 80
        host = InputBox("Enter IP of the Host", , "LocalHost")
        w1.Connect(host, port)
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        'For Host
        MsgBox("You are now Hosting")
        cs.Visible = False
        h.Enabled = False
        c.Visible = False
        MenuItem3.Enabled = True
        MenuItem6.Enabled = False
        MenuItem7.Enabled = False
        Me.Text = "Server Mode:  I'm " & nick 'Shows the mode and the Nick 
        w1.LocalPort = 80
        w1.Listen()
    End Sub
End Class
