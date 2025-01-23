Imports System.IO.Ports
Imports System.Management
Imports System.Runtime.CompilerServices


Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        TimerDateAndTime.Interval = 1000
        TimerDateAndTime.Start()
        Timer_Splash.Start()
    End Sub

    Private Sub TimerDateAndTime_Tick(sender As Object, e As EventArgs) Handles TimerDateAndTime.Tick
        LabelDateVal.Text = DateTime.Now.ToShortDateString()
        LabelTimeVal.Text = DateTime.Now.ToLongTimeString()
    End Sub

    Private Sub ButtonAutoConnect_Click(sender As Object, e As EventArgs) Handles ButtonAutoConnect.Click

        ' Connect to the identified Arduino port
        Try
            SerialPort1.BaudRate = 9600
            SerialPort1.PortName = ComboBoxPort.SelectedItem
            SerialPort1.Open()

            LabelStatus.Text = "CONNECTED"
            ButtonAutoConnect.SendToBack()
            ButtonDisconnect.BringToFront()

            'Enable controls to proceed
            ButtonDisconnect.Enabled = True


            '------------------------------'

            PictureBoxConnected.BringToFront()
            PictureBoxDisconnected.SendToBack()
            LabelStatus.Text = "CONNECTED"
            LabelStatus.ForeColor = Color.DarkBlue
            LabelStatus.Left = (LabelStatus.Parent.ClientSize.Width - LabelStatus.Width) \ 2
            ButtonAutoConnect.Enabled = False
            ButtonDisconnect.Enabled = True
        Catch ex As Exception
            MsgBox("Please check the Hardware, COM, Baud Rate and try again.", MsgBoxStyle.Critical, "Connection failed !!!")
        End Try

    End Sub

    Private Sub ButtonDisconnect_Click(sender As Object, e As EventArgs) Handles ButtonDisconnect.Click
        If SerialPort1.IsOpen Then
            TimerDateAndTime.Stop()
            SerialPort1.Close()
        End If
        PictureBoxDisconnected.BringToFront()
        PictureBoxConnected.SendToBack()
        LabelStatus.Text = "DISCONNECTED"
        LabelStatus.ForeColor = Color.Red
        LabelStatus.Left = (LabelStatus.Parent.ClientSize.Width - LabelStatus.Width) \ 2
        ButtonAutoConnect.Enabled = True
        ButtonDisconnect.Enabled = False
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' If result = DialogResult.No Then
        ' e.Cancel = True ' Prevent the form from closing
        '   End If
    End Sub

    Private Sub SerialPort1_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Try

            Dim data As String = SerialPort1.ReadLine() ' Read the received data

            ' Filter out any unwanted characters
            data = data.Trim() ' Remove leading and trailing whitespace
            data = data.Replace(vbCrLf, "") ' Remove newline characters

            ' Invoke a method to update the labels on the UI thread
            Me.Invoke(Sub() UpdateLabels(data))
        Catch ex As Exception
            MessageBox.Show("Error receiving data: " & ex.Message)
        End Try
    End Sub

    Private Sub UpdateLabels(data As String)
        If data.StartsWith("voltPos") Then
            Dim voltPosStr As String = data.Substring(7).Trim()
            LabelVoltPos.Text = voltPosStr

            Dim voltPosDouble As Double = Double.Parse(voltPosStr)
            If voltPosDouble <= 2.977 Then  ' if voltPosDouble is less than 2.977
                PictureBoxVoltPos.BackColor = Color.Red

            ElseIf voltPosDouble >= 2.999 Then
                PictureBoxVoltPos.BackColor = Color.Green  ' if voltPosDouble is more than 2.977
            End If

        ElseIf data.StartsWith("voltNeg") Then
            Dim voltNegStr As String = data.Substring(7).Trim()
            LabelVoltNeg.Text = voltNegStr

            ' Convert String to Double
            Dim voltNegDouble As Double = Double.Parse(voltNegStr)
            If voltNegDouble <= 2.973 Then
                PictureBoxVoltNeg.BackColor = Color.Red

            ElseIf voltNegDouble >= 2.992 Then
                PictureBoxVoltNeg.BackColor = Color.Green
            End If

        ElseIf data.StartsWith("ampPos:") Then
            Dim ampPosStr As String = data.Substring(7).Trim()
            LabelAmpPos.Text = ampPosStr

            ' Convert String to Double
            Dim ampPosDouble As Double = Double.Parse(ampPosStr)
            If ampPosDouble <= 3.307 Then
                PictureBoxAmpPos.BackColor = Color.Red

            ElseIf ampPosDouble >= 1.016 Then
                PictureBoxAmpPos.BackColor = Color.Green
            End If


        ElseIf data.StartsWith("ampNeg:") Then
            Dim ampNegStr As String = data.Substring(7).Trim()
            LabelNegAmp.Text = ampNegStr

            ' Convert String to Double
            Dim ampNegDouble As Double = Double.Parse(ampNegStr)
            If ampNegDouble <= 3.092 Then
                PictureBoxAmpNeg.BackColor = Color.Red

            ElseIf ampNegDouble >= 1.283 Then
                PictureBoxAmpNeg.BackColor = Color.Green
            End If

        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to Restart?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Application.Restart() ' Close the application
        End If

    End Sub

    Private Sub Timer_Splash_Tick(sender As Object, e As EventArgs) Handles Timer_Splash.Tick
        Timer_Splash.Interval = 3000
        PictureBoxPulse.Visible = False

        Timer_Splash.Stop()
    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ButtonScanPort_Click(sender As Object, e As EventArgs) Handles ButtonScanPort.Click
        If LabelStatus.Text = "STATUS: CONNECTED" Then
            MsgBox("Conncetion in progress, please Disconnect to scan the new port.", MsgBoxStyle.Critical, "Warning !!!")
            Return
        End If
        ComboBoxPort.Items.Clear()
        Dim myPort As Array
        Dim i As Integer
        myPort = System.IO.Ports.SerialPort.GetPortNames()
        ComboBoxPort.Items.AddRange(myPort)
        i = ComboBoxPort.Items.Count
        i = i - i
        Try
            ComboBoxPort.SelectedIndex = i
            ButtonAutoConnect.Enabled = True
        Catch ex As Exception
            MsgBox("Com port not detected", MsgBoxStyle.Critical, "Warning !!!")
            ComboBoxPort.Text = ""
            ComboBoxPort.Items.Clear()
            Return
        End Try
        ComboBoxPort.DroppedDown = True

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If SerialPort1.IsOpen Then
            SerialPort1.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub PictureBoxAmpNeg_Click(sender As Object, e As EventArgs) Handles PictureBoxAmpNeg.Click

    End Sub


    '  End Sub

End Class
