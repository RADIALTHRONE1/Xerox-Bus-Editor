Imports System.IO
Imports System.Data.OleDb
Imports System.Text
Imports System.Windows.Forms.VisualStyles

Public Class Form1

    Public SourceDBFPath As String


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub




    'This is the "Load File" button. This will handle the loading and setting of the SOURCE.DBF file
    Private Sub btnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()

        Dim OpenFileDialog1 As New OpenFileDialog()

        OpenFileDialog1.InitialDirectory = "C:\Xerox\Bus\bus2.xrxgsn.com\"
        OpenFileDialog1.Filter = "SOURCE.DBF|source.dbf"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            SourceDBFPath = New FileInfo(OpenFileDialog1.FileName).FullName
            GlobalVariables.FILE_NAME = New FileInfo(OpenFileDialog1.FileName).Name
            GlobalVariables.FILE_PATH = New FileInfo(OpenFileDialog1.FileName).Directory.ToString & "\"

            LoadSourceDBF1()
        End If

    End Sub

    'This will read the values of the SOURCE.DBF file, convert, shove data into Raw Data, then Decoded Names
    Private Sub LoadSourceDBF1()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)

        Dim path As String = GlobalVariables.FILE_PATH & GlobalVariables.FILE_NAME
        Dim s As String
        Try
            s = System.IO.File.ReadAllText(path, encodingANSI)
        Catch
            MsgBox("You need to close IE and PWSLock first!")
            Exit Sub
        End Try

        If ckbBAKCreation.Checked = True Then
            'Make a backup of original file
            GlobalVariables.FileNameBak = GlobalVariables.FILE_NAME & "bak"
            Dim append As Boolean = False

            Dim objWriter As System.IO.StreamWriter
            objWriter = New System.IO.StreamWriter(GlobalVariables.FileNameBak, append, encodingANSI)
            objWriter.Write(s)
            objWriter.Close()
        End If

        Dim B As Byte() = encodingANSI.GetBytes(s)
        Dim sb As New StringBuilder()

        'Convert the ANSI text to Hex
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        'Get number of Records
        Dim records As Integer = (sb.Length - 836) / 1258

        'Add the header, then each split record
        ListBox1.Items.Add(sb.ToString().Substring(0, 834))
        Dim Y As Integer = 836
        For X As Integer = 0 To (records) - 1
            ListBox1.Items.Add(sb.ToString().Substring(Y, 1258))
            Y += 1258
        Next X

        'Populate the Name List
        For X As Integer = 1 To ListBox1.Items.Count - 1
            Y = 0
            Dim recordNameCorrected As String = ListBox1.Items(X).ToString().Substring(24, 256)
            Dim recordNameCorrectedASCII As String = ""

            For Z As Integer = 0 To (recordNameCorrected.ToString.Length() / 2) - 1
                Dim HexBuffer As String = recordNameCorrected.Substring(Y, 2)
                If HexBuffer = "00" Then
                    HexBuffer = ""
                Else
                    HexBuffer = (Chr(CInt("&H" & HexBuffer)))
                End If
                recordNameCorrectedASCII &= HexBuffer
                Y += 2
            Next Z
            ListBox3.Items.Add(recordNameCorrectedASCII)
        Next X
        Me.Focus()
    End Sub

    'Old code, no longer used?
    Private Sub LoadSourceDBF2()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(20127)

        Dim path As String = GlobalVariables.FILE_PATH & GlobalVariables.FILE_NAME
        Dim s As String = System.IO.File.ReadAllText(path, encodingANSI)

        If ckbBAKCreation.Checked = True Then
            'Make a backup of original file
            GlobalVariables.FileNameBak = GlobalVariables.FILE_NAME & "bak"
            Dim append As Boolean = False

            Dim objWriter As System.IO.StreamWriter
            objWriter = New System.IO.StreamWriter(GlobalVariables.FileNameBak, append, encodingANSI)
            objWriter.Write(s)
            objWriter.Close()
        End If

        'Dim B As Byte() = encodingANSI.GetBytes(s)
        Dim sb As New StringBuilder()

        'Convert the ANSI text to Hex
        'For Each singlebyte As Byte In B
        '    sb.AppendFormat("{0:x2}", singlebyte)
        'Next

        'Get number of Records
        Dim records As Integer = (sb.Length - 417) / 629

        'Add the header, then each split record

        'ListBox1.Items.Add(s)
        txtDebug.Text = s

        'ListBox1.Items.Add(sb.ToString().Substring(0, 416))
        'Dim Y As Integer = 418
        'For X As Integer = 0 To (records) - 1
        '    ListBox1.Items.Add(sb.ToString().Substring(Y, 629))
        '    Y += 629
        'Next X

        'Populate the Name List
        'For X As Integer = 1 To ListBox1.Items.Count - 1
        '    Y = 0
        '    Dim recordNameCorrected As String = ListBox1.Items(X).ToString().Substring(24, 256)
        '    Dim recordNameCorrectedASCII As String = ""

        '    For Z As Integer = 0 To (recordNameCorrected.ToString.Length() / 2) - 1
        '        Dim HexBuffer As String = recordNameCorrected.Substring(Y, 2)
        '        If HexBuffer = "00" Then
        '            HexBuffer = ""
        '        Else
        '            HexBuffer = (Chr(CInt("&H" & HexBuffer)))
        '        End If
        '        recordNameCorrectedASCII &= HexBuffer
        '        Y += 2
        '    Next Z
        '    ListBox3.Items.Add(recordNameCorrectedASCII)
        'Next X
    End Sub

    Private Sub LoadCATSRCDBF()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)

        Dim path As String = GlobalVariables.FILE_PATH & GlobalVariables.FILE_NAME
        Dim s As String = System.IO.File.ReadAllText(path, encodingANSI)

        If ckbBAKCreation.Checked = True Then
            'Make a backup of original file
            GlobalVariables.FileNameBak = GlobalVariables.FILE_NAME & "bak"
            Dim append As Boolean = False

            Dim objWriter As System.IO.StreamWriter
            objWriter = New System.IO.StreamWriter(GlobalVariables.FileNameBak, append, encodingANSI)
            objWriter.Write(s)
            objWriter.Close()
        End If

        Dim B As Byte() = encodingANSI.GetBytes(s)
        Dim sb As New StringBuilder()

        'Convert the ANSI text to Hex
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        'Get number of Records
        Dim records As Integer = (sb.Length - 194) / 50

        'Add the header, then each split record
        ListBox1.Items.Add(sb.ToString().Substring(0, 192))
        Dim Y As Integer = 194
        For X As Integer = 0 To (records) - 1
            ListBox1.Items.Add(sb.ToString().Substring(Y, 50))
            Y += 50
        Next X

        Dim source_path As String = GlobalVariables.FILE_PATH & "SOURCE.DBF"
        Dim ss As String = System.IO.File.ReadAllText(source_path, encodingANSI)

        Dim BB As Byte() = encodingANSI.GetBytes(ss)
        Dim sb2 As New StringBuilder()

        'Convert the ANSI text to Hex
        For Each singlebyte As Byte In BB
            sb2.AppendFormat("{0:x2}", singlebyte)
        Next

        'Get number of Records
        Dim records_source As Integer = (sb2.Length - 836) / 1258

        'Add the header, then each split record
        ListBox2.Items.Add(sb2.ToString().Substring(0, 834))
        Y = 836
        For X As Integer = 0 To (records_source) - 1
            ListBox2.Items.Add(sb2.ToString().Substring(Y, 1258))
            Y += 1258
        Next X

        'For X As Integer = 1 To ListBox2.Items.Count - 1
        '    Y = 0
        '    Dim recordNameCorrected As String = ListBox2.Items(X).ToString().Substring(24, 256)
        '    Dim recordNameCorrectedASCII As String = ""

        '    For Z As Integer = 0 To (recordNameCorrected.ToString.Length() / 2) - 1
        '        Dim HexBuffer As String = recordNameCorrected.Substring(Y, 2)
        '        If HexBuffer = "00" Then
        '            HexBuffer = ""
        '        Else
        '            HexBuffer = (Chr(CInt("&H" & HexBuffer)))
        '        End If
        '        recordNameCorrectedASCII &= HexBuffer
        '        Y += 2
        '    Next Z
        '    ListBox3.Items.Add(recordNameCorrectedASCII)
        'Next X

        For X As Integer = 1 To ListBox1.Items.Count - 1
            Y = 0
            Dim recordNameCorrected As String = ListBox1.Items(X).ToString().Substring(26, 24)
            Dim recordNameCorrectedASCII As String = ""

            For Z As Integer = 0 To (recordNameCorrected.ToString.Length() / 2) - 1
                Dim HexBuffer As String = recordNameCorrected.Substring(Y, 2)
                If HexBuffer = "00" Then
                    HexBuffer = ""
                Else
                    HexBuffer = (Chr(CInt("&H" & HexBuffer)))
                End If
                recordNameCorrectedASCII &= HexBuffer
                Y += 2
            Next Z
            ListBox3.Items.Add(recordNameCorrectedASCII)
        Next X
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Dim FILE_PATH As String = "P:\Original Files\Xerox\Bus\bus2.xrxgsn.com\SOURCEWORKING.dbf"
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)

        Try
            Using writer As New StreamWriter(SourceDBFPath, False, encodingANSI)
                Dim HeaderToWriteANSI As String = ListBox1.Items(0) & "20"
                Dim HeaderToWriteHEX As String
                For Y As Integer = 0 To HeaderToWriteANSI.Length - 1 Step 2
                    HeaderToWriteHEX = (Chr(CInt("&H" & HeaderToWriteANSI.Substring(Y, 2))))
                    writer.Write(HeaderToWriteHEX)
                Next


                For X As Integer = 1 To ListBox1.Items.Count - 1
                    Dim DataToWriteANSI As String = ListBox1.Items(X)
                    Dim DataToWriteHEX As String

                    For Y As Integer = 0 To DataToWriteANSI.Length - 1 Step 2
                        DataToWriteHEX = (Chr(CInt("&H" & DataToWriteANSI.Substring(Y, 2))))
                        writer.Write(DataToWriteHEX)
                    Next
                Next X
                writer.Close()
            End Using
            MessageBox.Show("Data exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"Error exporting: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnLoadRecord_Click(sender As Object, e As EventArgs) Handles btnLoadRecord.Click
        GlobalVariables.FirstLoad = True
        If GlobalVariables.FILE_NAME = "SOURCE.DBF" Then
            LoadRecordSourceDBF()
        ElseIf GlobalVariables.FILE_NAME = "CATSRC.DBF" Then
            LoadRecordCATSRCDBF()
        Else
            MsgBox("Error")
        End If
    End Sub

    Private Sub LoadRecordSourceDBF()

        Dim recordIndex As Integer = GlobalVariables.recordSelected
        txtCurrentRecord.Text = recordIndex


        Dim recordSourceID As String = ListBox1.Items(recordIndex).ToString().Substring(0, 24)
        Dim recordSourceIDASCII As String = ""

        Dim Y As Integer = 0
        For X As Integer = 0 To (recordSourceID.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordSourceID.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordSourceIDASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordSourceID.Text = recordSourceID
        txtRecordSourceIDASCII.Text = recordSourceIDASCII

        txtRecordSourceIDLength.Text = txtRecordSourceID.Text.Length
        txtRecordSourceIDASCIILength.Text = txtRecordSourceIDASCII.Text.Length

        '==========

        Dim recordName As String = ListBox1.Items(recordIndex).ToString().Substring(24, 256)
        Dim recordNameASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordName.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordName.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordNameASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordName.Text = recordName
        txtRecordNameASCII.Text = recordNameASCII

        txtRecordNameLength.Text = txtRecordName.Text.Length
        txtRecordNameASCIILength.Text = txtRecordNameASCII.Text.Length

        '==========

        Dim recordDescription As String = ListBox1.Items(recordIndex).ToString().Substring(280, 510)
        Dim recordDescriptionASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordDescription.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordDescription.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordDescriptionASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordDescription.Text = recordDescription
        txtRecordDescriptionASCII.Text = recordDescriptionASCII

        txtRecordDescriptionLength.Text = txtRecordDescription.Text.Length
        txtRecordDescriptionASCIILength.Text = txtRecordDescriptionASCII.Text.Length

        '==========

        Dim recordLastSync As String = ListBox1.Items(recordIndex).ToString().Substring(790, 38)
        Dim recordLastSyncASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordLastSync.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordLastSync.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordLastSyncASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordLastSync.Text = recordLastSync
        txtRecordLastSyncASCII.Text = recordLastSyncASCII

        txtRecordLastSyncLength.Text = txtRecordLastSync.Text.Length
        txtRecordLastSyncASCIILength.Text = txtRecordLastSyncASCII.Text.Length

        '==========

        Dim recordModified As String = ListBox1.Items(recordIndex).ToString().Substring(828, 38)
        Dim recordModifiedASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordModified.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordModified.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordModifiedASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordModified.Text = recordModified
        txtRecordModifiedASCII.Text = recordModifiedASCII

        txtRecordModifiedLength.Text = txtRecordModified.Text.Length
        txtRecordModifiedASCIILength.Text = txtRecordModifiedASCII.Text.Length

        '==========

        Dim recordDirectory As String = ListBox1.Items(recordIndex).ToString().Substring(866, 256)
        Dim recordDirectoryASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordDirectory.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordDirectory.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordDirectoryASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordDirectory.Text = recordDirectory
        txtRecordDirectoryASCII.Text = recordDirectoryASCII

        txtRecordDirectoryLength.Text = txtRecordDirectory.Text.Length
        txtRecordDirectoryASCIILength.Text = txtRecordDirectoryASCII.Text.Length

        '==========

        Dim recordWorkflow As String = ListBox1.Items(recordIndex).ToString().Substring(1122, 6)
        Dim recordWorkflowASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordWorkflow.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordWorkflow.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordWorkflowASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordWorkflow.Text = recordWorkflow
        txtRecordWorkflowASCII.Text = recordWorkflowASCII

        txtRecordWorkflowLength.Text = txtRecordWorkflow.Text.Length
        txtRecordWorkflowASCIILength.Text = txtRecordWorkflowASCII.Text.Length

        '==========

        Dim recordSubscription As String = ListBox1.Items(recordIndex).ToString().Substring(1128, 2)
        Dim recordSubscriptionASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordSubscription.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordSubscription.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordSubscriptionASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordSubscription.Text = recordSubscription
        txtRecordSubscriptionASCII.Text = recordSubscriptionASCII

        If recordSubscriptionASCII = 0 Then
            rdbSubScribeFalse.Checked = True
        Else
            rdbSubscribeTrue.Checked = True
        End If

        txtRecordSubscriptionLength.Text = txtRecordSubscription.Text.Length
        txtRecordSubscriptionASCIILength.Text = txtRecordSubscriptionASCII.Text.Length

        '==========

        Dim recordIcon As String = ListBox1.Items(recordIndex).ToString().Substring(1130, 50)
        Dim recordIconASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordIcon.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordIcon.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordIconASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordIcon.Text = recordIcon
        txtRecordIconASCII.Text = recordIconASCII

        txtRecordIconLength.Text = txtRecordIcon.Text.Length
        txtRecordIconASCIILength.Text = txtRecordIconASCII.Text.Length

        '==========

        Dim recordInvisible As String = ListBox1.Items(recordIndex).ToString().Substring(1180, 2)
        Dim recordInvisibleASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordInvisible.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordInvisible.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordInvisibleASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordInvisible.Text = recordInvisible
        txtRecordInvisibleASCII.Text = recordInvisibleASCII

        txtRecordInvisibleLength.Text = txtRecordInvisible.Text.Length
        txtRecordInvisibleASCIILength.Text = txtRecordInvisibleASCII.Text.Length

        '==========

        Dim recordSubCatID As String = ListBox1.Items(recordIndex).ToString().Substring(1182, 24)
        Dim recordSubCatIDASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordSubCatID.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordSubCatID.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordSubCatIDASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordSubCatID.Text = recordSubCatID
        txtRecordSubCatIDASCII.Text = recordSubCatIDASCII

        txtRecordSubCatIDLength.Text = txtRecordSubCatID.Text.Length
        txtRecordSubCatIDASCIILength.Text = txtRecordSubCatIDASCII.Text.Length

        '==========

        Dim recordMessage As String = ListBox1.Items(recordIndex).ToString().Substring(1206, 50)
        Dim recordMessageASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordMessage.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordMessage.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordMessageASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordMessage.Text = recordMessage
        txtRecordMessageASCII.Text = recordMessageASCII

        txtRecordMessageLength.Text = txtRecordMessage.Text.Length
        txtRecordMessageASCIILength.Text = txtRecordMessageASCII.Text.Length

        '==========

        Dim RecordLengthAll As Integer = (txtRecordSourceID.Text.Length + txtRecordName.Text.Length + txtRecordDescription.Text.Length + txtRecordLastSync.Text.Length + txtRecordModified.Text.Length + txtRecordDirectory.Text.Length + txtRecordWorkflow.Text.Length + txtRecordSubscription.Text.Length + txtRecordIcon.Text.Length + txtRecordInvisible.Text.Length + txtRecordSubCatID.Text.Length + txtRecordMessage.Text.Length)
        txtRecordLengthAll.Text = RecordLengthAll

        btnUpdateRecord.Enabled = False
        btnUndoUpdate.Enabled = False
        btnSubmitRecord.Enabled = False

        GlobalVariables.BackupRecordSourceID = txtRecordSourceIDASCII.Text
        GlobalVariables.BackupRecordName = txtRecordNameASCII.Text
        GlobalVariables.BackupRecordDescription = txtRecordDescriptionASCII.Text
        GlobalVariables.BackupRecordDirectory = txtRecordDirectoryASCII.Text
        GlobalVariables.BackupRecordWorkflow = txtRecordWorkflowASCII.Text
        GlobalVariables.BackupRecordSubscription = txtRecordSubscriptionASCII.Text
        GlobalVariables.BackupRecordIcon = txtRecordIconASCII.Text
        GlobalVariables.BackupRecordInvisible = txtRecordInvisibleASCII.Text
        GlobalVariables.BackupRecordSubcatID = txtRecordSubCatIDASCII.Text
        GlobalVariables.BackupRecordMessage = txtRecordMessageASCII.Text
    End Sub

    Private Sub LoadRecordCATSRCDBF()
        GlobalVariables.FirstLoad = True
        Dim recordIndex As Integer = GlobalVariables.recordSelected
        txtCurrentRecord.Text = recordIndex


        Dim recordSourceID As String = ListBox1.Items(recordIndex).ToString().Substring(0, 26)
        Dim recordSourceIDASCII As String = ""

        Dim Y As Integer = 0
        For X As Integer = 0 To (recordSourceID.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordSourceID.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordSourceIDASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordSourceID.Text = recordSourceID
        txtRecordSourceIDASCII.Text = recordSourceIDASCII

        txtRecordSourceIDLength.Text = txtRecordSourceID.Text.Length
        txtRecordSourceIDASCIILength.Text = txtRecordSourceIDASCII.Text.Length



        Dim recordSubCatID As String = ListBox1.Items(recordIndex).ToString().Substring(26, 24)
        Dim recordSubCatIDASCII As String = ""

        Y = 0
        For X As Integer = 0 To (recordSubCatID.Length() / 2) - 1
            Dim ASCIIBuffer As String = recordSubCatID.Substring(Y, 2)
            If ASCIIBuffer = "00" Then
                ASCIIBuffer = "2A"
            End If
            recordSubCatIDASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
            Y += 2
        Next X

        txtRecordSubCatID.Text = recordSubCatID
        txtRecordSubCatIDASCII.Text = recordSubCatIDASCII

        txtRecordSubCatIDLength.Text = txtRecordSubCatID.Text.Length
        txtRecordSubCatIDASCIILength.Text = txtRecordSubCatIDASCII.Text.Length
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        ListBox1.SelectedIndex = ListBox3.SelectedIndex() + 1
        GlobalVariables.recordSelected = ListBox1.SelectedIndex()
        btnLoadRecord.Enabled = True
    End Sub

    Private Sub txtRecordSourceIDASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordSourceIDASCII.TextChanged
        txtRecordSourceIDASCIILength.Text = txtRecordSourceIDASCII.Text.Length

        If txtRecordSourceIDASCII.Text.Length = 12 Then
            txtRecordSourceIDASCIILength.BackColor = Color.FromName("Control")

            If txtRecordSourceIDASCII.Text <> GlobalVariables.BackupRecordSourceID And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordSourceIDASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordNameASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordNameASCII.TextChanged
        txtRecordNameASCIILength.Text = txtRecordNameASCII.Text.Length

        If txtRecordNameASCII.Text.Length = 128 Then
            txtRecordNameASCIILength.BackColor = Color.FromName("Control")

            If txtRecordNameASCII.Text <> GlobalVariables.BackupRecordName And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordNameASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordDescriptionASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordDescriptionASCII.TextChanged
        txtRecordDescriptionASCIILength.Text = txtRecordDescriptionASCII.Text.Length

        If txtRecordDescriptionASCII.Text.Length = 255 Then
            txtRecordDescriptionASCIILength.BackColor = Color.FromName("Control")

            If txtRecordDescriptionASCII.Text <> GlobalVariables.BackupRecordDescription And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordDescriptionASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordDirectoryASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordDirectoryASCII.TextChanged
        txtRecordDirectoryASCIILength.Text = txtRecordDirectoryASCII.Text.Length

        If txtRecordDirectoryASCII.Text.Length = 128 Then
            txtRecordDirectoryASCIILength.BackColor = Color.FromName("Control")

            If txtRecordDirectoryASCII.Text <> GlobalVariables.BackupRecordDirectory And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordDirectoryASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If

    End Sub

    Private Sub txtRecordWorkflowASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordWorkflowASCII.TextChanged
        txtRecordWorkflowASCIILength.Text = txtRecordWorkflowASCII.Text.Length

        If txtRecordWorkflowASCII.Text.Length = 3 Then
            txtRecordWorkflowASCIILength.BackColor = Color.FromName("Control")

            If txtRecordWorkflowASCII.Text <> GlobalVariables.BackupRecordWorkflow And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordWorkflowASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordSubscriptionASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordSubscriptionASCII.TextChanged
        txtRecordSubscriptionASCIILength.Text = txtRecordSubscriptionASCII.Text.Length

        If txtRecordSubscriptionASCII.Text.Length = 1 Then
            txtRecordSubscriptionASCIILength.BackColor = Color.FromName("Control")

            If txtRecordSubscriptionASCII.Text <> GlobalVariables.BackupRecordSubscription And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordSubscriptionASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordIconASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordIconASCII.TextChanged
        txtRecordIconASCIILength.Text = txtRecordIconASCII.Text.Length

        If txtRecordIconASCII.Text.Length = 25 Then
            txtRecordIconASCIILength.BackColor = Color.FromName("Control")

            If txtRecordIconASCII.Text <> GlobalVariables.BackupRecordIcon And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordIconASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordInvisibleASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordInvisibleASCII.TextChanged
        txtRecordInvisibleASCIILength.Text = txtRecordInvisibleASCII.Text.Length

        If txtRecordInvisibleASCII.Text.Length = 1 Then
            txtRecordInvisibleASCIILength.BackColor = Color.FromName("Control")

            If txtRecordInvisibleASCII.Text <> GlobalVariables.BackupRecordInvisible And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordInvisibleASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordSubCatIDASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordSubCatIDASCII.TextChanged
        txtRecordSubCatIDASCIILength.Text = txtRecordSubCatIDASCII.Text.Length

        If txtRecordSubCatIDASCII.Text.Length = 12 Then
            txtRecordSubCatIDASCIILength.BackColor = Color.FromName("Control")

            If txtRecordSubCatIDASCII.Text <> GlobalVariables.BackupRecordSubcatID And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordSubCatIDASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
    End Sub

    Private Sub txtRecordMessageASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordMessageASCII.TextChanged
        txtRecordMessageASCIILength.Text = txtRecordMessageASCII.Text.Length

        If txtRecordMessageASCII.Text.Length = 25 Then
            txtRecordMessageASCIILength.BackColor = Color.FromName("Control")

            If txtRecordMessageASCII.Text <> GlobalVariables.BackupRecordMessage And GlobalVariables.FirstLoad = False Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If

        Else
            txtRecordMessageASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If

        GlobalVariables.FirstLoad = False
    End Sub

    Private Sub btnUpdateRecord_Click(sender As Object, e As EventArgs) Handles btnUpdateRecord.Click
        UpdateSourceID()
        UpdateName()
        UpdateDescription()
        UpdateDirectory()
        UpdateWorkflow()
        UpdateSubscription()
        UpdateIcon()
        UpdateInvisible()
        UpdateSubcatID()
        UpdateMessage()

        btnUndoUpdate.Enabled = True
        btnUpdateRecord.Enabled = False
        btnSubmitRecord.Enabled = True
    End Sub

    Private Sub UpdateSourceID()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordSourceIDASCII As String = txtRecordSourceIDASCII.Text
        Dim newRecordSourceIDHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordSourceIDASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordSourceIDHex = sb.ToString

        For X As Integer = 0 To (newRecordSourceIDHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordSourceIDHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordSourceIDHex = HexBuffer
        txtRecordSourceID.Text = newRecordSourceIDHex
    End Sub

    Private Sub UpdateName()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordNameASCII As String = txtRecordNameASCII.Text
        Dim newRecordNameHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordNameASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordNameHex = sb.ToString

        For X As Integer = 0 To (newRecordNameHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordNameHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordNameHex = HexBuffer
        txtRecordName.Text = newRecordNameHex
    End Sub

    Private Sub UpdateDescription()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordDescriptionASCII As String = txtRecordDescriptionASCII.Text
        Dim newRecordDescriptionHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordDescriptionASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordDescriptionHex = sb.ToString

        For X As Integer = 0 To (newRecordDescriptionHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordDescriptionHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordDescriptionHex = HexBuffer
        txtRecordDescription.Text = newRecordDescriptionHex
    End Sub

    Private Sub UpdateDirectory()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordDirectoryASCII As String = txtRecordDirectoryASCII.Text
        Dim newRecordDirectoryHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordDirectoryASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordDirectoryHex = sb.ToString

        For X As Integer = 0 To (newRecordDirectoryHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordDirectoryHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordDirectoryHex = HexBuffer
        txtRecordDirectory.Text = newRecordDirectoryHex
    End Sub

    Private Sub UpdateWorkflow()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordWorkflowASCII As String = txtRecordWorkflowASCII.Text
        Dim newRecordWorkflowHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordWorkflowASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordWorkflowHex = sb.ToString

        For X As Integer = 0 To (newRecordWorkflowHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordWorkflowHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordWorkflowHex = HexBuffer
        txtRecordWorkflow.Text = newRecordWorkflowHex
    End Sub

    Private Sub UpdateSubscription()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordSubscriptionASCII As String = txtRecordSubscriptionASCII.Text
        Dim newRecordSubscriptionHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordSubscriptionASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordSubscriptionHex = sb.ToString

        For X As Integer = 0 To (newRecordSubscriptionHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordSubscriptionHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordSubscriptionHex = HexBuffer
        txtRecordSubscription.Text = newRecordSubscriptionHex
    End Sub

    Private Sub UpdateIcon()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordIconASCII As String = txtRecordIconASCII.Text
        Dim newRecordIconHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordIconASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordIconHex = sb.ToString

        For X As Integer = 0 To (newRecordIconHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordIconHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordIconHex = HexBuffer
        txtRecordIcon.Text = newRecordIconHex
    End Sub

    Private Sub UpdateInvisible()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordInvisibleASCII As String = txtRecordInvisibleASCII.Text
        Dim newRecordInvisibleHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordInvisibleASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordInvisibleHex = sb.ToString

        For X As Integer = 0 To (newRecordInvisibleHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordInvisibleHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordInvisibleHex = HexBuffer
        txtRecordInvisible.Text = newRecordInvisibleHex
    End Sub

    Private Sub UpdateSubcatID()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordSubcatIDASCII As String = txtRecordSubCatIDASCII.Text
        Dim newRecordSubcatIDHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordSubcatIDASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordSubcatIDHex = sb.ToString

        For X As Integer = 0 To (newRecordSubcatIDHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordSubcatIDHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordSubcatIDHex = HexBuffer
        txtRecordSubCatID.Text = newRecordSubcatIDHex
    End Sub

    Private Sub UpdateMessage()
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim newRecordMessageASCII As String = txtRecordMessageASCII.Text
        Dim newRecordMessageHex As String = ""
        Dim HexBuffer As String = ""

        Dim Y As Integer = 0
        Dim sb As New StringBuilder()


        Dim B As Byte() = encodingANSI.GetBytes(newRecordMessageASCII)
        For Each singlebyte As Byte In B
            sb.AppendFormat("{0:x2}", singlebyte)
        Next

        newRecordMessageHex = sb.ToString

        For X As Integer = 0 To (newRecordMessageHex.Length() / 2) - 1
            Dim ASCIIBuffer As String = newRecordMessageHex.Substring(Y, 2)
            If ASCIIBuffer = "2a" Then
                ASCIIBuffer = "00"
            End If
            HexBuffer &= ASCIIBuffer
            Y += 2
        Next X

        newRecordMessageHex = HexBuffer
        txtRecordMessage.Text = newRecordMessageHex

        btnUpdateRecord.Enabled = False
    End Sub

    Private Sub btnUndoUpdate_Click(sender As Object, e As EventArgs) Handles btnUndoUpdate.Click
        txtRecordSourceIDASCII.Text = GlobalVariables.BackupRecordSourceID
        txtRecordNameASCII.Text = GlobalVariables.BackupRecordName
        txtRecordDescriptionASCII.Text = GlobalVariables.BackupRecordDescription
        txtRecordDirectoryASCII.Text = GlobalVariables.BackupRecordDirectory
        txtRecordWorkflowASCII.Text = GlobalVariables.BackupRecordWorkflow
        txtRecordSubscriptionASCII.Text = GlobalVariables.BackupRecordSubscription
        txtRecordIconASCII.Text = GlobalVariables.BackupRecordIcon
        txtRecordInvisibleASCII.Text = GlobalVariables.BackupRecordInvisible
        txtRecordSubCatIDASCII.Text = GlobalVariables.BackupRecordSubcatID
        txtRecordMessageASCII.Text = GlobalVariables.BackupRecordMessage

        UpdateSourceID()
        UpdateName()
        UpdateDescription()
        UpdateDirectory()
        UpdateWorkflow()
        UpdateSubscription()
        UpdateIcon()
        UpdateInvisible()
        UpdateSubcatID()
        UpdateMessage()

        btnUndoUpdate.Enabled = False
        btnSubmitRecord.Enabled = True
        btnUpdateRecord.Enabled = False

    End Sub

    Private Sub btnSubmitRecord_Click(sender As Object, e As EventArgs) Handles btnSubmitRecord.Click
        Dim CompleteRecord As String = ""
        Dim Index As Integer = txtCurrentRecord.Text

        CompleteRecord &= txtRecordSourceID.Text
        CompleteRecord &= txtRecordName.Text
        CompleteRecord &= txtRecordDescription.Text
        CompleteRecord &= txtRecordLastSync.Text
        CompleteRecord &= txtRecordModified.Text
        CompleteRecord &= txtRecordDirectory.Text
        CompleteRecord &= txtRecordWorkflow.Text
        CompleteRecord &= txtRecordSubscription.Text
        CompleteRecord &= txtRecordIcon.Text
        CompleteRecord &= txtRecordInvisible.Text
        CompleteRecord &= txtRecordSubCatID.Text
        CompleteRecord &= txtRecordMessage.Text
        CompleteRecord &= "20"

        ListBox1.Items.RemoveAt(Index)
        ListBox1.Items.Insert(Index, CompleteRecord)

        btnSave.Enabled = True
        btnSubmitRecord.Enabled = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
        Dim recordsAllAllAll As String = ""
        Dim append As Boolean = False

        Dim HexToString As String = ListBox1.Items(0) & "20"

        For X As Integer = 1 To ListBox1.Items.Count - 1
            HexToString &= ListBox1.Items(X)
        Next X

        Dim Y As Integer = 0

        For Z As Integer = 0 To (HexToString.ToString.Length() / 2) - 1
            Dim HexBuffer As String = HexToString.Substring(Y, 2)

            HexBuffer = (Chr(CInt("&H" & HexBuffer)))

            recordsAllAllAll &= HexBuffer
            Y += 2
        Next Z
    End Sub

    Private Sub btnDebug_Click(sender As Object, e As EventArgs) Handles btnDebug.Click
        btnLoadRecord.Enabled = True
        btnUpdateRecord.Enabled = True
        btnUndoUpdate.Enabled = True
        btnSubmitRecord.Enabled = True
        btnSave.Enabled = True
    End Sub
    Private Sub rdbSubscribeTrue_CheckedChanged(sender As Object, e As EventArgs) Handles rdbSubscribeTrue.CheckedChanged
        txtRecordSubscriptionASCII.Text = 1
    End Sub
    Private Sub rdbSubScribeFalse_CheckedChanged(sender As Object, e As EventArgs) Handles rdbSubScribeFalse.CheckedChanged
        txtRecordSubscriptionASCII.Text = 0
    End Sub

    Private Sub btnFormUsage_Click(sender As Object, e As EventArgs) Handles btnFormUsage.Click
        MsgBox("This is for fixing the PWS Searchlite EDOC source list when you load new equipment into a PWS (through the EDOC Manager or Individual Installs)

ALWAYS MAKE A BACKUP COPY OF FILES BEFORE MODIFYING THEM!")
    End Sub

    Private Sub btnExSubCatID_Click(sender As Object, e As EventArgs) Handles btnExSubCatID.Click
        MsgBox("The folder the record will be stored in on the source list.
You can find out the ID of a folder by right-clicking it and selecting 'properties'.
12 Characters, front space filled.")
    End Sub

    Private Sub btnExWorkflow_Click(sender As Object, e As EventArgs) Handles btnExWorkflow.Click
        MsgBox("Not sure exactly what this does, but these are the typical values:
*** = Blank
CS* = Customer Solutions
EKA = Eureka Tech Tip
BUL = Eureka Bulletin
ED* = EDOC
SPL = Supplies?
SWF = Software?")
    End Sub
    Private Sub btnExIcon_Click(sender As Object, e As EventArgs) Handles btnExIcon.Click
        MsgBox("This assigns the little icon for the record:
*** = Blank
CS* = Customer Solutions
EKA* = Eureka Tech Tip
BUL* = Eureka Bulletin
EDO* = EDOC
EDOC* = EDOC
SW* = Supplies? Software?

Must have terminating Star, then space fill")
    End Sub






























    'Private Sub X()
    '    Dim encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)
    '    Dim newRecordASCII As String = txtRecordASCII.Text
    '    Dim newRecordHex As String = ""
    '    Dim HexBuffer As String = ""

    '    Dim Y As Integer = 0
    '    Dim sb As New StringBuilder()


    '    Dim B As Byte() = encodingANSI.GetBytes(newRecordASCII)
    '    For Each singlebyte As Byte In B
    '        sb.AppendFormat("{0:x2}", singlebyte)
    '    Next

    '    newRecordHex = sb.ToString

    '    For X As Integer = 0 To (newRecordHex.Length() / 2) - 1
    '        Dim ASCIIBuffer As String = newRecordHex.Substring(Y, 2)
    '        If ASCIIBuffer = "2a" Then
    '            ASCIIBuffer = "00"
    '        End If
    '        HexBuffer &= ASCIIBuffer
    '        Y += 2
    '    Next X

    '    newRecordHex = HexBuffer
    '    txtRecord.Text = newRecordHex
    'End Sub





End Class



Public Class GlobalVariables
    Public Shared dt As New DataTable
    Public Shared HexTemp As String
    Public Shared recordSelected As Integer
    Public Shared FileName As String
    Public Shared FileNameBak As String
    Public Shared FILE_NAME As String
    Public Shared FILE_PATH As String

    Public Shared BackupRecordSourceID As String
    Public Shared BackupRecordName As String
    Public Shared BackupRecordDescription As String
    Public Shared BackupRecordDirectory As String
    Public Shared BackupRecordWorkflow As String
    Public Shared BackupRecordSubscription As String
    Public Shared BackupRecordIcon As String
    Public Shared BackupRecordInvisible As String
    Public Shared BackupRecordSubcatID As String
    Public Shared BackupRecordMessage As String


    Public Shared FirstLoad As Boolean = True

End Class


