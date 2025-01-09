Imports System.IO
Imports System.Data.OleDb
Imports System.Text
Imports System.Windows.Forms.VisualStyles

Public Class EditSourceDBFWindow
    Public encodingANSI As Encoding = System.Text.Encoding.GetEncoding(1252)

    Public SourceDBFPath As String
    Public SourceDBFFullPath As String

    Public SourceDBFHeader As String
    Public SourceDBFData As New List(Of Integer)

    Public UpdatedRecord As Boolean = False


    Private Sub XeroxBusEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.FirstTimeUser = True Then
            MsgBox("Warning! It is highly recommended to make a backup of your Bus files before using this program. It has a minumum amount of safety checks, and can break your SOURCE.DBF file if not careful.", MsgBoxStyle.OkOnly)
            My.Settings.FirstTimeUser = False
        End If
    End Sub

    Private Sub XeroxBusEditor_DisgracefulClosing(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.FormClosing
        End
    End Sub

    'This is the "Load File" button. This will handle the loading and setting of the SOURCE.DBF file
    Private Sub btnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click
        Dim OpenSourceDBFFile As New OpenFileDialog()

        OpenSourceDBFFile.InitialDirectory = "C:\Xerox\Bus\bus2.xrxgsn.com\"
        OpenSourceDBFFile.Filter = "SOURCE.DBF|source.dbf"

        If OpenSourceDBFFile.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            SourceDBFData.Clear()
            lbxRawData.Items.Clear()
            lbxConvertedNames.Items.Clear()
            btnLoadRecord.Enabled = False
            ckbBAKCreation.Enabled = False

            SourceDBFFullPath = New FileInfo(OpenSourceDBFFile.FileName).FullName
            SourceDBFPath = New FileInfo(OpenSourceDBFFile.FileName).DirectoryName
            LoadSourceDBF()
        End If
    End Sub

    'This will read the values of the SOURCE.DBF file, convert, shove data into Raw Data, then Decoded Names
    Private Sub LoadSourceDBF()
        Dim SourceDBF_Data As String
        Try
            SourceDBF_Data = System.IO.File.ReadAllText(SourceDBFFullPath, encodingANSI)
        Catch
            MsgBox("You need to close IE and PWSLock first!")
            Exit Sub
        End Try

        If ckbBAKCreation.Checked = True Then
            'Make a backup of original file
            Dim FileNameBackup As String = SourceDBFPath & "_bak" & Now
            My.Computer.FileSystem.CopyFile(SourceDBFFullPath, SourceDBFPath & "\SOURCE_bak.DBF", overwrite:=True)
            My.Settings.ConfirmedBackup = True
        End If


        Dim ByteFormat As Byte() = encodingANSI.GetBytes(SourceDBF_Data)
        Dim sb As New StringBuilder()

        'Convert the ANSI text to Hex
        Try
            For Each singlebyte As Byte In ByteFormat
                sb.AppendFormat("{0:x2}", singlebyte)
            Next
            If (sb.Length - 836) Mod 1258 <> 0 Then
                Dim ConfirmWeird = MsgBox("Something seems wrong with your SOURCE.DBF file, it has a weird number of characters." & vbCrLf &
                                          "You can continue if you want, but it's likely to error out.", MsgBoxStyle.OkCancel)
                If ConfirmWeird = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try

        'Get number of Records. 836 is the length of the header, and each record should be 1258 characters.
        Dim records As Integer = (sb.Length - 836) / 1258
        txtTotalRecords.Text = records

        'Add the header, then each split record
        lbxRawData.Items.Add(sb.ToString().Substring(0, 834))

        Dim Y As Integer = 836
        Try
            For X As Integer = 0 To (records) - 1
                lbxRawData.Items.Add(sb.ToString().Substring(Y, 1258))
                Y += 1258
            Next X
        Catch ex As Exception
            MsgBox("Your SOURCE.DBF file is broken as shit yo")
        End Try

        'Populate the Name List
        For X As Integer = 1 To lbxRawData.Items.Count - 1
            Y = 0
            Dim recordNameCorrected As String = lbxRawData.Items(X).ToString().Substring(24, 256)
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
            lbxConvertedNames.Items.Add(recordNameCorrectedASCII)
        Next X

        'Save the header data, then remove it from the list
        SourceDBFHeader = lbxRawData.Items(0) & "20"
        lbxRawData.Items.RemoveAt(0)

        lbxRawData.SetSelected(0, True)
        lbxConvertedNames.SetSelected(0, True)
        btnLoadRecord.Enabled = True
        Me.Activate()
    End Sub



    Private Sub btnLoadRecord_Click(sender As Object, e As EventArgs) Handles btnLoadRecord.Click
        If UpdatedRecord = True Then
            Dim Question = MsgBox("You've changed a record, but haven't saved its changes. Continue?", MsgBoxStyle.YesNo)
            If Question = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        Dim recordIndex As Integer = lbxConvertedNames.SelectedIndex
        txtCurrentRecord.Text = lbxConvertedNames.SelectedIndex

        LoadSourceDBFRecord_SourceID(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Name(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Description(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_LastSync(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Modified(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Directory(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Workflow(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Subscription(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Icon(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Invisible(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_SubCatID(lbxConvertedNames.SelectedIndex)
        LoadSourceDBFRecord_Message(lbxConvertedNames.SelectedIndex)

        txtRecordLengthAll_TextChanged()

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

    Private Sub LoadSourceDBFRecord_SourceID(RecordIndex)
        Dim recordSourceID As String = lbxRawData.Items(RecordIndex).ToString().Substring(0, 24)
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
    End Sub
    Private Sub LoadSourceDBFRecord_Name(RecordIndex)
        Dim recordName As String = lbxRawData.Items(RecordIndex).ToString().Substring(24, 256)
        Dim recordNameASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_Description(RecordIndex)
        Dim recordDescription As String = lbxRawData.Items(RecordIndex).ToString().Substring(280, 510)
        Dim recordDescriptionASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_LastSync(RecordIndex)
        Dim recordLastSync As String = lbxRawData.Items(RecordIndex).ToString().Substring(790, 38)
        Dim recordLastSyncASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_Modified(RecordIndex)
        Dim recordModified As String = lbxRawData.Items(RecordIndex).ToString().Substring(828, 38)
        Dim recordModifiedASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_Directory(RecordIndex)
        Dim recordDirectory As String = lbxRawData.Items(RecordIndex).ToString().Substring(866, 256)
        Dim recordDirectoryASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_Workflow(RecordIndex)
        Dim recordWorkflow As String = lbxRawData.Items(RecordIndex).ToString().Substring(1122, 6)
        Dim recordWorkflowASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_Subscription(RecordIndex)
        Dim recordSubscription As String = lbxRawData.Items(RecordIndex).ToString().Substring(1128, 2)
        Dim recordSubscriptionASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_Icon(RecordIndex)
        Dim recordIcon As String = lbxRawData.Items(RecordIndex).ToString().Substring(1130, 50)
        Dim recordIconASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_Invisible(RecordIndex)
        Dim recordInvisible As String = lbxRawData.Items(RecordIndex).ToString().Substring(1180, 2)
        Dim recordInvisibleASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub
    Private Sub LoadSourceDBFRecord_SubCatID(RecordIndex)
        Dim recordSubCatID As String = lbxRawData.Items(RecordIndex).ToString().Substring(1182, 24)
        Dim recordSubCatIDASCII As String = ""
        Dim Y As Integer = 0

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
    Private Sub LoadSourceDBFRecord_Message(RecordIndex)
        Dim recordMessage As String = lbxRawData.Items(RecordIndex).ToString().Substring(1206, 50)
        Dim recordMessageASCII As String = ""
        Dim Y As Integer = 0

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
    End Sub

    '========================================
    '========= Live Record Updating =========
    '========================================

    Private Sub txtRecordSourceIDASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordSourceIDASCII.TextChanged
        txtRecordSourceIDASCIILength.Text = txtRecordSourceIDASCII.Text.Length

        If txtRecordSourceIDASCII.Text.Length = 12 Then
            txtRecordSourceIDASCIILength.BackColor = Color.FromName("Control")

            If txtRecordSourceIDASCII.Text <> GlobalVariables.BackupRecordSourceID Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordSourceIDASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordNameASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordNameASCII.TextChanged
        txtRecordNameASCIILength.Text = txtRecordNameASCII.Text.Length

        If txtRecordNameASCII.Text.Length = 128 Then
            txtRecordNameASCIILength.BackColor = Color.FromName("Control")

            If txtRecordNameASCII.Text <> GlobalVariables.BackupRecordName Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordNameASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordDescriptionASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordDescriptionASCII.TextChanged
        txtRecordDescriptionASCIILength.Text = txtRecordDescriptionASCII.Text.Length

        If txtRecordDescriptionASCII.Text.Length = 255 Then
            txtRecordDescriptionASCIILength.BackColor = Color.FromName("Control")

            If txtRecordDescriptionASCII.Text <> GlobalVariables.BackupRecordDescription Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordDescriptionASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordDirectoryASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordDirectoryASCII.TextChanged
        txtRecordDirectoryASCIILength.Text = txtRecordDirectoryASCII.Text.Length

        If txtRecordDirectoryASCII.Text.Length = 128 Then
            txtRecordDirectoryASCIILength.BackColor = Color.FromName("Control")

            If txtRecordDirectoryASCII.Text <> GlobalVariables.BackupRecordDirectory Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordDirectoryASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordWorkflowASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordWorkflowASCII.TextChanged
        txtRecordWorkflowASCIILength.Text = txtRecordWorkflowASCII.Text.Length

        If txtRecordWorkflowASCII.Text.Length = 3 Then
            txtRecordWorkflowASCIILength.BackColor = Color.FromName("Control")

            If txtRecordWorkflowASCII.Text <> GlobalVariables.BackupRecordWorkflow Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordWorkflowASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordSubscriptionASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordSubscriptionASCII.TextChanged
        txtRecordSubscriptionASCIILength.Text = txtRecordSubscriptionASCII.Text.Length

        If txtRecordSubscriptionASCII.Text.Length = 1 Then
            txtRecordSubscriptionASCIILength.BackColor = Color.FromName("Control")

            If txtRecordSubscriptionASCII.Text <> GlobalVariables.BackupRecordSubscription Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordSubscriptionASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordIconASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordIconASCII.TextChanged
        txtRecordIconASCIILength.Text = txtRecordIconASCII.Text.Length

        If txtRecordIconASCII.Text.Length = 25 Then
            txtRecordIconASCIILength.BackColor = Color.FromName("Control")

            If txtRecordIconASCII.Text <> GlobalVariables.BackupRecordIcon Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordIconASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordInvisibleASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordInvisibleASCII.TextChanged
        txtRecordInvisibleASCIILength.Text = txtRecordInvisibleASCII.Text.Length

        If txtRecordInvisibleASCII.Text.Length = 1 Then
            txtRecordInvisibleASCIILength.BackColor = Color.FromName("Control")

            If txtRecordInvisibleASCII.Text <> GlobalVariables.BackupRecordInvisible Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordInvisibleASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordSubCatIDASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordSubCatIDASCII.TextChanged
        txtRecordSubCatIDASCIILength.Text = txtRecordSubCatIDASCII.Text.Length

        If txtRecordSubCatIDASCII.Text.Length = 12 Then
            txtRecordSubCatIDASCIILength.BackColor = Color.FromName("Control")

            If txtRecordSubCatIDASCII.Text <> GlobalVariables.BackupRecordSubcatID Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordSubCatIDASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordMessageASCII_TextChanged(sender As Object, e As EventArgs) Handles txtRecordMessageASCII.TextChanged
        txtRecordMessageASCIILength.Text = txtRecordMessageASCII.Text.Length

        If txtRecordMessageASCII.Text.Length = 25 Then
            txtRecordMessageASCIILength.BackColor = Color.FromName("Control")

            If txtRecordMessageASCII.Text <> GlobalVariables.BackupRecordMessage Then
                btnUpdateRecord.Enabled = True
            Else
                btnUpdateRecord.Enabled = False
            End If
        Else
            txtRecordMessageASCIILength.BackColor = Color.Red
            btnUpdateRecord.Enabled = False
        End If
        txtRecordLengthAll_TextChanged()
    End Sub
    Private Sub txtRecordLengthAll_TextChanged()
        txtRecordLengthAll.Text = (
            txtRecordSourceIDASCII.Text.Length +
            txtRecordNameASCII.Text.Length +
            txtRecordDescriptionASCII.Text.Length +
            txtRecordLastSyncASCII.Text.Length +
            txtRecordModifiedASCII.Text.Length +
            txtRecordDirectoryASCII.Text.Length +
            txtRecordWorkflowASCII.Text.Length +
            txtRecordSubscriptionASCII.Text.Length +
            txtRecordIconASCII.Text.Length +
            txtRecordInvisibleASCII.Text.Length +
            txtRecordSubCatIDASCII.Text.Length +
            txtRecordMessageASCII.Text.Length)

        If txtRecordLengthAll.Text = 628 Then
            txtRecordLengthAll.BackColor = Color.FromName("Control")
        Else
            txtRecordLengthAll.BackColor = Color.Red
        End If
    End Sub

    '========================================
    '========= Submit Record Changes ========
    '========================================

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

        lbxRawData.Items.RemoveAt(Index)
        lbxRawData.Items.Insert(Index, CompleteRecord)

        btnSave.Enabled = True
        btnSubmitRecord.Enabled = False

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If My.Settings.ConfirmedBackup = False Then
            Dim AcknowledgedDanger = MsgBox("Last chance to make sure you have a backup file. Continue?", MsgBoxStyle.YesNo)
            If AcknowledgedDanger = MsgBoxResult.Yes Then
                My.Settings.ConfirmedBackup = True
            Else
                Exit Sub
            End If
        End If
        Try
            Using writer As New StreamWriter(SourceDBFFullPath, False, encodingANSI)
                Dim HeaderToWriteHEX As String
                For Y As Integer = 0 To SourceDBFHeader.Length - 1 Step 2
                    HeaderToWriteHEX = (Chr(CInt("&H" & SourceDBFHeader.Substring(Y, 2))))
                    writer.Write(HeaderToWriteHEX)
                Next

                For X As Integer = 0 To lbxRawData.Items.Count - 1
                    Dim DataToWriteANSI As String = lbxRawData.Items(X)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnDebug_Click(sender As Object, e As EventArgs) Handles btnDebug.Click
        btnLoadRecord.Enabled = True
        btnUpdateRecord.Enabled = True
        btnUndoUpdate.Enabled = True
        btnSubmitRecord.Enabled = True
        btnSave.Enabled = True

        gbxDebug.Visible = True
    End Sub
    Private Sub rdbSubscribeTrue_CheckedChanged(sender As Object, e As EventArgs) Handles rdbSubscribeTrue.CheckedChanged
        txtRecordSubscriptionASCII.Text = 1
    End Sub
    Private Sub rdbSubScribeFalse_CheckedChanged(sender As Object, e As EventArgs) Handles rdbSubScribeFalse.CheckedChanged
        txtRecordSubscriptionASCII.Text = 0
    End Sub

    Private Sub lbxConvertedNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxConvertedNames.SelectedIndexChanged
        lbxRawData.SelectedIndex = lbxConvertedNames.SelectedIndex()
        txtHighlightedRecord.Text = lbxConvertedNames.SelectedIndex()
    End Sub
    Private Sub btnExName_Click(sender As Object, e As EventArgs) Handles btnExName.Click

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
    Private Sub btnExSubCatID_Click(sender As Object, e As EventArgs) Handles btnExSubCatID.Click
        MsgBox("The folder the record will be stored in on the source list.
You can find out the ID of a folder by right-clicking it and selecting 'properties'.
12 Characters, front space filled.")
    End Sub


    Private Sub btnFormUsage_Click(sender As Object, e As EventArgs) Handles btnFormUsage.Click
        MsgBox("This is for fixing the PWS Searchlite EDOC source list when you load new equipment into a PWS (through the EDOC Manager or Individual Installs)

ALWAYS MAKE A BACKUP COPY OF FILES BEFORE MODIFYING THEM!")
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


