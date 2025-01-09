Imports System.Text

Public Class Form2

    Private Sub CATSRCEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    'Private Sub LoadRecordCATSRCDBF()
    '    GlobalVariables.FirstLoad = True
    '    Dim recordIndex As Integer = GlobalVariables.recordSelected
    '    txtCurrentRecord.Text = recordIndex


    '    Dim recordSourceID As String = lbxRawData.Items(recordIndex).ToString().Substring(0, 26)
    '    Dim recordSourceIDASCII As String = ""

    '    Dim Y As Integer = 0
    '    For X As Integer = 0 To (recordSourceID.Length() / 2) - 1
    '        Dim ASCIIBuffer As String = recordSourceID.Substring(Y, 2)
    '        If ASCIIBuffer = "00" Then
    '            ASCIIBuffer = "2A"
    '        End If
    '        recordSourceIDASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
    '        Y += 2
    '    Next X

    '    txtRecordSourceID.Text = recordSourceID
    '    txtRecordSourceIDASCII.Text = recordSourceIDASCII

    '    txtRecordSourceIDLength.Text = txtRecordSourceID.Text.Length
    '    txtRecordSourceIDASCIILength.Text = txtRecordSourceIDASCII.Text.Length



    '    Dim recordSubCatID As String = lbxRawData.Items(recordIndex).ToString().Substring(26, 24)
    '    Dim recordSubCatIDASCII As String = ""

    '    Y = 0
    '    For X As Integer = 0 To (recordSubCatID.Length() / 2) - 1
    '        Dim ASCIIBuffer As String = recordSubCatID.Substring(Y, 2)
    '        If ASCIIBuffer = "00" Then
    '            ASCIIBuffer = "2A"
    '        End If
    '        recordSubCatIDASCII &= (Chr(CInt("&H" & ASCIIBuffer)))
    '        Y += 2
    '    Next X

    '    txtRecordSubCatID.Text = recordSubCatID
    '    txtRecordSubCatIDASCII.Text = recordSubCatIDASCII

    '    txtRecordSubCatIDLength.Text = txtRecordSubCatID.Text.Length
    '    txtRecordSubCatIDASCIILength.Text = txtRecordSubCatIDASCII.Text.Length
    'End Sub

    'Private Sub LoadCATSRCDBF()

    '    Dim s As String = System.IO.File.ReadAllText(SourceDBFFullPath, encodingANSI)

    '    If ckbBAKCreation.Checked = True Then

    '    End If

    '    Dim B As Byte() = encodingANSI.GetBytes(s)
    '    Dim sb As New StringBuilder()

    '    'Convert the ANSI text to Hex
    '    For Each singlebyte As Byte In B
    '        sb.AppendFormat("{0:x2}", singlebyte)
    '    Next

    '    'Get number of Records
    '    Dim records As Integer = (sb.Length - 194) / 50

    '    'Add the header, then each split record
    '    lbxRawData.Items.Add(sb.ToString().Substring(0, 192))
    '    Dim Y As Integer = 194
    '    For X As Integer = 0 To (records) - 1
    '        lbxRawData.Items.Add(sb.ToString().Substring(Y, 50))
    '        Y += 50
    '    Next X

    '    Dim source_path As String = SourceDBFPath & "SOURCE.DBF"
    '    Dim ss As String = System.IO.File.ReadAllText(source_path, encodingANSI)

    '    Dim BB As Byte() = encodingANSI.GetBytes(ss)
    '    Dim sb2 As New StringBuilder()

    '    'Convert the ANSI text to Hex
    '    For Each singlebyte As Byte In BB
    '        sb2.AppendFormat("{0:x2}", singlebyte)
    '    Next

    '    'Get number of Records
    '    Dim records_source As Integer = (sb2.Length - 836) / 1258

    '    'Add the header, then each split record
    '    ListBox2.Items.Add(sb2.ToString().Substring(0, 834))
    '    Y = 836
    '    For X As Integer = 0 To (records_source) - 1
    '        ListBox2.Items.Add(sb2.ToString().Substring(Y, 1258))
    '        Y += 1258
    '    Next X

    '    'For X As Integer = 1 To ListBox2.Items.Count - 1
    '    '    Y = 0
    '    '    Dim recordNameCorrected As String = ListBox2.Items(X).ToString().Substring(24, 256)
    '    '    Dim recordNameCorrectedASCII As String = ""

    '    '    For Z As Integer = 0 To (recordNameCorrected.ToString.Length() / 2) - 1
    '    '        Dim HexBuffer As String = recordNameCorrected.Substring(Y, 2)
    '    '        If HexBuffer = "00" Then
    '    '            HexBuffer = ""
    '    '        Else
    '    '            HexBuffer = (Chr(CInt("&H" & HexBuffer)))
    '    '        End If
    '    '        recordNameCorrectedASCII &= HexBuffer
    '    '        Y += 2
    '    '    Next Z
    '    '    ListBox3.Items.Add(recordNameCorrectedASCII)
    '    'Next X

    '    For X As Integer = 1 To lbxRawData.Items.Count - 1
    '        Y = 0
    '        Dim recordNameCorrected As String = lbxRawData.Items(X).ToString().Substring(26, 24)
    '        Dim recordNameCorrectedASCII As String = ""

    '        For Z As Integer = 0 To (recordNameCorrected.ToString.Length() / 2) - 1
    '            Dim HexBuffer As String = recordNameCorrected.Substring(Y, 2)
    '            If HexBuffer = "00" Then
    '                HexBuffer = ""
    '            Else
    '                HexBuffer = (Chr(CInt("&H" & HexBuffer)))
    '            End If
    '            recordNameCorrectedASCII &= HexBuffer
    '            Y += 2
    '        Next Z
    '        lbxConvertedNames.Items.Add(recordNameCorrectedASCII)
    '    Next X
    'End Sub


End Class