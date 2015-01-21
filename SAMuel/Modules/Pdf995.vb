Imports System.IO

Namespace Modules
    Module Pdf995

        Private ModifiedIni As String = Directory.GetCurrentDirectory() + "\pdf995.ini"

        ''' <summary>
        ''' Creates the pdf995ini file configured to work as desired.
        ''' </summary>
        ''' <param name="outputFolder">Where the created PDFs will be saved to.</param>
        ''' <remarks></remarks>
        Sub CreatePdf995Ini(ByVal outputFolder As String)
            'Removes the ending /
            outputFolder = outputFolder.Remove(outputFolder.Length - 1)

            'Creates or Overwrites the XML file
            Dim fs1 As FileStream = New FileStream(Pdf995Ini, FileMode.Create, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            'Create ini file
            s1.Write("[Parameters]" & vbCrLf)
            s1.Write("Display Readme=0" & vbCrLf)
            s1.Write("AutoLaunch=0" & vbCrLf) 'Disable PDF from opening after print.
            s1.Write("Quiet=1" & vbCrLf) 'Disables UI
            s1.WriteLine("Use GPL Ghostscript=1") '?
            's1.Write("Document Name=" & documentName & vbCrLf) 'Source file
            s1.Write("Initial Dir=" & outputFolder & vbCrLf) 'Output folder?
            s1.Write("Output File=SAMEASDOCUMENT" & vbCrLf) 'Output file name
            s1.WriteLine("Output Folder=" & outputFolder) 'Output Folder

            s1.WriteLine("[OmniFormat]")
            s1.WriteLine("Accept EULA=1")

            s1.Close()
            fs1.Close()
        End Sub



        ''' <summary>
        ''' Restores the copy of the original PDF995ini file and deletes the backup.
        ''' </summary>
        ''' <remarks></remarks>
        Sub RestorePdf995Ini()
            If (File.Exists(ModifiedIni)) Then
                File.Copy(ModifiedIni, Pdf995Ini, True)
                File.Delete(ModifiedIni)
            End If
        End Sub

        ''' <summary>
        ''' Deletes the existing PDF995ini file.
        ''' </summary>
        ''' <remarks>PDF995 creates a default configuration file if there is no existing.</remarks>
        Sub DefaultPdf995Ini()
            If (File.Exists(Pdf995Ini)) Then
                File.Delete(Pdf995Ini)
            End If
        End Sub

        ''' <summary>
        ''' Backup the copy of the original PDF995ini file.
        ''' </summary>
        ''' <remarks></remarks>
        Sub BackupOriginalPdf995Ini()
            ModifiedIni = SAVE_FOLDER + "original_pdf995.ini"
            If (File.Exists(Pdf995Ini)) Then
                File.Copy(Pdf995Ini, ModifiedIni, True)
            End If
        End Sub

        Public Sub SetupPdf995Ini()
            'Delete the ini then copy over the modified one.
            If (File.Exists(ModifiedIni)) Then
                If (File.Exists(Pdf995Ini)) Then
                    File.Delete(Pdf995Ini)
                End If
                File.Copy(ModifiedIni, My.Settings.Pdf995ini_File)
            Else
                MsgBox("Modified PDF995.ini is missing. Please contact Zachary Karpinski.",, "Missing PDF995.ini!")
            End If
        End Sub
    End Module
End Namespace