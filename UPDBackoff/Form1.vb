Imports System.IO
Imports ComponentAce.Compression.Archiver
Imports ComponentAce.Compression.ZipForge
Imports Ionic.Zip
Imports MySql.Data.MySqlClient
Public Class Form1
    Dim conlocal As MySqlConnection
    Dim zip As New ZipForge()
    Private totalFiles As Integer
    Dim isadaupdpos As Boolean
    Dim isadaupdiki As Boolean
    Dim cabang As String
    Dim sizepos, sizeiki As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox1.Visible = False
        cabang = GetIniSetting(My.Application.Info.DirectoryPath + "\setting.ini", "INFO", "Cabang")
        If cabang = "NOT_FOUND" Then
            SetIniSettings(My.Application.Info.DirectoryPath + "\setting.ini", "INFO", "Cabang", "0000")
        End If
        If cabang = "0000" Then
            MsgBox("setting.ini harus disetting dahulu")
            Me.Close()
        End If
        System.Windows.Forms.Form.CheckForIllegalCrossThreadCalls = False
        BackgroundWorker1.WorkerSupportsCancellation = True
        isadaupdiki = False
        isadaupdpos = False
        If Not BackgroundWorker1.IsBusy Then
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Public Sub autoupd()
        Try
            lblload.Text = "Mulai aplikasi"
            Dim tmp As String = Application.StartupPath & "\tmp\"
            If Not Directory.Exists(tmp) Then
                Directory.CreateDirectory(tmp)
            End If

            Dim tmpi As String = Application.StartupPath & "\tmp2\"
            If Not Directory.Exists(tmpi) Then
                Directory.CreateDirectory(tmpi)
            End If

            Dim tmp2 As String = Application.StartupPath & "\OpenZip\"
            If Not Directory.Exists(tmp2) Then
                Directory.CreateDirectory(tmp2)
            End If
            Dim pbackoff As String = Application.StartupPath & "\BACKOFF\"
            If Not Directory.Exists(pbackoff) Then
                Directory.CreateDirectory(pbackoff)
            End If

            Dim pikiosk As String = Application.StartupPath & "\IKIOSK\"
            If Not Directory.Exists(pikiosk) Then
                Directory.CreateDirectory(pikiosk)
            End If
            Dim filezippos As String = "Backoff" & Format(Now, "yyMMdd") & ".zip"
            Dim filezipiki As String = "Ikiosk" & Format(Now, "yyMMdd") & ".zip"


            Dim filePaths As String() = Directory.GetFiles(tmp, "*.zip")
            Dim nupath As Integer = filePaths.Count

            ProgressBar1.Maximum = nupath
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            If nupath <> 0 Then

                For Each data As String In filePaths
                    Dim fi As FileInfo = New FileInfo(data)
                    Dim temp As String = fi.Name.Length

                    If UnzipFileF(tmp & fi.Name, tmp2) = "done" Then
                        ProgressBar1.Value += 1
                        lblload.Text = fi.Name
                        Dim filePaths2 As String() = Directory.GetFiles(tmp2, "*.*")
                        Dim nupath2 As Integer = filePaths2.Count
                        ProgressBar2.Maximum = nupath2
                        ProgressBar2.Minimum = 0
                        ProgressBar2.Value = 0
                        If nupath2 <> 0 Then
                            For Each data2 As String In filePaths2
                                Dim fi2 As FileInfo = New FileInfo(data2)
                                Dim temp2 As String = fi2.Name.Length

                                ProgressBar2.Value += 1
                                If File.Exists(pbackoff & fi2.Name) Then
                                    If DelFileF(pbackoff & fi2.Name) = "done" Then
                                        File.Copy(tmp2 & fi2.Name, pbackoff & fi2.Name, True)
                                    Else
                                        File.Copy(tmp2 & fi2.Name, pbackoff & fi2.Name, True)
                                    End If

                                Else
                                    File.Copy(tmp2 & fi2.Name, pbackoff & fi2.Name, True)
                                End If
                                File.Delete(tmp2 & fi2.Name)

                            Next
                        End If
                    End If
                    File.Delete(tmp & fi.Name)
                Next
                cekprogram()
                isadaupdpos = True
            End If


            '-------------------IKIOSK
            Dim filePathi As String() = Directory.GetFiles(tmpi, "*.zip")
            Dim nupathi As Integer = filePathi.Count

            ProgressBar1.Maximum = nupathi
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            If nupathi <> 0 Then

                For Each datai As String In filePathi
                    Dim fii As FileInfo = New FileInfo(datai)
                    Dim tempi As String = fii.Name.Length

                    If UnzipFileF(tmpi & fii.Name, tmp2) = "done" Then
                        ProgressBar1.Value += 1
                        lblload.Text = fii.Name
                        Dim filePaths2i As String() = Directory.GetFiles(tmp2, "*.*")
                        Dim nupath2i As Integer = filePaths2i.Count
                        ProgressBar2.Maximum = nupath2i
                        ProgressBar2.Minimum = 0
                        ProgressBar2.Value = 0
                        If nupath2i <> 0 Then
                            For Each data2i As String In filePaths2i
                                Dim fi2i As FileInfo = New FileInfo(data2i)
                                Dim temp2i As String = fi2i.Name.Length

                                ProgressBar2.Value += 1
                                If File.Exists(pikiosk & fi2i.Name) Then
                                    If DelFileF(pikiosk & fi2i.Name) = "done" Then
                                        File.Copy(tmp2 & fi2i.Name, pikiosk & fi2i.Name, True)
                                    Else
                                        File.Copy(tmp2 & fi2i.Name, pikiosk & fi2i.Name, True)
                                    End If

                                Else
                                    File.Copy(tmp2 & fi2i.Name, pikiosk & fi2i.Name, True)
                                End If
                                File.Delete(tmp2 & fi2i.Name)

                            Next
                        End If
                    End If
                    File.Delete(tmpi & fii.Name)
                Next
                cekprogram2()
                isadaupdiki = True
            End If


            If isadaupdpos Then
                PictureBox1.Visible = True
                lblload.Text = "Zip Pos"
                Try
                    Dim tempZipPath As String = Application.StartupPath & "\temp_backoff.zip"
                    If File.Exists(tempZipPath) Then
                        File.Delete(tempZipPath)
                    End If
                    Dim sevenZipPath As String = Application.StartupPath & "\7z.exe"
                    Dim arguments As String = "a -tzip """ & tempZipPath & """ """ & pbackoff & "*"""
                    Dim processInfo As New ProcessStartInfo()
                    processInfo.FileName = sevenZipPath
                    processInfo.Arguments = arguments
                    processInfo.UseShellExecute = False
                    processInfo.RedirectStandardOutput = True
                    processInfo.RedirectStandardError = True
                    processInfo.CreateNoWindow = True
                    Dim process As Process = Process.Start(processInfo)
                    process.WaitForExit()

                    If process.ExitCode <> 0 Then
                        Dim errorOutput As String = process.StandardError.ReadToEnd()
                        MsgBox("Error during zipping: " & errorOutput)
                        Return
                    End If

                    If Directory.Exists("E:\PRG_TAMPUNG\Backoff\") Then
                        'If File.Exists("E:\PRG_TAMPUNG\Backoff\" & filezippos) Then
                        '    File.Delete("E:\PRG_TAMPUNG\Backoff\" & filezippos)
                        'End If
                        Dim folderPath As String = "E:\PRG_TAMPUNG\Backoff\"
                        Dim pattern As String = "Backoff??????.zip"
                        For Each filePath As String In Directory.GetFiles(folderPath, pattern)
                            File.Delete(filePath)
                        Next
                        File.Move(tempZipPath, "E:\PRG_TAMPUNG\Backoff\" & filezippos)
                        Dim fileInfo As New FileInfo("E:\PRG_TAMPUNG\Backoff\" & filezippos)
                        If fileInfo.Exists Then
                            Dim fileSizeInBytes As Long = fileInfo.Length
                            Dim fileSizeInKB As Double = fileSizeInBytes / 1024
                            Dim fileSizeInMB As Double = fileSizeInKB / 1024
                            sizepos = fileSizeInMB.ToString("F2") & " MB"
                            getsize(filezippos, sizepos, "backoff")
                        End If
                    Else
                        'If File.Exists("D:\PRG_TAMPUNG\Backoff\" & filezippos) Then
                        '    File.Delete("D:\PRG_TAMPUNG\Backoff\" & filezippos)
                        'End If
                        Dim folderPath As String = "D:\PRG_TAMPUNG\Backoff\"
                        Dim pattern As String = "Backoff??????.zip"
                        For Each filePath As String In Directory.GetFiles(folderPath, pattern)
                            File.Delete(filePath)
                        Next
                        File.Move(tempZipPath, "D:\PRG_TAMPUNG\Backoff\" & filezippos)
                        Dim fileInfo As New FileInfo("D:\PRG_TAMPUNG\Backoff\" & filezippos)
                        If fileInfo.Exists Then
                            Dim fileSizeInBytes As Long = fileInfo.Length
                            Dim fileSizeInKB As Double = fileSizeInBytes / 1024
                            Dim fileSizeInMB As Double = fileSizeInKB / 1024
                            sizepos = fileSizeInMB.ToString("F2") & " MB"
                            getsize(filezippos, sizepos, "backoff")
                        End If
                    End If

                    File.Delete(tempZipPath)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return
                End Try

                PictureBox1.Visible = False
            End If
            If isadaupdiki Then
                PictureBox1.Visible = True
                lblload.Text = "Zip Ikiosk"
                Try
                    Dim tempZipPath As String = Application.StartupPath & "\temp_ikiosk.zip"
                    If File.Exists(tempZipPath) Then
                        File.Delete(tempZipPath)
                    End If
                    Dim sevenZipPath As String = Application.StartupPath & "\7z.exe"
                    Dim arguments As String = "a -tzip """ & tempZipPath & """ """ & pikiosk & "*"""
                    Dim processInfo As New ProcessStartInfo()
                    processInfo.FileName = sevenZipPath
                    processInfo.Arguments = arguments
                    processInfo.UseShellExecute = False
                    processInfo.RedirectStandardOutput = True
                    processInfo.RedirectStandardError = True
                    processInfo.CreateNoWindow = False
                    Dim process As Process = Process.Start(processInfo)
                    process.WaitForExit()

                    If process.ExitCode <> 0 Then
                        Dim errorOutput As String = process.StandardError.ReadToEnd()
                        MsgBox("Error during zipping: " & errorOutput)
                        Return
                    End If

                    If Directory.Exists("E:\PRG_TAMPUNG\i-kiosk\") Then
                        'If File.Exists("E:\PRG_TAMPUNG\i-kiosk\" & filezipiki) Then
                        '    File.Delete("E:\PRG_TAMPUNG\i-kiosk\" & filezipiki)
                        'End If
                        Dim folderPath As String = "E:\PRG_TAMPUNG\i-kiosk\"
                        Dim pattern As String = "ikiosk??????.zip"
                        For Each filePath As String In Directory.GetFiles(folderPath, pattern)
                            File.Delete(filePath)
                        Next
                        File.Move(tempZipPath, "E:\PRG_TAMPUNG\i-kiosk\" & filezipiki)
                        Dim fileInfo As New FileInfo("E:\PRG_TAMPUNG\i-kiosk\" & filezipiki)
                        If fileInfo.Exists Then
                            Dim fileSizeInBytes As Long = fileInfo.Length
                            Dim fileSizeInKB As Double = fileSizeInBytes / 1024
                            Dim fileSizeInMB As Double = fileSizeInKB / 1024
                            sizeiki = fileSizeInMB.ToString("F2") & " MB"
                            getsize(filezipiki, sizeiki, "ikiosk")
                        End If
                    Else
                        'If File.Exists("D:\PRG_TAMPUNG\i-kiosk\" & filezipiki) Then
                        '    File.Delete("D:\PRG_TAMPUNG\i-kiosk\" & filezipiki)
                        'End If
                        Dim folderPath As String = "D:\PRG_TAMPUNG\i-kiosk\"
                        Dim pattern As String = "ikiosk??????.zip"
                        For Each filePath As String In Directory.GetFiles(folderPath, pattern)
                            File.Delete(filePath)
                        Next
                        File.Move(tempZipPath, "D:\PRG_TAMPUNG\i-kiosk\" & filezipiki)
                        Dim fileInfo As New FileInfo("D:\PRG_TAMPUNG\i-kiosk\" & filezipiki)
                        If fileInfo.Exists Then
                            Dim fileSizeInBytes As Long = fileInfo.Length
                            Dim fileSizeInKB As Double = fileSizeInBytes / 1024
                            Dim fileSizeInMB As Double = fileSizeInKB / 1024
                            sizeiki = fileSizeInMB.ToString("F2") & " MB"
                            getsize(filezipiki, sizeiki, "ikiosk")
                        End If
                    End If

                    File.Delete(tempZipPath)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return
                End Try
                PictureBox1.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message & ex.StackTrace)
            Me.Close()
        End Try
        Me.Close()
    End Sub
    Public Function UnzipFileF(fileZip As String, folder As String) As String
        Dim ket As String
        Try
            Using zip As ZipFile = ZipFile.Read(fileZip)
                zip.ExtractExistingFile = ExtractExistingFileAction.OverwriteSilently
                zip.ExtractAll(folder)
            End Using
            ket = "done"
        Catch ex As Exception
            ket = "error"
        End Try
        Return ket
    End Function
    Public Sub konekLocal()
        Dim sql As String = "server=192.168.190.100;uid=root;pwd=15032012;database=coba;Pooling=false"
        conlocal = New MySqlConnection(sql)
        Try
            conLocal.Open()
        Catch ex As Exception
            MsgBox("Tidak Terhubung ke database")
        End Try
    End Sub

    Private Sub getsize(sfile As String, ssize As String, jenis As String)
        Try
            lblload.Text = "Kirim Size"
            konekLocal()
            Dim sqlb As String = "delete from db_scrap.master_ukuran_cabang where cabang='" + cabang + "' and namafile like '%" & jenis & "%'"
            Dim cmdx As MySqlCommand = New MySqlCommand(sqlb, conlocal)
            cmdx.ExecuteNonQuery()

            Dim sqlc As String = "REPLACE INTO db_scrap.MASTER_UKURAN_CABANG (cabang,namafile,ukuran,addtime) values(?, ?, ?, ?)"
            Using cmdC As New MySqlCommand(sqlc, conlocal)
                cmdC.Parameters.AddWithValue("@cabang", cabang)
                cmdC.Parameters.AddWithValue("@namafile", sfile)
                cmdC.Parameters.AddWithValue("@ukuran", ssize)
                cmdC.Parameters.AddWithValue("@addtime", DateTime.Now)
                cmdC.ExecuteNonQuery()
            End Using
            lblload.Text = "Selesi size"
        Catch ex As Exception
            MsgBox(ex.Message)
            lblload.Text = "Error : " & ex.Message
        End Try
    End Sub
    Private Sub cekprogram()
        Try
            lblload.Text = "Kirim versi"
            konekLocal()
            Dim xsql As String = "SELECT program, exist_version FROM program WHERE jenis != 'IKIOSK' ORDER BY program ASC"
            Dim ds As New DataTable
            Dim da As New MySqlDataAdapter(xsql, conlocal)

            da.Fill(ds)

            Dim a As Integer = 0
            Dim b As Integer = ds.Rows.Count
            Dim vtoko, vserver, Status, vtgl As String
            Dim totnok As Integer = 0
            Dim totok As Integer = 0
            Dim totwew As Integer = 0
            Dim totsim As Integer = 0
            Dim sql1 As String = "delete from db_scrap.master_backoff_cabang where cabang='" & cabang & "'"
            Dim cmd1 As MySqlCommand = New MySqlCommand(sql1, conlocal)
            cmd1.ExecuteNonQuery()

            While a < b
                vserver = ds.Rows(a).Item(1).ToString()
                Status = ""
                vtgl = ""

                Try
                    vtoko = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\Backoff\" & ds.Rows(a).Item(0).ToString()).FileVersion
                    vtgl = Format(File.GetCreationTime(Application.StartupPath & "\Backoff\" & ds.Rows(a).Item(0).ToString()), "yyyy-MM-dd HH:mm:ss ")
                Catch ex As Exception
                    vtoko = "0"
                    totwew += 1
                    vtgl = "0000-00-00 00:00:00"
                End Try

                Dim xvtoko As Double = Convert.ToDouble(vtoko.Replace(".", ""))
                Dim xvserver As Double = Convert.ToDouble(vserver.Replace(".", ""))

                If xvtoko < xvserver Then
                    If xvtoko <> 0 Then
                        Status = "NOK"
                        totnok += 1
                    End If
                ElseIf xvtoko > xvserver Then
                    Status = "Melebihi Master"
                    totsim += 1
                Else
                    totok += 1
                    Status = "OK"
                End If

                Dim ck As Boolean = (Status <> "OK")


                Dim sql As String = "REPLACE INTO db_scrap.master_backoff_cabang (cabang, namaprogram, versi, addtime) VALUES (?, ?, ?, ?)"
                Using cmd As New MySqlCommand(sql, conlocal)
                    cmd.Parameters.AddWithValue("@cabang", cabang)
                    cmd.Parameters.AddWithValue("@namaprogram", ds.Rows(a).Item(0).ToString())
                    cmd.Parameters.AddWithValue("@versi", vtoko)
                    cmd.Parameters.AddWithValue("@addtime", DateTime.Now)
                    cmd.ExecuteNonQuery()
                End Using

                a += 1
            End While
            lblload.Text = "Selesi versi"
        Catch ex As Exception
            MsgBox("Gagal cek versi program: " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub
    Private Sub cekprogram2()
        Try
            lblload.Text = "Kirim versi"
            konekLocal()
            Dim xsql As String = "SELECT program, exist_version FROM program WHERE jenis like '%ikiosk%' ORDER BY program ASC"
            Dim ds As New DataTable
            Dim da As New MySqlDataAdapter(xsql, conlocal)

            da.Fill(ds)

            Dim a As Integer = 0
            Dim b As Integer = ds.Rows.Count
            Dim vtoko, vserver, Status, vtgl As String
            Dim totnok As Integer = 0
            Dim totok As Integer = 0
            Dim totwew As Integer = 0
            Dim totsim As Integer = 0
            Dim sql1 As String = "delete from db_scrap.master_ikiosk_cabang where cabang='" & cabang & "'"
            Dim cmd1 As MySqlCommand = New MySqlCommand(sql1, conlocal)
            cmd1.ExecuteNonQuery()

            While a < b
                vserver = ds.Rows(a).Item(1).ToString()
                Status = ""
                vtgl = ""

                Try
                    vtoko = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\ikiosk\" & ds.Rows(a).Item(0).ToString()).FileVersion
                    vtgl = Format(File.GetCreationTime(Application.StartupPath & "\ikiosk\" & ds.Rows(a).Item(0).ToString()), "yyyy-MM-dd HH:mm:ss ")
                Catch ex As Exception
                    vtoko = "0"
                    totwew += 1
                    vtgl = "0000-00-00 00:00:00"
                End Try

                Dim xvtoko As Double = Convert.ToDouble(vtoko.Replace(".", ""))
                Dim xvserver As Double = Convert.ToDouble(vserver.Replace(".", ""))

                If xvtoko < xvserver Then
                    If xvtoko <> 0 Then
                        Status = "NOK"
                        totnok += 1
                    End If
                ElseIf xvtoko > xvserver Then
                    Status = "Melebihi Master"
                    totsim += 1
                Else
                    totok += 1
                    Status = "OK"
                End If

                Dim ck As Boolean = (Status <> "OK")


                Dim sql As String = "REPLACE INTO db_scrap.master_ikiosk_cabang (cabang, namaprogram, versi, addtime) VALUES (?, ?, ?, ?)"
                Using cmd As New MySqlCommand(sql, conlocal)
                    cmd.Parameters.AddWithValue("@cabang", cabang)
                    cmd.Parameters.AddWithValue("@namaprogram", ds.Rows(a).Item(0).ToString())
                    cmd.Parameters.AddWithValue("@versi", vtoko)
                    cmd.Parameters.AddWithValue("@addtime", DateTime.Now)
                    cmd.ExecuteNonQuery()
                End Using

                a += 1
            End While
            lblload.Text = "Selesi versi"
        Catch ex As Exception
            MsgBox("Gagal cek versi program: " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub
    Public Function DelFileF(lok1 As String) As String
        Dim ket As String
        Try
            File.Delete(lok1)
            ket = "done"
        Catch ex As Exception
            ket = "error"
        End Try
        Return ket
    End Function


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        autoupd()
    End Sub
End Class
