' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2020-12-15 17:18
' ## File update Date: 2021-8-27 19:52
' ## Filename: Error_cCrypt.vb (F:\++++ Code Share\classes\Error_cCrypt.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Option Strict On

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports DevComponents.DotNetBar
Imports MySql.Data.MySqlClient
Imports System.DateTime

Public Class Error_cCrypt
    Private Shared _saltPrefix As String
    Friend Shared Property SaltPrefix() As String
        Get
            Return _saltPrefix
        End Get
        Set(ByVal value As String)
            _saltPrefix = value
        End Set
    End Property
#Region "Zustandsvariablen"
    Private Shared EncryptedString_ As String
    Private Shared DecryptedString_ As String
#End Region


#Region "Methoden"
    ' Verschlüsseln
    Public Shared Sub Encrypt(ByVal AESKeySize As Int32, ByVal DecryptedStringRes As String)
        Dim Password As String = "wvlIelaulv2fHRMFWPICqNW5d"
        Dim Salt() As Byte
        If SaltPrefix Is Nothing Then
            SaltPrefix = ""
        End If
        Salt = System.Text.Encoding.UTF8.GetBytes(SaltPrefix & "9Tbd4JWZe6FNPibLadXzE6lJSZbC6bRZAELL4iqtBTLu5nN")
        Dim GenerierterKey As New Rfc2898DeriveBytes(Password, Salt)
        Using AES As New AesManaged
            AES.KeySize = AESKeySize
            AES.BlockSize = 128
            AES.Key = GenerierterKey.GetBytes(AES.KeySize \ 8)
            AES.IV = GenerierterKey.GetBytes(AES.BlockSize \ 8)
            Using ms As New IO.MemoryStream
                Using cs As New CryptoStream(ms, AES.CreateEncryptor(), CryptoStreamMode.Write)
                    Dim Data() As Byte
                    Data = System.Text.Encoding.UTF8.GetBytes(DecryptedStringRes)
                    cs.Write(Data, 0, Data.Length)
                    cs.FlushFinalBlock()
                End Using
                Try
                    EncryptedString_ = System.Convert.ToBase64String(ms.ToArray)
                Catch ex As Exception
                    EncryptedString_ = ""
                End Try
            End Using
            AES.Clear()
        End Using
    End Sub
    ' Entschlüsseln
    Public Shared Sub Decrypt(ByVal AESKeySize As Int32, ByVal EncryptedStringVal As String)
            Dim Password As String = "wvlIelaulv2fHRMFWPICqNW5d"
            Dim Salt() As Byte

        Try
            Using AES As New AesManaged
                If EncryptedStringVal.Trim = "" OrElse EncryptedStringVal.Trim = "Oy4ZYRTB1iU2iMhA0GjAzw==" Then
                    DecryptedString_ = ""
                    Exit Sub
                End If
                Salt = System.Text.Encoding.UTF8.GetBytes("9Tbd4JWZe6FNPibLadXzE6lJSZbC6bRZAELL4iqtBTLu5nN")
                Dim GenerierterKey As New Rfc2898DeriveBytes(Password, Salt)
                AES.KeySize = AESKeySize
                AES.BlockSize = 128
                AES.Key = GenerierterKey.GetBytes(AES.KeySize \ 8)
                AES.IV = GenerierterKey.GetBytes(AES.BlockSize \ 8)
                Using ms As New IO.MemoryStream
                    Using cs As New CryptoStream(ms, AES.CreateDecryptor(), CryptoStreamMode.Write)
                        Dim Data() As Byte
                        Data = System.Convert.FromBase64String(EncryptedStringVal)
                        cs.Write(Data, 0, Data.Length)
                        cs.FlushFinalBlock()
                    End Using
                    Try
                        DecryptedString_ = System.Text.Encoding.UTF8.GetString(ms.ToArray)
                    Catch ex As Exception
                        DecryptedString_ = ""
                    End Try
                End Using
                AES.Clear()
            End Using

        Catch ex As Exception
            Dim rb As Boolean = ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Error_Helper.IsIDE() Then Stop
        End Try

    End Sub
    Public Shared Function EncryptString(ByVal Decrypted As String, Optional ByVal KeyLength As Integer = 256) As String
       'Dim cCrypt As cCrypt = New cCrypt
        If Not String.IsNullOrWhiteSpace(Decrypted) Then
            Error_cCrypt.Encrypt(KeyLength, Decrypted)
            Return Error_cCrypt.EncryptedString
        Else
            Return ""
        End If
    End Function
    Public Shared Function DecryptString(ByVal Encrypted As String, Optional ByVal KeyLength As Integer = 256) As String
       'Dim cCrypt As cCrypt = New cCrypt
        If Not String.IsNullOrWhiteSpace(Encrypted) Then
            Error_cCrypt.Decrypt(KeyLength, Error_VarConvert.ConvertToString(Encrypted, False, ""))
            Return Error_cCrypt.DecryptedString
        Else
            Return ""
        End If
    End Function

    Public Shared Function Verschluesseln(ByVal ZuVerschluesselndes As String, ByVal Schluessel As String) As String
        Dim Ausgabe As String=""
        Try
            If Not String.IsNullOrWhiteSpace(Zuverschluesselndes) AndAlso Not String.IsNullOrWhiteSpace(Schluessel) Then
                Using rd As New RijndaelManaged
                    Using md5 As New MD5CryptoServiceProvider
                        Dim key() As Byte = md5.ComputeHash(Encoding.UTF8.GetBytes(Schluessel)) ' das hier
                        rd.Key = key
                    End Using
                    rd.GenerateIV()
                    Dim iv() As Byte = rd.IV
                    Using ms As New MemoryStream
                        ms.Write(iv, 0, iv.Length)
                        Using cs As New CryptoStream(ms, rd.CreateEncryptor, CryptoStreamMode.Write)
                            Dim data() As Byte = System.Text.Encoding.UTF8.GetBytes(ZuVerschluesselndes) ' mit dem getauscht
                            cs.Write(data, 0, data.Length)
                            cs.FlushFinalBlock()
                            Dim encdata() As Byte = ms.ToArray()
                            Ausgabe = System.Convert.ToBase64String(encdata)
                        End Using
                    End Using
                End Using
            End If

        Catch ex As Exception
            Dim rb As Boolean  = ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Error_Helper.IsIDE() Then Stop
        End Try
        Return Ausgabe
    End Function
    Public Shared Function Entschluesseln(ByVal ZuEntschluesselndes As String, ByVal Schluessel As String) As String
        Dim rijndaelIvLength As Integer = 16
        Dim Ausgabe As String = ""

        Try
            If Not String.IsNullOrWhiteSpace(ZuEntschluesselndes) AndAlso Not String.IsNullOrWhiteSpace(Schluessel) Then
                Using rd As New RijndaelManaged
                    Using md5 As New MD5CryptoServiceProvider
                        Dim key() As Byte = md5.ComputeHash(Encoding.UTF8.GetBytes(Schluessel))
                        rd.Key = key
                    End Using
                    Dim encdata() As Byte = System.Convert.FromBase64String(ZuEntschluesselndes)
                    Using ms As New MemoryStream(encdata)
                        Dim iv(15) As Byte
                        ms.Read(iv, 0, rijndaelIvLength)
                        rd.IV = iv
                        Using cs As New CryptoStream(ms, rd.CreateDecryptor, CryptoStreamMode.Read)
                            Dim data(ms.Length.ToInteger() - rijndaelIvLength) As Byte
                            Dim i As Integer = cs.Read(data, 0, data.Length) '############HIER#############
                            Ausgabe= System.Text.Encoding.UTF8.GetString(data, 0, i)
                        End Using
                    End Using
                End Using
            End If

        Catch ex As Exception
            Dim rb As Boolean  = ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Error_Helper.IsIDE() Then Stop
        End Try
        Return Ausgabe
    End Function
#End Region

#Region "Eigenschaften"
    Public Shared ReadOnly Property EncryptedString() As String
        Get
            Return EncryptedString_
        End Get
    End Property
    Public Shared ReadOnly Property DecryptedString() As String
        Get
            Return DecryptedString_
        End Get
    End Property
    Public Shared Function CreateMD5Sum(ByVal filename As String) As String
        Dim md5checksum As String = ""
        Using md5 As MD5 = MD5.Create()
            Using stream As FileStream = File.OpenRead(filename)
                md5checksum = BitConverter.ToString(md5.ComputeHash(stream)).Replace("-", String.Empty)
            End Using
        End Using
        Return md5checksum
    End Function
#End Region

    'Public Shared Sub CheckApplicationChecksum()
    '    Dim query_cccac1up As String = ""
    '   'Dim cCrypt As cCrypt = New cCrypt

    '    Error_cCrypt.Encrypt(256, Application.ProductName)
    '    Dim application_name As String = Error_cCrypt.EncryptedString

    '    Error_cCrypt.Encrypt(256, My.Application.Info.Version.ToString)
    '    Dim application_version As String = Error_cCrypt.EncryptedString

    '    Dim application_checksum As String = ""

    '    Dim md5checksum As String = Error_cCrypt.CreateMD5Sum(Application.ExecutablePath)

    '    Error_cCrypt.Encrypt(256, String.Concat(Application.ProductName, My.Application.Info.Version.ToString))
    '    Dim unique As String = Error_cCrypt.EncryptedString

    '    If Not clsError_Helper.IsIDE Then
    '        Using cccac1up As New Helper_DBconnect
    '            If cccac1up.connect(clsError_DbConnectLocal.SelectDatabase.Updater) Then
    '                query_cccac1up = "SELECT * FROM `" & clsError_DbConnectLocal.db_table_checksums & "` WHERE "
    '                query_cccac1up = query_cccac1up & "`application` LIKE ?application? AND "
    '                query_cccac1up = query_cccac1up & "`version` LIKE ?version? "
    '                query_cccac1up = query_cccac1up & ";"
    '                cccac1up.cmd.CommandText = query_cccac1up
    '                cccac1up.cmd.Parameters.Clear()
    '                cccac1up.cmd.Parameters.AddWithValue("?application?", application_name)
    '                cccac1up.cmd.Parameters.AddWithValue("?version?", application_version)
    '                'Debug.Print(HelperDB.ParameterQuery(cccac1up))
    '                Using reader_cccac1up As MySqlDataReader = cccac1up.cmd.ExecuteReader
    '                    While reader_cccac1up.Read()
    '                        Application.DoEvents()
    '                        application_checksum = Helper_Convert.ConvertToString(reader_cccac1up("checksum"))
    '                    End While
    '                End Using
    '                If String.IsNullOrEmpty(application_checksum) Then
    '                    query_cccac1up = "INSERT INTO `" & clsError_DbConnectLocal.db_table_checksums & "` ("
    '                    query_cccac1up = query_cccac1up & "`unique`, "
    '                    query_cccac1up = query_cccac1up & "`application`, "
    '                    query_cccac1up = query_cccac1up & "`version`, "
    '                    query_cccac1up = query_cccac1up & "`checksum` "
    '                    query_cccac1up = query_cccac1up & ") VALUES ( "
    '                    query_cccac1up = query_cccac1up & "?unique?, "
    '                    query_cccac1up = query_cccac1up & "?application?, "
    '                    query_cccac1up = query_cccac1up & "?version?, "
    '                    query_cccac1up = query_cccac1up & "?checksum? "
    '                    query_cccac1up = query_cccac1up & ") ON DUPLICATE KEY UPDATE "
    '                    query_cccac1up = query_cccac1up & "`application` = ?application?, "
    '                    query_cccac1up = query_cccac1up & "`version` = ?version?, "
    '                    query_cccac1up = query_cccac1up & "`checksum` = ?checksum? "
    '                    query_cccac1up = query_cccac1up & ";"
    '                    cccac1up.cmd.CommandText = query_cccac1up
    '                    cccac1up.cmd.Parameters.Clear()
    '                    cccac1up.cmd.Parameters.AddWithValue("?unique?", unique)
    '                    cccac1up.cmd.Parameters.AddWithValue("?application?", application_name)
    '                    cccac1up.cmd.Parameters.AddWithValue("?version?", application_version)
    '                    cccac1up.cmd.Parameters.AddWithValue("?checksum?", md5checksum)
    '                    'Debug.Print(HelperDB.ParameterQuery(cccac1up))
    '                    ri = cccac1up.cmd.ExecuteNonQuery
    '                Else
    '                    If String.Compare(md5checksum, application_checksum, False) <> 0 Then
    '                        MessageBoxEx.Show("Die Prüfsumme der Anwendung ist nicht korrekt." & Environment.NewLine & "" & Environment.NewLine & "Bitte prüfen Sie, ob die Anwendung manipuliert wurde" & Environment.NewLine & " und informieren Sie Ihren Administrator." & Environment.NewLine & "" & Environment.NewLine & "Die Anwendung wird jetzt beendet!", "Prüfsummenfehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        Application.Exit()
    '                    End If
    '                End If

    '            End If
    '        End Using
    '    End If

    'End Sub

End Class