'   Создание оболочки шифрования
' Подключите пространство имен криптографии в начале файла.
Imports System.Security.Cryptography
' Создайте класс для инкапсуляции методов шифрования и расшифровки.
Public NotInheritable Class Simple3Des
    ' Добавьте закрытое поле для хранения поставщика служб шифрования 3DES.
    Private TripleDes As New TripleDESCryptoServiceProvider
    ' Добавьте закрытый метод, который создает массив байтов указанной длины из хэша указанного ключа.
    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte = _
            System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    ' Добавьте конструктор для инициализации поставщика служб шифрования 3DES.
    ' Параметр key управляет методами EncryptData и DecryptData.
    Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub
    ' Добавьте открытый метод, который шифрует строку.
    Public Function EncryptData(ByVal plaintext As String) As String
        If Len(Trim(plaintext)) = 0 Then
            Return plaintext
        Else
            ' Convert the plaintext string to a byte array.
            Dim plaintextBytes() As Byte = _
                System.Text.Encoding.Unicode.GetBytes(plaintext)

            ' Create the stream.
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream.
            Dim encStream As New CryptoStream(ms, _
                TripleDes.CreateEncryptor(), _
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()

            ' Convert the encrypted stream to a printable string.
            Return Convert.ToBase64String(ms.ToArray)
        End If
    End Function

    ' Добавьте открытый метод, который расшифровывает строку.
    Public Function DecryptData(ByVal encryptedtext As String) As String
        If Len(Trim(encryptedtext)) = 0 Then
            Return encryptedtext
        Else
            Dim rt1 As String
            ' Convert the encrypted text string to a byte array.
            Try
                Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)
                ' Create the stream.
                Dim ms As New System.IO.MemoryStream
                ' Create the decoder to write to the stream.
                Dim decStream As New CryptoStream(ms, _
                    TripleDes.CreateDecryptor(), _
                    System.Security.Cryptography.CryptoStreamMode.Write)
                ' Use the crypto stream to write the byte array to the stream.
                decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
                decStream.FlushFinalBlock()
                rt1 = System.Text.Encoding.Unicode.GetString(ms.ToArray)
            Catch ex As Exception
                rt1 = ""
            End Try
            ' Convert the plaintext stream to a string.
            Return rt1
        End If
    End Function
End Class