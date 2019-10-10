' An implementation of AES-256-CBC in VBScript.

Set utf8 = CreateObject("System.Text.UTF8Encoding")
Set b64Enc = CreateObject("System.Security.Cryptography.ToBase64Transform")
Set b64Dec = CreateObject("System.Security.Cryptography.FromBase64Transform")
set aes = CreateObject("System.Security.Cryptography.RijndaelManaged")


' Return the minimum value between two integer values.
'
' Arguments:
'   a (Long): An integer.
'   b (Long): Another integer.
'
' Return:
'   Long: Minimum of the two integer values.
Function Min(a, b)
    Min = a
    If b < a Then Min = b
End Function


' Convert a byte array to a Base64 string representation of it.
'
' Arguments:
'   bytes (Byte()): Byte array.
'
' Returns:
'   String: Base64 representation of the input byte array.
Function B64Encode(bytes)
    blockSize = b64Enc.InputBlockSize
    For offset = 0 To LenB(bytes) - 1 Step blockSize
        length = Min(blockSize, LenB(bytes) - offset)
        b64Block = b64Enc.TransformFinalBlock((bytes), offset, length)
        result = result & utf8.GetString((b64Block))
    Next
    B64Encode = result
End Function


' Convert a Base64 string to a byte array.
'
' Arguments:
'   b64Str (String): Base64 string.
'
' Returns:
'   Byte(): An array of bytes that the Base64 string decodes to.
Function B64Decode(b64Str)
    bytes = utf8.GetBytes_4(b64Str)
    B64Decode = b64Dec.TransformFinalBlock((bytes), 0, LenB(bytes))
End Function


' Encrypt a given plaintext with a given key.
'
' The encryption key must be 256-bit (32-byte) long. It must be provided
' as a Base64 encoded string. On macOS or Linux, enter this command to
' generate a Base64 encoded 256-bit key:
'
'   head -c32 /dev/urandom | base64
'
' The return value of Encrypt() is a single string that contains a
' Base 64 encoded 128-bit randomly generated initiazation vector (IV)
' and the ciphertext joined with a colon.
'
' A 256-bit key after Base64 encoding contains 44 characters including
' an '=' padding at the end.
'
' A 128-bit IV after Base64 encoding contains 24 characters including
' two '=' characters as padding at the end.
'
' Arguments:
'   plaintext (String): Text to be encrypted.
'   key (String): Encryption key encoded as a Base64 string.
'
' Returns:
'   String: IV and ciphertext joined with a colon in between.
Function Encrypt(plaintext, key)
    aes.GenerateIV()
    set aesEnc = aes.CreateEncryptor_2(B64Decode(key), aes.IV)
    bytes = utf8.GetBytes_4(plaintext)
    bytes = aesEnc.TransformFinalBlock((bytes), 0, LenB(bytes))
    Encrypt = B64Encode(aes.IV) & ":" & B64Encode(bytes)
End Function


' Decrypt a given ciphertext with a given IV and key.
'
' Both IV and key must be encoded in Base64. Both are provided together
' as a single string with values separated by a colon. See the comment
' for Encrypt() function to read more about their formats.
'
' Arguments:
'   ivAndCipherText (String): Colon separated IV and ciphertext.
'   key (String): Encryption key encoded as a Base64 string.
'
' Returns:
'   String: Plaintext that the given ciphertext decrypts to.
Function Decrypt(ivCiphertext, key)
    tokens = Split(ivCiphertext, ":")
    set aesDec = aes.CreateDecryptor_2(B64Decode(key), B64Decode(tokens(0)))
    bytes = B64Decode(tokens(1))
    bytes = aesDec.TransformFinalBlock((bytes), 0, LenB(bytes))
    Decrypt = utf8.GetString((bytes))
End Function


' Show results of Encrypt() and Decrypt() calls for demo purpose.
Function CryptoDemo
    demoPlaintext = "foo"
    demoAESKey = "CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg="
    demoIVCiphertext = "2dmVWVT++xbgaDq7ktdUNg==:LEd9iJAHo6bhkpkY/CcrlQ=="

    encrypted1 = Encrypt(demoPlaintext, demoAESKey)
    decrypted1 = Decrypt(encrypted1, demoAESKey)
    encrypted2 = demoIVCiphertext
    decrypted2 = Decrypt(encrypted2, demoAESKey)

    CryptoDemo = "demoAESKey: " & demoAESKey & vbCrLf & _
                 "encrypted1: " & encrypted1 & vbCrLf & _
                 "decrypted1: " & decrypted1 & vbCrLf & _
                 "encrypted2: " & encrypted2 & vbCrLf & _
                 "decrypted2: " & decrypted2 & vbCrLf
End Function


' Show interesting properties of cryptography objects used here.
Function CryptoInfo
    set enc = aes.CreateEncryptor_2(aes.Key, aes.IV)
    set dec = aes.CreateDecryptor_2(aes.Key, aes.IV)

    CryptoInfo = "aes.BlockSize: " & aes.BlockSize & vbCrLf & _
                 "aes.FeedbackSize: " & aes.FeedbackSize & vbCrLf & _
                 "aes.KeySize: " & aes.KeySize & vbCrLf & _
                 "aes.Mode: " & aes.Mode & vbCrLf & _
                 "aes.Padding: " & aes.Padding & vbCrLf &_
                 "aesEnc.InputBlockSize: " & enc.InputBlockSize & vbCrLf & _
                 "aesEnc.OutputBlockSize: " & enc.OutputBlockSize & vbCrLf & _
                 "aesDec.InputBlockSize: " & enc.InputBlockSize & vbCrLf & _
                 "aesDec.OutputBlockSize: " & enc.OutputBlockSize & vbCrLf & _
                 "b64Enc.InputBlockSize: " & b64Enc.InputBlockSize & vbCrLf & _
                 "b64Eec.OutputBlockSize: " & b64Enc.OutputBlockSize & vbCrLf & _
                 "b64Dec.InputBlockSize: " & b64Dec.InputBlockSize & vbCrLf & _
                 "b64Dec.OutputBlockSize: " & b64Dec.OutputBlockSize & vbCrLf
End Function
