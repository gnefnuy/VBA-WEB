Attribute VB_Name = "BDemo"
Option Compare Database
Option Explicit
'
' BDemo V1.1.2
' 关于BCrypt和BStorage中函数实现的各种示例。
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Cryptography
'
' 需要:
'   Module  BCrypt
'   Module  BStorage
'   Table   Library 这个是access的表
'   Query   LibraryContent
'

' 打开列出保存在Library表中的加密内容的查询。
' 用于加密的密钥是一个空格。
'
' 加密内容使用以下两个函数进行解密：
'   VDecryptBase64
'   VDecryptBinary
'
' Typical query:
'
'   PARAMETERS
'       [Key] LongText;
'   SELECT
'       Library.Id,
'       Library.Date,
'       VDecryptBase64([ContentBase64],[Key]) AS Base64Content,
'       VDecryptBinary([ContentBinary],[Key]) AS BinaryContent
'   FROM
'       Library;
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub ListLibraryContent()

    Const QueryName         As String = "LibraryContent"
    Const ParameterName     As String = "Key"
    Const ParameterValue    As String = " "
    
    Dim Expression          As String
    
    ' Build string expression for the parameter value.
    Expression = """" & ParameterValue & """"
    ' Preset the parameter.
    DoCmd.SetParameter ParameterName, Expression
    ' Open the query for display.
    DoCmd.OpenQuery QueryName
    
End Sub

' Read, for an ID, the encrypted text from a text field.
' Returns the decrypted text if the key is right.
' Returns an empty string, if the ID is not found, or the key is wrong.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadEncryptedBase64( _
    ByVal Id As Long, _
    ByVal Key As String) _
    As String
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBase64"
    
    Dim EncryptedText       As String
    Dim Text                As String
    
    EncryptedText = ReadTextField(CurrentDb, DefaultTableName, DefaultFieldName, Id)
    Text = Decrypt(EncryptedText, (Key))
    
    ReadEncryptedBase64 = Text
    
End Function

' Read, for an ID, the encrypted text from a  stored hash value of the password.
' Returns a byte array if a hash value is found.
' Returns an empty byte array, if the ID is not found, or the password is empty.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadEncryptedBinary( _
    ByVal Id As Long, _
    ByVal Key As String) _
    As String
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBinary"
    
    Dim EncryptedData()     As Byte
    Dim DecryptedData()     As Byte
    Dim Text                As String
    
    EncryptedData = ReadBinaryField(CurrentDb, DefaultTableName, DefaultFieldName, Id)
    
    If DecryptData(EncryptedData, (Key), DecryptedData) Then
        ' Return the decrypted text.
        Text = DecryptedData
    Else
        ' Return an empty string.
    End If
    
    ReadEncryptedBinary = Text
    
End Function

' 为指定的ID保存传递给文本字段的文本的AES加密值。
' 必须传递一个密钥。
' 如果参数Text为空，则将存储Null。
' 如果成功，返回True。
' 如果没有传递密钥，返回False。
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveEncryptedBase64( _
    ByVal Id As Long, _
    ByVal Key As String, _
    Optional ByVal Text As String) _
    As Boolean
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBase64"
    
    Dim EncryptedText       As String
    Dim Success             As Boolean
    
    EncryptedText = Encrypt(Text, Key)
    Success = SaveTextField(CurrentDb, DefaultTableName, DefaultFieldName, Id, EncryptedText)
    
    SaveEncryptedBase64 = Success
    
End Function

' Save, for an ID, the AES encrypted value of the text passed to a binary field.
' A key must be passed.
' If argument Text is empty, Null will be stored.
' Returns True if success.
' Returns False if no key is passed.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveEncryptedBinary( _
    ByVal Id As Long, _
    ByVal Key As String, _
    Optional ByVal Text As String) _
    As Boolean
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBinary"
    
    Dim EncryptedData()     As Byte
    
    Dim Success             As Boolean
    
    If EncryptData((Text), (Key), EncryptedData) = True Then
        Success = SaveBinaryField(CurrentDb, DefaultTableName, DefaultFieldName, Id, EncryptedData)
    End If
    
    SaveEncryptedBinary = Success
    
End Function

