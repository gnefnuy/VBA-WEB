' 数据库文件路径
Private Const DbFile As String = "database.db3"
' 执行数据库查询，获取第一行第一列内容
Public Function ExecuteScalar(ByVal sql As String) As String
    Dim result As String
    Dim myDbHandleV2 As LongPtr
    Dim myStmtHandle As LongPtr
    Dim RetVal As Long
    
    ' 初始化SQLite
    Dim InitReturn As Long
    InitReturn = SQLite3Initialize("D:\App\JBHTools\x64")
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "初始化SQLite错误:" & Err.LastDllError
        Exit Function
    End If
    
    ' 打开数据库
    RetVal = SQLite3OpenV2(DbFile, myDbHandleV2, SQLITE_OPEN_READONLY, "")
    If RetVal <> SQLITE_INIT_OK Then
        Debug.Print "打开数据库错误: " & SQLite3ErrMsg(myDbHandleV2)
        Exit Function
    End If
    
    ' 准备SQL语句
    RetVal = SQLite3PrepareV2(myDbHandleV2, sql, myStmtHandle)
    If RetVal <> SQLITE_INIT_OK Then
        Debug.Print "准备SQL语句错误: " & SQLite3ErrMsg(myDbHandleV2)
        SQLite3Close myDbHandleV2
        Exit Function
    End If
    
    ' 执行SQL语句
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        result = SQLite3ColumnText(myStmtHandle, 0)
    Else
        Debug.Print "执行SQL语句错误: " & SQLite3ErrMsg(myDbHandleV2)
        result = ""
    End If
    
    ' 清理资源
    RetVal = SQLite3Finalize(myStmtHandle)
    RetVal = SQLite3Close(myDbHandleV2)
    
    ExecuteScalar = result
End Function

' 将二维数组导入到数据库
Public Sub ImportArrayToDatabase(sql As String, dataArray As Variant)
    Dim myDbHandleV2 As LongPtr
    Dim myStmtHandle As LongPtr
    Dim RetVal As Long
    Dim i As Long
    Dim j As Long
    Dim Value As Variant
    
    ' 初始化SQLite
    Dim InitReturn As Long
    InitReturn = SQLite3Initialize("D:\App\x64")
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "初始化SQLite错误:" & Err.LastDllError
        Exit Sub
    End If
    
    ' 打开数据库
    RetVal = SQLite3OpenV2(DbFile, myDbHandleV2, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE, "")
    If RetVal <> SQLITE_INIT_OK Then
        Debug.Print "打开数据库错误: " & SQLite3ErrMsg(myDbHandleV2)
        Exit Sub
    End If
    
    ' 准备SQL语句
    RetVal = SQLite3PrepareV2(myDbHandleV2, sql, myStmtHandle)
    If RetVal <> SQLITE_INIT_OK Then
        Debug.Print "准备SQL语句错误: " & SQLite3ErrMsg(myDbHandleV2)
        SQLite3Close myDbHandleV2
        Exit Sub
    End If
    
    ' 导入数组数据
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            Value = dataArray(i, j)
            RetVal = SQLite3BindText(myStmtHandle, j + 1, Value)
            If RetVal <> SQLITE_INIT_OK Then
                Debug.Print "绑定参数错误: " & SQLite3ErrMsg(myDbHandleV2)
                SQLite3Finalize myStmtHandle
                SQLite3Close myDbHandleV2
                Exit Sub
            End If
        Next j
        
        RetVal = SQLite3Step(myStmtHandle)
        If RetVal <> SQLITE_DONE Then
            Debug.Print "执行SQL语句错误: " & SQLite3ErrMsg(myDbHandleV2)
            SQLite3Finalize myStmtHandle
            SQLite3Close myDbHandleV2
            Exit Sub
        End If
        
        RetVal = SQLite3Reset(myStmtHandle)
        If RetVal <> SQLITE_INIT_OK Then
            Debug.Print "重置语句错误: " & SQLite3ErrMsg(myDbHandleV2)
            SQLite3Finalize myStmtHandle
            SQLite3Close myDbHandleV2
            Exit Sub
        End If
    Next i
    
    ' 清理资源
    RetVal = SQLite3Finalize(myStmtHandle)
    RetVal = SQLite3Close(myDbHandleV2)
End Sub