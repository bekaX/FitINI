Imports System.IO
Imports System.Text

' FitINI Library
' A simple library for accessing ini file and ini style file
' 
' 1.2.2

Namespace FitINI

    ''' <summary>
    ''' 表示一个 INI 对象
    ''' </summary>
    Public Class INI
        Implements IEnumerable(Of KeyValuePair(Of String, Section))
        Protected _Sections As New Dictionary(Of String, Section)
        Protected _Default As New Section()
        Protected _CommentPrefix As String

        '默认的注释前缀
        Public Const DefaultCommentPrefix = ";"

        ''' <summary>
        ''' 获取或设置此 INI 中的某一个区块 (不存在时返回 Nothing)
        ''' </summary>
        ''' <param name="name">区块的名字</param>
        ''' <returns>返回名字为 name 的区块对象</returns>
        Default Public Property Section(name As String) As Section
            Get
                Try
                    Return _Sections.Item(name)
                Catch ex As KeyNotFoundException
                    Return Nothing
                End Try
            End Get
            Set(ByVal value As Section)
                If value Is Nothing Then
                    Throw New ArgumentNullException()
                End If
                If name Is Nothing Then name = ""
                _Sections.Item(name) = value
            End Set
        End Property

        ''' <summary>
        ''' 缺省状态下的区块(没有指定区块名的键值对和数据会被存在这里)
        ''' </summary>
        ''' <returns>返回缺省的区块</returns>
        Public ReadOnly Property DefaultSection() As Section
            Get
                Return _Default
            End Get
        End Property

        ''' <summary>
        ''' 访问此 INI 中的条目(不存在的区块和键值对会被自动添加)
        ''' </summary>
        ''' <param name="name">区块名字</param>
        ''' <param name="key">键名</param>
        ''' <returns>返回值或 Nothing (区块或键不存在时)</returns>
        Public Property Entry(name As String, key As String) As String
            Get
                Dim section As Section = Nothing
                If Not _Sections.TryGetValue(name, section) Then Return Nothing
                Return section.Item(key)
            End Get
            Set(value As String)
                If Not _Sections.ContainsKey(name) Then
                    _Sections.Add(name, New Section())
                End If
                _Sections.Item(name).Item(key) = value
            End Set
        End Property

        ''' <summary>
        ''' 获取此 INI 的行注释前缀字符
        ''' </summary>
        ''' <returns></returns>
        Public Property CommentPrefix() As String
            Get
                Return _CommentPrefix
            End Get
            Set(value As String)
                If value Is Nothing Then
                    Throw New ArgumentNullException
                End If
                _CommentPrefix = value
            End Set
        End Property

        ''' <summary>
        ''' 构造一个 INI 对象
        ''' </summary>
        Public Sub New()
            _CommentPrefix = DefaultCommentPrefix
        End Sub

        ''' <summary>
        ''' 新增一个区块(如果已存在则不会新建)
        ''' </summary>
        ''' <param name="name">区块的名字</param>
        ''' <returns>返回这个区块</returns>
        Public Function Add(name As String) As Section
            If Not _Sections.ContainsKey(name) Then
                _Sections.Add(name, New Section())
            End If
            Return _Sections.Item(name)
        End Function

        ''' <summary>
        ''' 获取此 INI 中的条目的值
        ''' </summary>
        ''' <param name="name">区块名字</param>
        ''' <param name="key">键名</param>
        ''' <param name="def">当区块或键不存在时的默认返回值</param>
        ''' <returns>返回结果</returns>
        Public Function GetEntry(name As String,
                                 key As String,
                                 Optional def As String = Nothing) As String
            Dim section As Section = Nothing
            If Not _Sections.TryGetValue(name, section) Then Return def
            Dim value As String = section.Item(key)
            Return IIf(value Is Nothing, def, value)
        End Function

        ''' <summary>
        ''' 检测此 INI 中是否包含某一区块
        ''' </summary>
        ''' <param name="sectionName">区块的名字</param>
        ''' <returns>包含时返回 True, 不包含返回 False</returns>
        Public Function Contains(sectionName As String) As Boolean
            Return _Sections.ContainsKey(sectionName)
        End Function

        ''' <summary>
        ''' 移除某一个区块
        ''' </summary>
        ''' <param name="sectionName">区块的名字</param>
        ''' <returns>成功移除返回 True, 否则返回 False</returns>
        Public Function Remove(sectionName As String) As Boolean
            Return _Sections.Remove(sectionName)
        End Function

        ''' <summary>
        ''' 修改某一个区块的名字
        ''' </summary>
        ''' <param name="oldName">区块的名字</param>
        ''' <param name="newName">区块的新名字</param>
        ''' <returns>成功修改返回 True, 否则返回 False</returns>
        Public Function Rename(oldName As String, newName As String) As Boolean
            If oldName Is Nothing Then oldName = ""
            If newName Is Nothing Then newName = ""
            If oldName <> newName AndAlso _Sections.ContainsKey(oldName) Then
                Dim section = _Sections.Item(oldName)
                _Sections.Remove(oldName)
                _Sections.Add(newName, section)
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' 创建一个当前 INI 对象的副本(深拷贝)
        ''' </summary>
        ''' <returns></returns>
        Public Function Clone() As INI
            Dim ini As New INI() With {
                    ._CommentPrefix = Me.CommentPrefix,
                    ._Default = Me._Default.Clone()
                }
            For Each kv In Me._Sections
                ini._Sections.Add(kv.Key, kv.Value.Clone())
            Next
            Return ini
        End Function

        ''' <summary>
        ''' 清空 INI 对象中的所有数据
        ''' </summary>
        Public Sub Clear()
            _Default.Clear()
            Me.ClearSections()
        End Sub

        ''' <summary>
        ''' 清空并移除所有区块 (缺省的区块除外)
        ''' </summary>
        Public Sub ClearSections()
            For Each v In _Sections.Values
                v.Clear()
            Next
            _Sections.Clear()
        End Sub

        ''' <summary>
        ''' 移除所有的注释和文本 (除键值对数据以外的内容)
        ''' </summary>
        Public Sub RemoveAllComments()
            _Default.RemoveAllComments()
            For Each v In _Sections.Values
                v.RemoveAllComments()
            Next
        End Sub

        ''' <summary>
        ''' 获取所有的区块名
        ''' </summary>
        ''' <returns></returns>
        Public Function GetSectionNames() As String()
            Return _Sections.Keys.ToArray()
        End Function

        ''' <summary>
        ''' 按行生成文本的列表
        ''' </summary>
        ''' <returns>返回包含每行内容的列表</returns>
        Public Function ToList() As List(Of String)
            Dim list As New List(Of String)
            list.AddRange(_Default.ToList(_CommentPrefix))
            For Each kv In _Sections
                list.Add(String.Format("[{0}]", kv.Key))
                list.AddRange(kv.Value.ToList(_CommentPrefix))
            Next
            Return list
        End Function

        ''' <summary>
        ''' 根据内容生成字符串
        ''' </summary>
        ''' <returns>返回文本内容的字符串</returns>
        Public Function ToContentString() As String
            Dim text As New StringBuilder
            If _Default.Count <> 0 Then text.Append(_Default.ToContentString(_CommentPrefix))
            If _Default.Count <> 0 AndAlso _Sections.Count <> 0 Then text.AppendLine()
            Dim index As Integer = 1
            For Each kv In _Sections
                text.Append(String.Format("[{0}]{1}", kv.Key, vbCrLf))
                text.Append(kv.Value.ToContentString(_CommentPrefix))
                If index <> _Sections.Count Then
                    text.AppendLine()
                    index += 1
                End If
            Next
            Return text.ToString()
        End Function

        ''' <summary>
        ''' 将本对象的数据写入到文件中保存
        ''' </summary>
        ''' <param name="path">文件保存的路径</param>
        Public Sub SaveToFile(path As String)
            Using file As New System.IO.StreamWriter(path)
                Dim list As List(Of String) = ToList()
                Dim index As Integer = 1
                For Each line In list
                    If index = list.Count Then
                        file.Write(line)
                    Else
                        file.WriteLine(line)
                    End If
                    index += 1
                Next
            End Using
        End Sub

        Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, Section)) _
            Implements IEnumerable(Of KeyValuePair(Of String, Section)).GetEnumerator
            Return DirectCast(_Sections, IEnumerable(Of KeyValuePair(Of String, Section))).GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

        '
        ' 数据 / 文件解析
        '

        ''' <summary>
        ''' 生成文件的行迭代器
        ''' </summary>
        ''' <param name="path">文件的路径</param>
        ''' <returns></returns>
        Private Shared Iterator Function YieldFile(path As String) As IEnumerable(Of String)
            Using file As New System.IO.StreamReader(path)
                While Not file.EndOfStream
                    Yield file.ReadLine
                End While
            End Using
        End Function

        ''' <summary>
        ''' 根据内容生成一个 INI 对象
        ''' </summary>
        ''' <param name="enumerable">可迭代内容</param>
        ''' <returns></returns>
        Protected Shared Function BuildINI(enumerable As IEnumerable(Of String), arg As LoadOption) As INI
            Dim ini As New INI()
            Dim sec As New Section()
            Dim secName As String = Nothing
            '按行分析文本
            For Each line In enumerable
                If Left(line, 1) = "[" AndAlso Right(line, 1) = "]" Then
                    '区块开头
                    If secName Is Nothing Then
                        '如果还没有遇到过区块开头放入默认区块
                        ini._Default = sec
                    Else
                        '已经读取过区块就按照区块名存入对象中
                        ini.Section(secName) = sec
                    End If
                    secName = Mid(line, 2, line.Length - 2)
                    sec = New Section()
                    Continue For
                ElseIf line.StartsWith(";") OrElse line.StartsWith("#") OrElse line.StartsWith(ini.CommentPrefix) Then
                    '注释
                    If line.StartsWith(ini.CommentPrefix) Then
                        line = line.Substring(1)
                    Else
                        line = line.Substring(ini.CommentPrefix.Length)
                    End If
                    sec.AddComment(line)
                Else
                    Dim kv As String() = line.Split("=", 2, StringSplitOptions.None)
                    If kv.Length = 2 Then
                        '键值对
                        sec.Item(kv(0)) = kv(1)
                    Else
                        '其他文本
                        Select Case arg.UnknowLineOption
                            Case LoadOption.LineOption.Keep
                                sec.AddComment(line, False)
                            Case LoadOption.LineOption.AsMultiLine
                                sec.AppendToLast(vbCrLf + line)
                            Case LoadOption.LineOption.AsMultiLineCombined
                                sec.AppendToLast(line)
                            Case LoadOption.LineOption.ForceToComment
                                sec.AddComment(line, True)
                            Case LoadOption.LineOption.Drop
                                Continue For
                        End Select
                    End If
                End If
            Next
            '保存最后一个读取的区块
            If secName Is Nothing Then
                '如果还没有遇到过区块开头放入默认区块
                ini._Default = sec
            Else
                '已经读取过区块就按照区块名存入对象中
                ini.Section(secName) = sec
            End If
            Return ini
        End Function

        ''' <summary>
        ''' 从文件加载并生成一个 INI 对象，当文件IO错误时内容为空
        ''' </summary>
        ''' <param name="path">文件路径</param>
        ''' <returns></returns>
        Public Shared Function LoadFromFile(path As String) As INI
            Try
                Return BuildINI(YieldFile(path), New LoadOption With {
                    .UnknowLineOption = LoadOption.LineOption.AsMultiLine
                })
            Catch ex As System.IO.IOException
                Return New INI()
            End Try
        End Function

        ''' <summary>
        ''' 从文件加载并生成一个 INI 对象，并指定加载所用的选项
        ''' </summary>
        ''' <param name="path">文件路径</param>
        ''' <param name="arg">加载的选项</param>
        ''' <returns></returns>
        Public Shared Function LoadFromFile(path As String, arg As LoadOption) As INI
            Try
                Dim ini As INI = BuildINI(YieldFile(path), arg)
                ini.CommentPrefix = arg.DataCommentPrefix
                Return ini
            Catch ex As System.IO.IOException
                If arg.IgnoreFileMissed AndAlso Not File.Exists(path) Then
                    Return New INI() With {
                        .CommentPrefix = arg.DataCommentPrefix
                    }
                End If
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' 简化地从文件加载并生成一个 INI 对象，忽略任何非键值对的行
        ''' </summary>
        ''' <param name="path">文件路径</param>
        ''' <returns></returns>
        Public Shared Function LoadFromFileSimply(path As String) As INI
            Dim ini As New INI()
            Try
                Dim sec As New Section()
                Dim secName As String = Nothing
                '按行分析文本
                For Each line In YieldFile(path)
                    If Left(line, 1) = "[" AndAlso Right(line, 1) = "]" Then
                        '区块开头
                        If secName Is Nothing Then
                            ini._Default = sec
                        Else
                            ini.Section(secName) = sec
                        End If
                        secName = Mid(line, 2, line.Length - 2)
                        sec = New Section()
                    Else
                        Try
                            Dim kv As String() = line.Split("=", 2, StringSplitOptions.None)
                            sec.Item(kv(0)) = kv(1)
                        Catch ex As IndexOutOfRangeException
                            Continue For
                        Catch ex As ArgumentOutOfRangeException
                            Continue For
                        End Try
                    End If
                Next
                '保存最后一个读取的区块
                If secName Is Nothing Then
                    '如果还没有遇到过区块开头放入默认区块
                    ini._Default = sec
                Else
                    '已经读取过区块就按照区块名存入对象中
                    ini.Section(secName) = sec
                End If
            Catch ex As IOException

            End Try
            Return ini
        End Function

        ''' <summary>
        ''' 从列表生成一个 INI 对象
        ''' </summary>
        ''' <param name="list">要解析的列表</param>
        ''' <param name="arg">加载的选项</param>
        ''' <returns></returns>
        Public Shared Function LoadFromList(ByRef list As List(Of String),
                                            Optional arg As LoadOption = Nothing) As INI
            If arg Is Nothing Then arg = New LoadOption With {
                .UnknowLineOption = LoadOption.LineOption.ForceToComment
            }
            Return BuildINI(list, arg)
        End Function

        ''' <summary>
        ''' 由字符串生成一个 INI 对象
        ''' </summary>
        ''' <param name="content">要解析的字符串</param>
        ''' <param name="arg">加载的选项</param>
        ''' <returns></returns>
        Public Shared Function LoadFromString(content As String,
                                              Optional arg As LoadOption = Nothing) As INI
            '切割字符串
            Dim lines As String() = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            If arg Is Nothing Then arg = New LoadOption With {
                .UnknowLineOption = LoadOption.LineOption.AsMultiLine
            }
            Return BuildINI(lines, arg)
        End Function

        ''' <summary>
        ''' 直接读取键值对并返回值 (不存在时返回 Nothing)
        ''' </summary>
        ''' <param name="enumerable">可迭代的对象</param>
        ''' <param name="name">区块名</param>
        ''' <param name="key">键名</param>
        ''' <returns>根据区块和键名找出的值</returns>
        Public Shared Function DirectLoad(ByRef enumerable As IEnumerable(Of String),
                                          name As String,
                                          key As String) As String
            Dim foundSection As Boolean = False
            For Each line In enumerable
                If foundSection Then
                    Dim kv As String() = line.Split("=", 2, StringSplitOptions.None)
                    Try
                        If key = kv(0) Then Return kv(1)
                    Catch ex As ArgumentOutOfRangeException
                        Continue For
                    End Try
                ElseIf Left(line, 1) = "[" AndAlso Right(line, 1) = "]" Then
                    '区块开头
                    If name = Mid(line, 2, line.Length - 2) Then foundSection = True
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' 从文件直接读取键值对并返回值 (文件, 区块或键不存在时返回 Nothing), 不会单独处理注释
        ''' </summary>
        ''' <param name="path">文件路径</param>
        ''' <param name="name">区块名</param>
        ''' <param name="key">键名</param>
        ''' <returns>根据区块和键名找出的值</returns>
        Public Shared Function DirectLoadFromFile(path As String,
                                                  name As String,
                                                  key As String) As String
            If Not File.Exists(path) Then
                Return Nothing
            End If
            Return DirectLoad(YieldFile(path), name, key)
        End Function

    End Class

    ''' <summary>
    ''' 解析数据的相关选项
    ''' </summary>
    Public Class LoadOption

        ''' <summary>
        ''' 当遇到非标准类型的行文本时的选项
        ''' </summary>
        Public Enum LineOption
            ''' <summary>
            ''' 保留原文
            ''' </summary>
            Keep
            ''' <summary>
            ''' 合并到上一行（有换行符）
            ''' </summary>
            AsMultiLine
            ''' <summary>
            ''' 合并到上一行（无换行符）
            ''' </summary>
            AsMultiLineCombined
            ''' <summary>
            ''' 转换为注释
            ''' </summary>
            ForceToComment
            ''' <summary>
            ''' 丢弃此行
            ''' </summary>
            Drop
        End Enum

        ''' <summary>
        ''' 当文件不存在时不抛出异常
        ''' </summary>
        ''' <returns></returns>
        Public Property IgnoreFileMissed() As Boolean = True
        ''' <summary>
        ''' 遇到无法解析的行的操作
        ''' </summary>
        ''' <returns></returns>
        Public Property UnknowLineOption() As LineOption = LineOption.Keep
        ''' <summary>
        ''' 要解析的文件中注释的前缀
        ''' </summary>
        ''' <returns></returns>
        Public Property DataCommentPrefix() As String = INI.DefaultCommentPrefix

    End Class

    ''' <summary>
    ''' 表示 INI 对象内部的一个区块
    ''' </summary>
    Public Class Section
        Implements IEnumerable(Of KeyValuePair(Of String, String))
        Protected _Entries As New Dictionary(Of String, String)
        Protected _ItemList As New List(Of Record)

        ''' <summary>
        ''' 数据的种类
        ''' </summary>
        Public Enum RecordType
            KeyValue
            Comment
            Other
        End Enum

        ''' <summary>
        ''' 表示区块中的某条具体数据
        ''' </summary>
        Public Structure Record

            Public Property Type As RecordType
            Public Property Content As String

            ''' <summary>
            ''' 是否为键值对
            ''' </summary>
            ''' <returns></returns>
            Public Function IsKeyValue() As Boolean
                Return Type = RecordType.KeyValue
            End Function

            Public Overrides Function ToString() As String
                Return String.Format("{0}: Type = {1}, Content = {2}", Me.GetType, Type, Content)
            End Function
        End Structure

        ''' <summary>
        ''' 获取或设置此区块中的某一项
        ''' </summary>
        ''' <param name="key">键值对中的键</param>
        ''' <returns>键值对中 key 对应的值</returns>
        Default Public Property Item(key As String) As String
            Get
                If _Entries.ContainsKey(key) Then
                    Return _Entries.Item(key)
                Else
                    Return Nothing
                End If
            End Get
            Set(value As String)
                If value Is Nothing Then value = ""
                If _Entries.ContainsKey(key) Then
                    _Entries.Item(key) = value
                Else
                    _Entries.Add(key, value)
                    _ItemList.Add(New Record With {
                        .Type = RecordType.KeyValue,
                        .Content = key
                    })
                End If
            End Set
        End Property

        ''' <summary>
        ''' 获取此区块内的条目(键值对)总数
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Count() As Integer
            Get
                Return _Entries.Count
            End Get
        End Property

        ''' <summary>
        ''' 构造一个 Section 对象
        ''' </summary>
        Public Sub New()

        End Sub

        Protected Overrides Sub Finalize()
            _Entries.Clear()
            _ItemList.Clear()
        End Sub

        ''' <summary>
        ''' 获取此区块中一个条目的值
        ''' </summary>
        ''' <param name="key">键名</param>
        ''' <param name="def">当键不存在时的默认返回值</param>
        ''' <returns>返回结果</returns>
        Function GetValue(key As String, Optional def As String = Nothing) As String
            Try
                Return _Entries.Item(key)
            Catch ex As KeyNotFoundException
                Return def
            End Try
        End Function

        ''' <summary>
        ''' 检测此区块中是否包含某一条目(键值对)
        ''' </summary>
        ''' <param name="key">键的名字</param>
        ''' <returns>包含时返回 True, 不包含返回 False</returns>
        Public Function Contains(key As String) As Boolean
            Return _Entries.ContainsKey(key)
        End Function

        ''' <summary>
        ''' 移除某一个条目(键值对)
        ''' </summary>
        ''' <param name="key">键的名字</param>
        ''' <returns>成功移除返回 True, 否则返回 False</returns>
        Public Function Remove(key As String) As Boolean
            Try
                _Entries.Remove(key)
                For i As Integer = _ItemList.Count - 1 To 0 Step -1
                    If _ItemList.Item(i).IsKeyValue() AndAlso _ItemList.Item(i).Content = key Then
                        _ItemList.RemoveAt(i)
                    End If
                Next
            Catch ex As Exception
                Return False
            End Try
            Return True
        End Function

        ''' <summary>
        ''' 修改键值中对的键名
        ''' </summary>
        ''' <param name="oldName">键名</param>
        ''' <param name="newName">键的新名字</param>
        ''' <returns>成功修改返回 True, 否则返回 False</returns>
        Public Function Rename(oldName As String, newName As String) As Boolean
            If oldName Is Nothing Then oldName = ""
            If newName Is Nothing Then newName = ""
            If oldName <> newName AndAlso _Entries.ContainsKey(oldName) Then
                For i = 0 To _ItemList.Count - 1
                    Dim e As Record = _ItemList(i)
                    If e.Content = oldName AndAlso e.IsKeyValue Then
                        _ItemList(i) = New Record With {
                            .Content = newName,
                            .Type = e.Type
                        }
                    End If
                Next
                Dim value = _Entries.Item(oldName)
                _Entries.Remove(oldName)
                _Entries.Add(newName, value)
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' 创建一个当前 Section 对象的副本(深拷贝)
        ''' </summary>
        ''' <returns></returns>
        Public Function Clone() As Section
            Dim section = New Section() With {
                ._Entries = New Dictionary(Of String, String)(_Entries),
                ._ItemList = New List(Of Record)(_ItemList)
            }
            Return section
        End Function

        ''' <summary>
        ''' 清空当前 Section 对象的所有内容
        ''' </summary>
        Public Sub Clear()
            _Entries.Clear()
            _ItemList.Clear()
        End Sub

        ''' <summary>
        ''' 将文本内容追加到最后一条内容上（键值对的值/注释/其他文本） 
        ''' 此方法主要用于构造INI时处理多行内容
        ''' 若尝试合并失败会自动跳过
        ''' </summary>
        ''' <param name="content">要追加的内容</param>
        Friend Sub AppendToLast(content As String)
            Try
                Dim rec = _ItemList.Last()
                If content Is Nothing Then content = ""
                If rec.IsKeyValue() Then
                    _Entries.Item(rec.Content) += content
                Else
                    _ItemList.Item(_ItemList.Count - 1) = New Record() With {
                                        .Type = rec.Type,
                                        .Content = rec.Content + content
                    }
                End If
            Catch ex As InvalidOperationException

            End Try
        End Sub

        ''' <summary>
        ''' 添加一条注释
        ''' </summary>
        ''' <param name="content">注释的内容</param>
        Public Sub AddComment(content As String)
            _ItemList.Add(New Record With {
                .Content = content,
                .Type = RecordType.Comment
            })
        End Sub

        ''' <summary>
        ''' 添加一条注释或纯文本
        ''' </summary>
        ''' <param name="content">文本内容</param>
        ''' <param name="isComment">是普通的注释(True)还是非注释的纯文本(False)</param>
        Public Sub AddComment(content As String, isComment As Boolean)
            _ItemList.Add(New Record With {
                .Content = content,
                .Type = IIf(isComment, RecordType.Comment, RecordType.Other)
            })
        End Sub

        ''' <summary>
        ''' 获取所有的注释和文本 (除键值对数据以外的内容)
        ''' </summary>
        ''' <returns>返回字符串数组</returns>
        Public Function GetComments() As String()
            Dim list As New List(Of String)
            For Each record In _ItemList
                For Each rec In _ItemList
                    If Not rec.IsKeyValue() Then list.Add(rec.Content)
                Next
            Next
            Return list.ToArray
        End Function

        ''' <summary>
        ''' 移除所有的注释和文本 (除键值对数据以外的内容)
        ''' </summary>
        Public Sub RemoveAllComments()
            Dim list As New List(Of Integer)
            For i As Integer = 0 To _ItemList.Count - 1
                If Not _ItemList.Item(i).IsKeyValue Then list.Add(i)
            Next
            For Each i In list
                _ItemList.RemoveAt(i)
            Next
        End Sub

        ''' <summary>
        ''' 获取所有的键
        ''' </summary>
        ''' <returns></returns>
        Public Function GetKeys() As String()
            Return _Entries.Keys.ToArray()
        End Function

        ''' <summary>
        ''' 按行生成文本的列表
        ''' </summary>
        ''' <param name="commentPrefix">表示行注释的前缀字符</param>
        ''' <returns>返回包含每行内容的列表</returns>
        Public Function ToList(Optional commentPrefix As String = ";") As List(Of String)
            Dim list As New List(Of String)
            For Each rec In _ItemList
                If rec.IsKeyValue Then
                    list.Add(String.Format("{0}={1}", rec.Content, _Entries.Item(rec.Content)))
                ElseIf rec.Type = RecordType.Comment Then
                    list.Add(String.Format("{0}{1}", commentPrefix, rec.Content))
                Else
                    list.Add(rec.Content)
                End If
            Next
            Return list
        End Function

        ''' <summary>
        ''' 根据内容生成字符串
        ''' </summary>
        ''' <param name="commentPrefix">表示行注释的前缀字符</param>
        ''' <returns>返回文本内容的字符串</returns>
        Public Function ToContentString(Optional commentPrefix As String = ";") As String
            If commentPrefix Is Nothing Then commentPrefix = ""
            Dim text As New StringBuilder
            Dim index = 1
            For Each rec In _ItemList
                If rec.IsKeyValue Then
                    text.Append(String.Format("{0}={1}", rec.Content, _Entries.Item(rec.Content)))
                ElseIf rec.Type = RecordType.Comment Then
                    text.Append(commentPrefix)
                    text.Append(rec.Content)
                Else
                    text.Append(rec.Content)
                End If
                If index <> _ItemList.Count Then
                    text.AppendLine()
                    index += 1
                End If
            Next
            Return text.ToString()
        End Function

        Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, String)) _
            Implements IEnumerable(Of KeyValuePair(Of String, String)).GetEnumerator
            Return DirectCast(_Entries, IEnumerable(Of KeyValuePair(Of String, String))).GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function
    End Class
End Namespace