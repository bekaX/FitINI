# FitINI
一个简单且方便的 INI 文件库

[TOC]

## 关于 INI 文件
`.ini`文件是一种简单易用的数据存储方式。所有的数据以键值对的方式记录，同时支持注释。

通常一个`.ini`文件中包含有若干个区块(Section)，每个区块中都保存了各自的键值对，通过键名(Key Name)可以访问到唯一的值(Value)。

**FitINI** 将常用的操作封装成了库，可以方便快速的读取、修改和写入`.ini`文件。并且在一定程度上支持非标准的 INI 风格的文件。

## 示例(Visual Basic.Net)
加载一个标准的 .ini 文件，修改后保存到另一个文件

    ' 加载文件
    Dim ini = FitINI.INI.LoadFromFile("config.ini")
    ' 读取内容
    Dim sec As FitINI.Section = ini.Section("SecName1")
    Debug.WriteLine(sec.Item("KeyName1"))
    Debug.WriteLine(ini.Entry("SecName2", "KeyName2"))
    ' 修改内容
    sec = ini.Add("SecName3")
    sec.Item("KeyName3") = "Hello"
    ' 保存到文件
    ini.SaveToFile("config2.ini")

## API使用说明(Visual Basic.Net)

### 类
**INI**：表示一个 INI 对象

**Section**：表示 INI 对象内部的一个区块

**LoadOption**：用于在解析 INI 文件数据时指定相关选项

区块的名字包含在了 INI 对象中，而不在 Section 对象中。
所有的类都不是线程安全的。如果要在多线程中共享，建议使用 *SyncLock* 关键字。

### INI 类

此类实现了 IEnumerable(Of KeyValuePair(Of String, Section)) 接口，可以使用 For Each 进行迭代。

#### 属性

**Section** 获取或设置此 INI 中的某一个区块

参数：name - 区块的名字

类型：*Section*

备注：返回名字为 name 的区块对象，若不存在则返回 Nothing，若将属性设置为 Nothing 会抛出 ArgumentNullException

**DefaultSection**(只读属性) 缺省状态的区块

类型：*Section*

备注：没有指定区块名的键值对和数据会被存在这里

**Entry** 获取或设置此 INI 中的条目

参数：name - 区块的名字，key - 键名

类型：*String*

备注：获取时若区块或键不存在返回 Nothing，写入时不存在的区块和键值对会被自动添加

**CommentPrefix** 获取或设置此 INI 的行注释前缀字符

类型：*String*

备注：设置为 Nothing 时会抛出 ArgumentNullException

#### 函数和过程
**Add** 在 INI 中用指定的名字新增一个区块(如果已存在则不会新建)

类型：*Section*

参数：name - 区块的名字

返回：新增的区块(名字为 name 的区块)

**GetEntry** 获取此 INI 中的条目的值

类型：*String*

参数：name - 区块的名字，key - 键名，def[可选] - 当区块或键不存在时的默认返回值

备注：作用与 Entry 属性相同，但是可以指定出错时的返回值

**Contains** 检测此 INI 中是否包含某一区块

类型：*Boolean*

参数：sectionName - 区块的名字

**Remove** 移除某一个区块

类型：*Boolean*

参数：sectionName - 区块的名字

备注：成功移除返回 True, 否则返回 False

**Rename** 修改某一个区块的名字

类型：*Boolean*

参数：oldName - 区块的名字，newName - 新名字

备注：成功修改返回 True, 否则返回 False

**Clone** 创建一个当前 INI 对象的副本(深拷贝)

类型：*INI*

**Clear**过程 清空 INI 对象中的所有数据

**ClearSections**过程 移除所有区块 (不含缺省区块)

**RemoveAllComments**过程 移除所有的注释和文本(非键值对数据)

**GetSectionNames** 获取所有的区块名

类型：*String* 数组

**ToList** 按行生成文本的列表

类型：*List(Of String)*

备注：返回包含每行文本内容的 String 列表

**ToContentString** 根据内容生成文本字符串

类型：*String*

**SaveToFile**过程 将数据写入到文件中保存

参数：path - 文件保存的路径

**LoadFromFile**共享函数 从文件加载并生成一个 INI 对象

类型：*INI*

参数：path - 文件路径

备注：当发生IO错误时，会返回一个不含任何区块的 INI 对象

**LoadFromFile**共享函数 从文件加载并生成一个 INI 对象，并指定加载所用的选项

类型：*INI*

参数：path - 文件路径，arg - 加载的选项

备注：加载的选项通过 *LoadOption* 类指定

**LoadFromFileSimply**共享函数 简化地从文件加载并生成一个 INI 对象，忽略任何非键值对的行

类型：*INI*

参数：path - 文件路径

**LoadFromList**共享函数 读取 String 列表生成一个 INI 对象

类型：*INI*

参数：list - 要解析的列表，arg[可选] - 加载的选项

备注：默认的 arg 参数中，UnknowLineOption 属性为 LineOption.ForceToComment

**LoadFromString**共享函数 从字符串生成一个 INI 对象

类型：*INI*

参数：list - 要解析的列表，arg[可选] - 加载的选项

备注：默认的 arg 参数中，UnknowLineOption 属性为 LineOption.AsMultiLine

**DirectLoadFromFile**共享函数 直接从文件读取键值对的值并返回

类型：*String*

参数：path - 文件路径，name - 区块名，key - 键名

备注：文件、区块或键不存在时返回 Nothinge

### Section 类

此类可以维持写入的键值对和注释的顺序。新添加的数据会加入到末尾，在输出类中的数据时默认会保持这个顺序。

此类实现了 IEnumerable(Of KeyValuePair(Of String, String)) 接口，可以使用 For Each 进行迭代。

#### 属性

**Item** 获取或设置此区块中键值对的值

参数：key - 键值对中的键名

类型：*String*

备注：若指定的 key 不存在则返回 Nothing

**Count**(只读属性) 获取此区块内的键值对的数量

类型：*Integer*

备注：若指定的 key 不存在则返回 Nothing

#### 函数和过程

**GetValue** 获取此区块中一个键值对的值

类型：*String*

参数：key - 键名，def[可选] - 当键不存在时的默认返回值

备注：作用与 Item 属性相同，但是可以指定出错时的返回值

**Contains** 检测此区块中是否包含某一键值对

类型：*Boolean*

参数：key - 键的名字

**Remove** 移除某一个键值对

类型：*Boolean*

参数：key - 键的名字

备注：成功移除返回 True, 否则返回 False

**Rename** 修改键值中对的键名

类型：*Boolean*

参数：oldName - 键名，newName - 新名字
修改返回 True, 否则返回 False

备注：成功
**Clone** 创建一个当前 Section 对象的副本(深拷贝)

类型：*Section*

**Clear**过程 清空当前 Section 对象的所有内容

**AddComment**过程 添加一条注释

参数：content - 注释的内容

**AddComment**过程 添加一条注释或纯文本

参数：content - 内容，isComment - 是普通的注释(True)还是非注释的纯文本(False)

备注：这个功能用于兼容一些非标准的 INI 文件

**GetComments** 获取所有的注释和文本 (除键值对以外的内容)

类型：*String* 数组

**RemoveAllComments**过程 移除所有的注释和文本 (除键值对以外的内容)

**GetKeys** 获取所有的键值对的键名

类型：*String* 数组

备注：如果要遍历所有的键值对，推荐使用 *For Each* 遍历当前区块，访问到的 KeyValuePair 就是区块中键值对的封装

**ToList** 按行生成文本的列表

类型：*List(Of String)*

参数：commentPrefix[可选] - 表示行注释的前缀字符

备注：返回包含每行文本内容的 String 列表，默认的行注释前缀为“;”

**ToContentString** 根据内容生成文本字符串

类型：*String*

参数：commentPrefix[可选] - 表示行注释的前缀字符

备注：默认的行注释前缀为“;”

### LoadOption 类

#### 属性

**IgnoreFileMissed** 若读取文件时打开失败，是否抛出异常

类型：*Boolean*

备注：指定为 False 时出错不抛出异常，默认为 True

**UnknowLineOption** 遇到无法解析的行的操作

类型：*LoadOption.LineOption* (Enum 枚举类型)

备注：LineOption有以下取值：
- Keep(默认值)：保留原文
- AsMultiLine：合并到上一行（有换行符）
- AsMultiLineCombined：合并到上一行（无换行符）
- ForceToComment：转换为注释
- Drop：丢弃此行


**DataCommentPrefix** 要解析的文件中注释的前缀

类型：*String*

备注：默认为“;”

---

感谢使用 ~ :)

目前由 MFS 的 **beka** 进行维护
