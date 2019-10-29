# 在 Excel 中用公式及 VBA 将节点转化为链接

**NOTE：*本来此教程写在[有道云笔记分享](https://note.youdao.com/ynoteshare1/index.html?id=28611b2be7c1b63a28dba0f7cdb0aae7&type=note)的，似乎有段时间服务器开小差刷不出来， 故备份到 Github。***

---

在分步讲解之前，先预览一下全程。**整个过程如动图所示：**
![getss.gif](https://i.loli.net/2019/10/25/CSZiHQcayIU5eXq.gif)

下面把步骤分解。


## 1、复制节点信息，粘贴到工作簿

新建一个 Excel 工作簿，到<https://www.youneed.win/free-ss>等网站复制节点的信息，粘贴到工作簿。

![复制节点信息，粘贴到工作簿](https://i.loli.net/2019/10/25/mdjXsG35nNLxTuZ.png)


## 2、字段拼接

把`加密方式`、`密码`、`帐号`、`端口`拼接起来，格式为`加密方式:密码@帐号:端口`。所以单元格 `H2` 可以键入 `=D2&":"&C2&"@"&A2&":"&B2` 拼接第一个节点。其余节点可以通过向下拖动等方式填充。

![147.png](https://i.loli.net/2019/10/25/AbD8eUFP4hWtoKv.png)

> H 列的字段只是“中间产物”，我出于方便的目的，设置 H2 以下的单元格缩小字体填充。


## 3、用 VBA 自定义函数

首先，按 `Alt`+`F11`“召唤”出 VBE 窗口。

**其次，插入一个模块。**右击 `Sheet1 (sheet1)`，如下图所示插入一个模块。

![320.png](https://i.loli.net/2019/10/25/PijA5yNaeU7GDRg.png)

然后，把 VBA 代码复制粘贴到代码编辑区。

![343.png](https://i.loli.net/2019/10/25/zFCscmOnyGI8SfE.png)


代码如下：
```
Private Function getss(strText As String, Optional strCharset As String = "ASCII") As String
    Dim arrBytes

    With CreateObject("ADODB.Stream")
     .Type = 2
     .Open
     .Charset = strCharset
     .WriteText strText
     .Position = 0
     .Type = 1
     arrBytes = .Read
     .Close
    End With

    With CreateObject("Microsoft.XMLDOM").createElement("tmp")
     .DataType = "bin.base64"
     .nodeTypedValue = arrBytes
     getss = "ss://" & Replace(Replace(.text, vbCr, ""), vbLf, "")
    End With

End Function
```

最后，保存代码。按快捷键 `Ctrl`+`S` 或点击 VBE 窗口左上方的保存按钮保存，需要注意的是，文件名可以自定义，保存类型请选择 `Excel 启用宏的工作簿(*.xlsm)`。

![476.png](https://i.loli.net/2019/10/25/P7xuNfe6AGjH58r.png)


## 4、使用自定义的 getss 函数

在单元格 `I2` 可以键入 `=getss(H2)` 处理第一个节点。其余节点可以通过向下拖动等方式填充。

![513.png](https://i.loli.net/2019/10/25/CVtXoYZaOB9KJTE.png)
