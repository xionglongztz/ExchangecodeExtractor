Imports System.Text.RegularExpressions '正则表达式
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadContent()
    End Sub
    Private Function LoadContent() '载入内容
        If Clipboard.ContainsText = False Then '如果剪贴板没有文本
            If MsgBox("剪贴板不含任何文本内容。" & vbCrLf & "请重新复制内容，并点击重试。",
               vbInformation + vbRetryCancel + vbSystemModal, "提取失败") = vbRetry Then
                LoadContent() '重新进入函数
            End If
        Else
            CodeExtract() '提取兑换码
        End If
        Dispose() '关闭程序
        Return 0
    End Function
    Private Function CodeExtract()
        Dim _pattern = "[a-zA-Z0-9]{12}" '设置表达式内容
        Dim _text = Clipboard.GetText '从剪贴板得到数据
        If Regex.IsMatch(Clipboard.GetText, _pattern) = True Then '如果命中规则
            Dim m As Match
            m = Regex.Match(_text, _pattern)
            For i = 1 To 3
                Clipboard.SetDataObject(m.Value) '设置剪贴板内容
                MsgBox(m.Value & vbCrLf & "当前兑换码已复制到剪贴板，粘贴完成后，点击确定可读取下一个",
                vbInformation + vbSystemModal, "兑换码" & i)
                m = m.NextMatch
            Next
            Clipboard.SetDataObject(_text)
            MsgBox("全部兑换码读取完毕。" & vbCrLf & "已将剪贴板内容设置为开启程序时的内容。",
            vbInformation + vbSystemModal, "提取成功")
        Else '未命中规则
            If MsgBox("文本内未检测到有效兑换码。" & vbCrLf & "正则表达式规则：[a-zA-Z0-9]{12}",
               vbInformation + vbRetryCancel + vbSystemModal, "提取失败") = vbRetry Then
                LoadContent()
            End If
        End If
        Return 0
    End Function
End Class
