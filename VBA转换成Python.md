# 自动办公：VBA转换成Python

使用Python控制MS Office最便捷的方法就是录制宏，然后抄。

抄代码也得讲基本法：

部分功能是没有VBA代码的，这个时候只能上网查找VBA的实现了。



属性 参数

函数可以没有括号

Visible会影响效率吗？有什么不同。

performance hung crash

会

为什么都只讲这么做会怎么样，而不讲不这样会如何？

Word.Application.ScreenUpdating = False

The **ScreenUpdating** property controls most display changes 
on the monitor while a procedure is running. When screen updating is turned off, 
toolbars remain visible and Word still allows the procedure to display or 
retrieve information using status bar prompts, input boxes, dialog boxes, and 
message boxes. You can increase the speed of some procedures by keeping screen 
updating turned off. You must set the **ScreenUpdating** property to **True** when the 
procedure finishes or when it stops after an error.

You can use this method after using the **ScreenUpdating** 
property to disable screen updates. **ScreenRefresh** turns on screen 
updating for just one instruction and then immediately turns it off. Subsequent 
instructions don't update the screen until screen updating is turned on again 
with the **ScreenUpdating** property.

http://word.tips.net/T001261_Turning_Off_Background_Repagination.html

http://superuser.com/questions/989720/large-document-is-sluggish-and-its-pictures-not-visible-in-web-layout

http://word.mvps.org/faqs/interdev/makeappinvisible.htm

Application.Interactive

Blocking user input will prevent the user from interfering with the macro as it moves or activates Microsoft Excel objects.

This property is useful if you're using DDE or OLE Automation to communicate with Microsoft Excel from another application.

If you set this property to **False**, don't forget to set it back to **True**. Microsoft Excel won't automatically set this property back to **True** when your macro stops running.

https://msdn.microsoft.com/en-us/library/office/ff841248.aspx

Application.DisplayAlerts Property (Excel)

The default value is **True**. Set this property to **False** to suppress prompts and alert messages while a macro is running; when a message requires a response, Microsoft Excel chooses the default response.

When using the **SaveAs** method for workbooks to overwrite an existing file, the Confirm Save As dialog box has a default of No, while the Yes response is selected by Excel when the **DisplayAlerts** property is set to **False**. The Yes response overwrites the existing file.
When using the **SaveAs** method for workbooks to save a workbook that contains a Visual Basic for Applications (VBA) project in the Excel 5.0/95 file format, the Microsoft Excel dialog box has a default of Yes, while the Cancel response is selected by Excel when the **DisplayAlerts** property is set to **False**. You cannot save a workbook that contains a VBA project using the Excel 5.0/95 file format.

https://msdn.microsoft.com/en-us/library/office/ff839782.aspx

http://stackoverflow.com/questions/30422236/excel-vba-application-screenupdating-vs-application-visible

Increasing Word Automation Performance for Large Amounts of Data by Using Open XML SDK

https://msdn.microsoft.com/en-us/library/office/ff191178(v=office.14).aspx

## 使用常量

```python
>>>w=win32com.client.gencache.EnsureDispatch('Word.Application')
>>> win32com.client.constants.wdWindowStateMinimize
2
>>>
```

产生缓存

生成COM type lib

清理缓存

删除gen_py目录，`win32com.__gen_path__`

两种方法

1. makepy.py预先生成（不会）
2. 使用`EnsureDispatch`生成缓存(首次耗时较长，之后可以直接使用`Dispatch`)



http://stackoverflow.com/questions/23290216/error-xlprimary-not-defined-in-python-win32com/23306850#23306850

http://stackoverflow.com/questions/21465810/accessing-enumaration-constants-in-excel-com-using-python-and-win32com/21536701#21536701

http://stackoverflow.com/questions/28264548/how-to-use-win32com-client-constants-with-ms-word

## 错误处理 Error handling

> \>>> from win32com.client import Dispatch
> \>>> import pythoncom
> \>>> def OpenExcelSheet(filename):
> �     try:
> �         xl = Dispatch("Excel.Application")
> �         xl.Workbooks.Open(filename)
> �     except pythoncom.com_error, (hr, msg, exc, arg):
> �         print "The Excel call failed with code %d: %s" % (hr, msg)
> �         if exc is None:
> �            print "There is no extended error information"
> �         else:
> �            wcode, source, text, helpFile, helpId, scode = exc
> �            print "The source of the error is", source
> �            print "The error message is", text
> �            print "More info can be found in %s (id=%d)" % (helpFile, helpId)
> �
> \>>>

```python
from win32com.client.gencache import EnsureDispatch
import pythoncom

try:
    excel = EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Workbooks.Open("error_path.xls")
except pythoncom.com_error as _error:
    if _error.excepinfo is None:
        excepinfo = [None]*6
    else:
        excepinfo = _error.excepinfo 
    hresult_readable = u"""
hresult状态码：{0}
hresult状态码信息：{1}"""
    
    detail_readable = u"""
错误码：  {0}
错误源：  {1}
错误信息：{2}
帮助文档: {3}
帮助ID：  {4}
错误码：  {5}
                """   
    args_readable = u"""参数错误，参数所在位置：{0}"""
    print hresult_readable.format(_error.hresult, _error.strerror.decode('gbk'))
    print detail_readable.format(*excepinfo)
    if _error.argerror > 0:
        print args_readable.format(_error.argerror)

finally:
    excel.Quit()
```

win32api.FormatMessage(1).decode("gbk")

获取错误码的详细信息 1：函数不正确。



P211页

https://books.google.co.jp/books?id=fzUCGtyg0MMC&lpg=PA207&ots=Y46Xt6B9Bx&dq=VARIANT%20python&pg=PA212#v=onepage&q&f=true

## 参数传递

VARIANT

通常情况下，直接使用Python对象就可以了。不行的时候看下面这个链接。

http://docs.activestate.com/activepython/2.7/pywin32/html/com/win32com/HTML/variant.html

## With

```vb
Sub 选择所有英文字母()
'
' 选择所有英文字母 宏
'
'
    CommandBars("Navigation").Visible = False
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^$"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub
```

## 循环

看懂就可以了。

## 变量定义

不需要

```python
import os

from win32com.client import Dispatch

def wait():
    import time
    time.sleep(2)
    return None


def set_clipboard(text):
    import win32clipboard as wincb
    import win32con

    wincb.OpenClipboard()
    wincb.EmptyClipboard()
    wincb.SetClipboardData(win32con.CF_UNICODETEXT, text)
    wincb.CloseClipboard()
    return None


def excel_find_all(excel_app, what):
    """
    找到全部what所在位置(Row, Column)，以列表的形式返回

    VBA code
    expression.Find(What, After, LookIn, LookAt,
                    SearchOrder, SearchDirection,
                     MatchCase, MatchByte, SearchFormat)
    expression .FindNext(After)
    """
    _result = []
    first_occur = excel_app.Cells.Find(what)
    if first_occur is not None:
        _result.append((first_occur.Row, first_occur.Column))
        next_one = first_occur
        while True:
            next_one = excel_app.Cells.FindNext(next_one)
            if (next_one.Row, next_one.Column) not in _result:
                _result.append((next_one.Row, next_one.Column))
            else:
                return _result
    else:
        return None

# real work
# 打开Word 与 Excel
excel = Dispatch('Excel.Application')
excel.Visible = True

report_template_path = ur"E:\东莞理工学院\IP通信\IP实验模板.doc"

report_folder = ur"E:\东莞理工学院\IP通信\201441302623 彭未康"

guide_book_path = ur"E:\东莞理工学院\IP通信\中兴实物实验v2.0.xlsx"
exp_guide_book = excel.Workbooks.Open(guide_book_path)

for i in range(15, exp_guide_book.WorkSheets.Count+1):
    exp_guide_book.WorkSheets(i).Activate()
    experiment_name = exp_guide_book.WorkSheets(i).Name

    word = Dispatch("Word.Application")
    word.Visible = True
    report_doc = word.Documents.Open(report_template_path)

    # 查找设置
    # 全局查找常量
    wdFindContinue = 1

    my_find = word.Selection.Find
    my_find.ClearFormatting()
    my_find.Wrap = wdFindContinue

    # 粘贴常量
    wdFormatPlainText = 22
    # 实验名
    report_doc.Activate()
    my_find.Text = "{{experiment_name}}"
    my_find.Execute()

    set_clipboard(experiment_name)
    word.Selection.Paste()
    wait()

    # 拓扑图
    exp_guide_book.WorkSheets(i).Activate()
    exp_guide_book.WorkSheets(i).Shapes.SelectAll()
    wait()
    excel.Selection.Copy()

    report_doc.Activate()
    my_find.Text = "{{topological_graph}}"
    wait()
    my_find.Execute()
    wait()
    word.Selection.PasteSpecial(0,0,0,0,4)
    wait()

    # 代码
    exp_guide_book.WorkSheets(i).Activate()

    special_words = excel_find_all(excel, u"命令内容")
    start_cell = excel.Cells(special_words[0][0], special_words[0][1])

    special_words = excel_find_all(excel, u"exit")
    end_cell = excel.Cells(special_words[-1][0], special_words[-1][1])

    excel.Range(start_cell, end_cell).Select()

    wait()
    excel.Selection.Copy()

    wait()
    report_doc.Activate()
    wait()
    my_find.Text = "{{code}}"
    wait()
    my_find.Execute()
    wait()
    word.Selection.PasteAndFormat(wdFormatPlainText)
    wait()

    # 另存为 退出
    report_doc.Activate()
    wait()
    filename = os.path.join(report_folder, experiment_name)
    word.Documents(1).SaveAs(filename)
    word.Documents(1).Close()
```
### 常见问题

## 启动应用



