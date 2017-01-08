# 使用Python获取操作系统的标准路径


## 关键词

桌面路径 临时文件夹 系统 下载文件夹 文档 用户目录 HOME 用户

## Windows系统

### 标准路径

用户文件夹下

- 图片
- 文档
- 下载
- 音乐
- 视频
- 桌面
- 链接
- 联系人
- 保存的游戏
- 收藏夹
- 搜索

### 方法

- 注册表
- Win32 API
- 主动询问用户

### 具体实现

#### 注册表

windows系统会在两个地方记录软件列表：

64位：HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall

32位：HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall

```python
#!/usr/bin/python
 
import os
import socket
import _winreg
 
#将软件安装列表输出到网盘上
os.system(r'net use p: \\10.0.0.6\public password /user:Lc\tanjun')
 
#使用主机名命名软件安装列表
hostname = socket.gethostname()
file = open(r'P:\todo\temp\%s.txt' % hostname, 'a')
 
#需要遍历的两个注册表
sub_key = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', r'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall']
 
software_name = []
 
for i in sub_key:
    key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, i, 0, _winreg.KEY_ALL_ACCESS)
    for j in range(0, _winreg.QueryInfoKey(key)[0]-1):
        try:
            key_name = _winreg.EnumKey(key, j)
            key_path = i + '\\' + key_name
            each_key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, key_path, 0, _winreg.KEY_ALL_ACCESS)
            DisplayName, REG_SZ = _winreg.QueryValueEx(each_key, 'DisplayName')
            DisplayName = DisplayName.encode('utf-8')
            software_name.append(DisplayName)
        except WindowsError:
            pass
 
#去重排序
software_name = list(set(software_name))
software_name = sorted(software_name)
 
for result in software_name:
    file.write(result + '\n')
```



```python
import _winreg

def get_desktop():
    key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,\
                          r'Software\Microsoft\Windows\CurrentVersion\Uninstall',)
    return _winreg.QueryValueEx(key, "Desktop")[0]
```

```
_winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
```

#### Win32 API

```python
import win32api,win32con
def get_desktop():
    key =win32api.RegOpenKey(win32con.HKEY_CURRENT_USER,\
                              r'Software\Microsoft\Windows\CurrentVersion\Explorer\ShellFolders',\
                             0,win32con.KEY_READ)
    return win32api.RegQueryValueEx(key,'Desktop')[0]
```

```python
from win32com.shell import shell, shellcon
def GetDesktopPath():
    ilist =shell.SHGetSpecialFolderLocation(0, shellcon.CSIDL_DESKTOP)
    return shell.SHGetPathFromIDList(ilist)
```



#### 主动询问用户

#### 获取用户文件夹路径

```Python
import os

os.path.expanduser("~")
> 'C:\\Users\\pengw'

import os
def GetDesktopPath():
    return os.path.join(os.path.expanduser("~"), 'Desktop')
```



### 参考

[博文：python获取当前系统的桌面的路径的四种方法](http://blog.csdn.net/hcj116/article/details/7747940)

[博文：python获取Windows特殊文件夹路径](http://blog.csdn.net/jackeriss/article/details/42153159)

```python
# -*- coding: cp936 -*-  
from win32com.shell import shell  
from win32com.shell import shellcon  
  
#获取"启动"文件夹路径，关键是最后的参数CSIDL_STARTUP，这些参数可以在微软的官方文档中找到  
startup_path = shell.SHGetPathFromIDList(shell.SHGetSpecialFolderLocation(0,shellcon.CSIDL_STARTUP))  
#获取"桌面"文件夹路径，将最后的参数换成CSIDL_DESKTOP即可  
desktop_path = shell.SHGetPathFromIDList(shell.SHGetSpecialFolderLocation(0,shellcon.CSIDL_DESKTOP))  
  
if __name__ == "__main__":  
    print startup_path  
    print desktop_path  
```



 [Python 操作注册表](http://codecloud.net/75169.html)

https://docs.python.org/2.7/library/_winreg.html

[python获取windows软件安装列表](http://11072687.blog.51cto.com/11062687/1752445)

http://stackoverflow.com/questions/6227590/finding-the-users-my-documents-path

```python
import ctypes.wintypes
CSIDL_PERSONAL = 5       # My Documents
SHGFP_TYPE_CURRENT = 0   # Get current, not default value

buf= ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)

print(buf.value)
```



[32位64位处理技巧](http://www.cnblogs.com/JeffreySun/archive/2010/01/04/1639117.html)



https://mail.python.org/pipermail/python-win32/2005-September/003738.html

[constant defined in shellcon. The full list](http://tech-artists.org/wiki/Python_Recipes)

[WindowsError 2](http://www.gossamer-threads.com/lists/python/python/788454) 

```
WindowsError: [Error 2] The system cannot find the file specified
```