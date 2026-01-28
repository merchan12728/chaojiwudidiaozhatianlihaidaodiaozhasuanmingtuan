Set WshShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")

' 1. 先启动 Chrome 浏览器
objShell.ShellExecute "chrome.exe"
WScript.Sleep 20000 ' 等待 Chrome 启动

' 2. 打开 Chrome 后先按一次右键再按 Enter (选择用户头像)
WshShell.SendKeys "{RIGHT}"
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"
WScript.Sleep 2000

' 3. 打开登录页面并执行登录
objShell.ShellExecute "https://sites.google.com/a/kuencheng.edu.my/kchs/daily-notice"
WScript.Sleep 4000

WshShell.SendKeys "24955@kuencheng.edu.my"
WScript.Sleep 300
WshShell.SendKeys "{ENTER}"
WScript.Sleep 10000 ' 等待密码框加载

WshShell.SendKeys "kjt12728" ' 此处填入你的密码
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"
WScript.Sleep 5000 ' 等待登录跳转完成

' 4. 打开目标 Google Sheet 页面
objShell.ShellExecute "https://docs.google.com/spreadsheets/d/1PagVUShcQOfVLM0S3FzDrpYPisi89cdQaMtBUFk8kFk/edit?pli=1&gid=2014458434#gid=2014458434"
WScript.Sleep 8000 ' 重要：表格很大，需要较长时间加载才能响应快捷键

' 5. 换 Sheet 1 到 Sheet 4 (按 Alt + 向下键 3次)
' 注意：Google Sheets 官方切换表的快捷键是 Alt + Down
For i = 1 to 3
    WshShell.SendKeys "%{DOWN}" ' % 代表 Alt 键
    WScript.Sleep 800           ' 每次切换留一点反应时间
Next

' 6. 执行 Ctrl + F 并输入 c601
WshShell.SendKeys "^f"       ' ^ 代表 Ctrl 键
WScript.Sleep 1000           ' 等待查找框弹出
WshShell.SendKeys "c601"
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"