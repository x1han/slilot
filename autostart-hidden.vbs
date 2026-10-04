' AI PPT 助手 - 本地服务隐藏自启
' 放在用户"启动"文件夹中，登录 Windows 后自动以隐藏窗口运行 node server.js
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = "C:\Users\hxsci\ZCodeProject\ai-ppt-addin"
sh.Run "node server.js", 0, False
