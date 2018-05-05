''''''''''''''''''''''''''''''''''''''''''''''
'功能：锐捷登陆脚本
'适用：Xshell测试通过
'作者：DCChao
'邮箱：dcchao@126.com
''''''''''''''''''''''''''''''''''''''''''''''

Sub Main

	
		xsh.Screen.WaitForString(">")
		xsh.Screen.Send("en") & VbCr
		xsh.Screen.Send("password") & VbCr
		xsh.Screen.Send("sho int des") & VbCr
	
	'*** 弹出框 调试用 ***
	'xsh.Dialog.MsgBox(ScreenRow)
	
End Sub