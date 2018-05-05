''''''''''''''''''''''''''''''''''''''''''''''
'功能：锐捷交换机批量备份脚本
'适用：Xshell测试通过(测试环境s2952,s6000,s6010,s8610)
'作者：DCChao
'邮箱：dcchao@126.com
''''''''''''''''''''''''''''''''''''''''''''''

'*** 正则替换函数  ***'
'patern:正则表达式'
'str：原字符串'
'tagstr：替换字符'
'''''''''''''''''''''''
function replaceregex(patern,str,tagstr) 
	dim regex,matches 
	set regex=new regExp 
	regex.pattern=patern 
	regex.IgnoreCase=true 
	regex.global=true 
	matches=regex.replace(str,tagstr) 
	replaceregex=matches 
end function

Sub Main
	
	'*** 配置ssh用户名、密码、ip地址 ***
	Dim User, Pw, Ip, FirstAddress, EndAddress, BackupAddress, IpEnd
	User = "admin"
	Pw = "password"
	Ip = "192.168.1."             'IP段'
	FirstAddress = 1              '首ip'
	EndAddress = 38                '尾ip'
	
	'*** 配置tftp服务器地址 ***
	BackupAddress = "192.168.1.222"

	
	'*** 备份交换机配置 循环语句 ***
	For i = FirstAddress To EndAddress
	
		'*** 建立连接 *** 
		Dim SshUrl
		SshUrl = "ssh://" + User + ":" + Pw + "@" + Ip + cstr(i)
		xsh.Session.Open(SshUrl)
		xsh.Screen.Synchronous = true
		xsh.Session.Sleep(1000)
		
		
		'*** 获取交换机名称 ***
		Dim ScreenRow
		
		ScreenRow = xsh.Screen.CurrentRow
		ScreenRow = xsh.Screen.Get(ScreenRow,1,ScreenRow,30)
		
		'*** 进入配置模式 ***
		xsh.Screen.Send("en") & VbCr
		xsh.Screen.Send(Pw) & VbCr
		
		dim SwName
		SwName = replaceregex("\/",replaceregex("<|>",ScreenRow,""),"_")
		'*** 文件名称过长时开启 ***
		'SwName = replaceregex("HR-",SwName,"")
		
		'*** 下载交换机配置 *** 
		xsh.Screen.Send("copy flash:config.text tftp://" + BackupAddress + "/" + SwName + ".text")
		xsh.Screen.Send(VbCr)
		
		'*** 关闭本次连接 ***
		xsh.Screen.WaitForString("!")
		xsh.Session.Close()
		
		xsh.Session.Sleep(1000)
	Next
	'*** 弹出框 调试用 ***
	'xsh.Dialog.MsgBox(ScreenRow)
	
End Sub