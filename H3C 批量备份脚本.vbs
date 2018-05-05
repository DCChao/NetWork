Sub Main
	
	''''''''''''''''''''''''''''''''''''''''''''''
	'功能：华三交换机批量备份脚本
	'适用：Xshell测试通过(测试环境)
	'作者：DCChao
	'邮箱：dcchao@126.com
	''''''''''''''''''''''''''''''''''''''''''''''
	
	'*** 配置ssh用户名、密码、ip地址 ***
	Dim User, Pw, Ip, FirstAddress, EndAddress, BackupAddress, IpEnd
	User = "admin"
	Pw = "password"
	Ip = "192.168.0."             'IP段'
	FirstAddress = 1              '首ip'
	EndAddress = 254              '尾ip'
	
	'*** 配置备份服务器地址 ***
	BackupAddress = "192.168.0.11"
	
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
		ScreenRow = xsh.Screen.Get(ScreenRow,1,ScreenRow,10)
		
		'*** 替换正则  ***
		dim regex, SwName
		set regex=new regExp
		regex.pattern="<|>"    '替换符号 < 或 >'
		regex.IgnoreCase=true 
		regex.global=true 
		SwName=regex.replace(ScreenRow,"")
		
		'*** 下载交换机配置 ***
		'xsh.Screen.Send("tftp 192.168.0.11 put flash:/startup.cfg 1.cfg")
		xsh.Screen.Send("tftp 192.168.0.11 put flash:/startup.cfg " + SwName + ".cfg")
		xsh.Screen.Send(VbCr)
		
		'*** 关闭本次连接 ***
		xsh.Screen.WaitForString(ScreenRow)
		xsh.Session.Close()
		
		xsh.Session.Sleep(1000)
	Next
	'*** 弹出框 调试用 ***
	'xsh.Dialog.MsgBox(ScreenRow)
	
End Sub