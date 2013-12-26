<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="./classes/PayResponseHandler.asp"-->
<!--#include file="./classes/NotifyResponseHandler.asp"-->
<%
'---------------------------------------------------------
'财付通即时到帐处理回调示例，商户按照此示例进行开发即可
'---------------------------------------------------------

'密钥
Dim key
key = "8934e7d15453e97507ef794cf7b0519d"

'创建支付应答对象
Dim resHandler
Set resHandler = new PayResponseHandler
resHandler.setKey(key)

'判断签名
If resHandler.isTenpaySign() = True Then
	
	Dim transaction_id
	Dim total_fee
	Dim out_trade_no
	Dim discount
	Dim trade_state
	
	'商户交易单号
	out_trade_no = resHandler.getParameter("out_trade_no")	

	'财付通交易单号
	transaction_id = resHandler.getParameter("transaction_id")
	
	'支付结果
	trade_state = resHandler.getParameter("trade_state")
	trade_mode = resHandler.getParameter("trade_mode")
	notify_id = resHandler.getParameter("notify_id")
	
	partner = resHandler.getParameter("partner")
	
	If "0" = trade_state and "1" = trade_mode Then
		'先使用notify_id去查询财付通服务器，确认支付成功
		Dim tenNotifyURL
		Dim sign
		Dim sign_str
		tenNotifyURL = "https://gw.tenpay.com/gateway/verifynotifyid.xml?"
		sign_str = "notify_id=" & notify_id & "&partner=" & partner & "&key=" & key
		sign = LCase(ASP_MD5(sign_str))
		
		tenNotifyURL = tenNotifyURL &"notify_id=" & notify_id & "&partner=" & partner & "&sign=" & sign
		'Response.Write("<br/>tenpay notify URL: " & tenNotifyURL & "<br/>")
		Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		
		Retrieval.setOption 2, 13056 
		Retrieval.open "GET", tenNotifyURL, False, "", "" 
		Retrieval.send()
		
		'文档已经解析完毕，客户端可以接受返回消息
		If Retrieval.Readystate =4 Then
			If 200 = Retrieval.Status Then
				ResponseTxt = Retrieval.ResponseText
				'新建服务器XMLDOM文档解析对象
				Set xmlDoc = server.CreateObject("Microsoft.XMLDOM")
				'加载请求返回的XML文档
				xmlDoc.loadxml(ResponseTxt)
				Set notifyResp = new NotifyResponseHandler
				notifyResp.setKey(key)
				'获取文档根元素
				Set obj =  xmlDoc.selectSingleNode("root")
				'遍历root的所有子节点，获取返回的键值对
				For Each node in obj.childnodes
					notifyResp.setParameter node.nodename, node.text 
					'Response.Write("<br/>" & node.nodename & "=" & node.text & "<br/>")
				Next
				'Response.Write("<br/>ResponseTxt: " & ResponseTxt & "<br/>")
				Set Retrieval = Nothing
				
				trade_state = notifyResp.getParameter("trade_state")
				trade_mode = notifyResp.getParameter("trade_mode")	
				
				'商品金额,以分为单位
				total_fee = notifyResp.getParameter("total_fee")
				
				'如果有使用折扣券，discount有值，total_fee+discount=原请求的total_fee
				discount = notifyResp.getParameter("discount")
	
				'判断notify响应的签名
				If notifyResp.isTenpaySign() and "0" = notifyResp.getParameter("retcode") and "0" = trade_state and "1" = trade_mode Then
					'Response.Write("<br/>查询通知签名验证成功<br/>")
					'------------------------------
					'确定订单已支付成功，处理业务开始
					'------------------------------
					
					'注意交易单不要重复处理
					
					'注意判断返回金额
					
					'------------------------------
					'处理业务完毕
					'------------------------------	
					
					'处理成功
					'返回给财付通服务器信息，重复通知的时候，直接返回success
					Response.Write("success")
				Else
					'非财付通通知或notify_id超时,参看recode的值和retmsg来确定原因, 当做不成功处理
					Response.Write("<br/>查询通知签名验证失败<br/>")
					notifyDebugInfo = notifyResp.getDebugInfo()
					Response.Write("<br/>retcode: " & notifyResp.getParameter("retcode") & "<br/>")
					Response.Write("<br/>retmsg: " & notifyResp.getParameter("retmsg") & "<br/>")
					Response.Write("<br/>Debug info: " & notifyDebugInfo & "<br/>")
				End If
			Else 
				'网络连接错误，Http返回码不是200，查询notify_id失败，记录订单号和日志，以后可以调用查询订单接口
				Response.Write("<br/>Http code: " & Retrieval.Status & "<br/>")
				Response.Write("<br/> 网络连接错误，订单号为：" + transaction_id)
			End if
		Else 
			'网络连接错误，查询notify_id失败，记录订单号和日志，以后可以调用查询订单接口
			Response.Write("<br/> 网络连接错误，订单号为：" + transaction_id)
		End if
	Else
		'当做不成功处理
		Response.Write("支付失败")
		Response.Write("<br/>trade_state:" & trade_state & "<br/>")
		Response.Write("<br/>pay_info:" & resHandler.getParameter("pay_info") & "<br/>")
		
	End If	

Else

	'签名失败
	Response.Write("签名签证失败")
	Dim debugInfo
	debugInfo = resHandler.getDebugInfo()
	Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")

End If
%>