<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="./classes/PayRequestHandler.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk">
<title>财付通即时到帐支付请求示例</title>
</head>
<body>
<%
'---------------------------------------------------------
'财付通即时到帐支付请求示例，商户按照此示例进行开发即可
'---------------------------------------------------------

Dim strDate
Dim strTime
Dim randNumber
Dim key
Dim partner
Dim out_trade_no
Dim total_fee
Dim body
Dim return_url
Dim notify_url
Dim attach

'8位日期格式YYYYmmdd
strDate = getServerDate()

'6位时间,格式hhmiss
strTime = getTime()

'4位随机数
randNumber = getStrRandNumber(1000,9999)

'订单号，此处用时间加随机数生成，商户根据自己情况调整，只要保持全局唯一就行
out_trade_no = strDate & strTime & randNumber

'密钥，需要替换为商户自己的
key = "8934e7d15453e97507ef794cf7b0519d"

'商户号，需要替换为商户自己的
partner = "1900000109"

'回调地址，需要替换为商户自己的
return_url = "http://localhost/tenpay/return_url.asp"

'通知地址，需要替换为商户自己的
notify_url = "http://localhost/tenpay/notify_url.asp"


'商品价格，以分为单位
total_fee = "1"

'商品名称
body = "营业款上缴上海子公司"

'商户附加字段
attach = "付款人：张三"

'创建支付请求对象
Dim reqHandler
Set reqHandler = new PayRequestHandler
reqHandler.setGateUrl("https://gw.tenpay.com/gateway/pay.htm")
'初始化
reqHandler.init()

'设置密钥
reqHandler.setKey(key)

'-----------------------------
'设置支付参数
'-----------------------------
reqHandler.setParameter "partner", partner		'设置商户号
reqHandler.setParameter "out_trade_no", out_trade_no				'商户订单号
reqHandler.setParameter "body", body	'商品描述
reqHandler.setParameter "total_fee", total_fee				'商品总金额,以分为单位
reqHandler.setParameter "return_url", return_url			'回调地址
reqHandler.setParameter "notify_url", notify_url			'通知地址
reqHandler.setParameter "bank_type", "DEFAULT"						'银行类型
reqHandler.setParameter "fee_type", "1"						'银行类型
reqHandler.setParameter "spbill_create_ip", Request.ServerVariables("REMOTE_ADDR")  '支付机器IP

'系统可选参数
reqHandler.setParameter "sign_type", "MD5"
reqHandler.setParameter "service_version", "1.0"
reqHandler.setParameter "input_charset", "GBK"
reqHandler.setParameter "sign_key_index", "1"

'业务可选参数
reqHandler.setParameter "attach", attach

'请求的URL
Dim reqUrl
reqUrl = reqHandler.getRequestURL()

'debug信息
Dim debugInfo
debugInfo = reqHandler.getDebugInfo()
Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")
Response.Write("<br/>reqUrl" & reqUrl & "<br/>")


%>
<br/><a href="<%=reqUrl%>" target="_blank">财付通支付</a>
</body>
</html>