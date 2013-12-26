<%@ page language="java" contentType="text/html; charset=GBK" pageEncoding="GBK"%>
<%@ page import="com.tenpay.util.TenpayUtil" %>
<%@ page import="com.tenpay.RequestHandler"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
//---------------------------------------------------------
//财付通网关支付请求示例，商户按照此文档进行开发即可
//---------------------------------------------------------
/* 商户号，上线时务必将测试商户号替换为正式商户号 */
String partner = "1900000109";
//密钥
String key = "8934e7d15453e97507ef794cf7b0519d";

//交易完成后跳转的URL
String return_url = "http://127.0.0.1:8180/tenpay_api_b2c/payReturnUrl.jsp";
//接收财付通通知的URL
String notify_url = "http://127.0.0.1:8180/tenpay_api_b2c/payNotifyUrl.jsp";

//---------------生成订单号 开始------------------------
//当前时间 yyyyMMddHHmmss
String currTime = TenpayUtil.getCurrTime();
//8位日期
String strTime = currTime.substring(8, currTime.length());
//四位随机数
String strRandom = TenpayUtil.buildRandom(4) + "";
//10位序列号,可以自行调整。
String strReq = strTime + strRandom;
//订单号，此处用时间加随机数生成，商户根据自己情况调整，只要保持全局唯一就行
String out_trade_no = strReq;
//---------------生成订单号 结束------------------------

//创建支付请求对象
RequestHandler reqHandler = new RequestHandler(request, response);
reqHandler.init();
//设置密钥
reqHandler.setKey(key);
reqHandler.setGateUrl("https://gw.tenpay.com/gateway/pay.htm");

//-----------------------------
//设置支付参数
//-----------------------------
reqHandler.setParameter("partner", partner);		        //商户号
reqHandler.setParameter("out_trade_no", out_trade_no);		//商家订单号
reqHandler.setParameter("total_fee", "1");			        //商品金额,以分为单位
reqHandler.setParameter("return_url", return_url);		    //交易完成后跳转的URL
reqHandler.setParameter("notify_url", notify_url);		    //接收财付通通知的URL
reqHandler.setParameter("body", "测试");	                    //商品描述
reqHandler.setParameter("bank_type", "DEFAULT");		    //银行类型
reqHandler.setParameter("spbill_create_ip",request.getRemoteAddr());   //用户的公网ip
reqHandler.setParameter("fee_type", "1");

//系统可选参数
reqHandler.setParameter("sign_type", "MD5");
reqHandler.setParameter("service_version", "1.0");
reqHandler.setParameter("input_charset", "GBK");
reqHandler.setParameter("sign_key_index", "1");

//业务可选参数
reqHandler.setParameter("attach", "");
reqHandler.setParameter("product_fee", "1");
reqHandler.setParameter("transport_fee", "0");
reqHandler.setParameter("time_start", currTime);
reqHandler.setParameter("time_expire", "");

reqHandler.setParameter("buyer_id", "");
reqHandler.setParameter("goods_tag", "");
//reqHandler.setParameter("agentid", "");
//reqHandler.setParameter("agent_type", "");

//请求的url
String requestUrl = reqHandler.getRequestURL();

//获取debug信息,建议把请求和debug信息写入日志，方便定位问题
String debuginfo = reqHandler.getDebugInfo();

System.out.println("requestUrl:" + requestUrl);
System.out.println("debuginfo:" + debuginfo);

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK">
<title>财付通即时到帐程序演示</title>
</head>
<body>
<br/><a target="_blank" href="<%=requestUrl %>">财付通支付</a>
</body>
</html>
