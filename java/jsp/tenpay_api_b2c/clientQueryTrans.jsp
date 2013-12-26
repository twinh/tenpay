<%@ page language="java" contentType="text/html; charset=GBK" pageEncoding="GBK"%>

<%@ page import="com.tenpay.RequestHandler" %>
<%@ page import="com.tenpay.client.ClientResponseHandler" %>    
<%@ page import="com.tenpay.client.TenpayHttpClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
    //商户号 
    String partner = "1900000109";

    //密钥 
    String key = "8934e7d15453e97507ef794cf7b0519d";
    
    //创建查询请求对象
    RequestHandler reqHandler = new RequestHandler(null, null);
    //通信对象
    TenpayHttpClient httpClient = new TenpayHttpClient();
    //应答对象
    ClientResponseHandler resHandler = new ClientResponseHandler();
    
    //-----------------------------
    //设置请求参数
    //-----------------------------
    reqHandler.init();
    reqHandler.setKey(key);
    reqHandler.setGateUrl("https://gw.tenpay.com/gateway/normalorderquery.xml");
    
    //-----------------------------
    //设置接口参数
    //-----------------------------
    reqHandler.setParameter("partner", partner);    //商户号
    
    //out_trade_no和transaction_id至少一个必填，同时存在时transaction_id优先
    reqHandler.setParameter("out_trade_no", "1033537274");    	    	                //商家订单号
    //reqHandler.setParameter("transaction_id", "1900000109201101270026218385");	//财付通交易单号    

    //-----------------------------
    //设置通信参数
    //-----------------------------
    //设置请求返回的等待时间
    httpClient.setTimeOut(5);	
    
    //设置请求内容
    String requestUrl = reqHandler.getRequestURL();
    httpClient.setReqContent(requestUrl);
    String rescontent = "null";
    
    //后台调用
    if(httpClient.call()) {
    	//设置结果参数
    	rescontent = httpClient.getResContent();
    	resHandler.setContent(rescontent);
    	resHandler.setKey(key);
    	   	
    	//获取返回参数
    	String retcode = resHandler.getParameter("retcode");
    	
    	//判断签名及结果
    	if(resHandler.isTenpaySign()&& "0".equals(retcode)) {
    		out.println("订单查询成功</br>");
    		
    		//商户订单号
                String out_trade_no = resHandler.getParameter("out_trade_no");
                //财付通订单号
                String transaction_id = resHandler.getParameter("transaction_id");
                //金额,以分为单位
                String total_fee = resHandler.getParameter("total_fee");
                //如果有使用折扣券，discount有值，total_fee+discount=原请求的total_fee
                String discount = resHandler.getParameter("discount");
                //支付结果
                String trade_state = resHandler.getParameter("trade_state");
                //支付成功
                if("0".equals(trade_state)) {
                       //业务处理
                       out.println("transaction_id=" + resHandler.getParameter("transaction_id") +
                                  " out_trade_no=" + resHandler.getParameter("out_trade_no") + "</br>");
                       out.println("total_fee=" + resHandler.getParameter("total_fee") +
                                  " discount=" + resHandler.getParameter("discount"));
                                  
                }
    	} else {
    		//错误时，返回结果未签名，记录retcode、retmsg看失败详情。
    		System.out.println("验证签名失败或业务错误");
    		System.out.println("retcode:" + resHandler.getParameter("retcode")+
    	    	                    " retmsg:" + resHandler.getParameter("retmsg"));
    	}	
    } else {
    	System.out.println("后台调用通信失败");   	
    	System.out.println(httpClient.getResponseCode());
    	System.out.println(httpClient.getErrInfo());
    	//有可能因为网络原因，请求已经处理，但未收到应答。
    }
    
    //获取debug信息,建议把请求、应答内容、debug信息，通信返回码写入日志，方便定位问题
    System.out.println("http res:" + httpClient.getResponseCode() + "," + httpClient.getErrInfo());
    System.out.println("req url:" + requestUrl);
    System.out.println("req debug:" + reqHandler.getDebugInfo());
    System.out.println("res content:" + rescontent);
    System.out.println("res debug:" + resHandler.getDebugInfo());
    
%>

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk">
	<title>订单查询后台调用示例</title>
</head>
<body>
</body>
</html>
