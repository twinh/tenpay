<?php
//---------------------------------------------------------
//财付通即时到帐支付请求示例，商户按照此文档进行开发即可
//---------------------------------------------------------

require_once ("classes/RequestHandler.class.php");

/* 商户号，上线时务必将测试商户号替换为正式商户号 */
$partner = "1900000109";

/* 密钥 */
$key = "8934e7d15453e97507ef794cf7b0519d";




//4位随机数
$randNum = rand(1000, 9999);

//订单号，此处用时间加随机数生成，商户根据自己情况调整，只要保持全局唯一就行
$out_trade_no = date("YmdHis") . $randNum;



/* 创建支付请求对象 */
$reqHandler = new RequestHandler();
$reqHandler->init();
$reqHandler->setKey($key);
$reqHandler->setGateUrl("https://gw.tenpay.com/gateway/pay.htm");

//----------------------------------------
//设置支付参数 
//----------------------------------------
$reqHandler->setParameter("total_fee", "2");  //总金额
//用户ip
$reqHandler->setParameter("spbill_create_ip", $_SERVER['REMOTE_ADDR']);//客户端IP
$reqHandler->setParameter("return_url", "http://localhost:8080/payReturnUrl.php");//支付成功后返回
$reqHandler->setParameter("partner", $partner);
$reqHandler->setParameter("out_trade_no", $out_trade_no);
$reqHandler->setParameter("notify_url", "http://localhost:8080/payNotifyUrl.php");
$reqHandler->setParameter("body", "测试");
$reqHandler->setParameter("bank_type", "DEFAULT");  	  //银行类型，默认为财付通
$reqHandler->setParameter("fee_type", "1");               //币种
//系统可选参数
$reqHandler->setParameter("sign_type", "MD5");  	 	  //签名方式，默认为MD5，可选RSA
$reqHandler->setParameter("service_version", "1.0"); 	  //接口版本号
$reqHandler->setParameter("input_charset", "GBK");   	  //字符集
$reqHandler->setParameter("sign_key_index", "1");    	  //密钥序号

//业务可选参数
$reqHandler->setParameter("attach", "");             	  //附件数据，原样返回就可以了
$reqHandler->setParameter("product_fee", "");        	  //商品费用
$reqHandler->setParameter("transport_fee", "");      	  //物流费用
$reqHandler->setParameter("time_start", date("YmdHis"));  //订单生成时间
$reqHandler->setParameter("time_expire", "");             //订单失效时间

$reqHandler->setParameter("buyer_id", "");                //买方财付通帐号
$reqHandler->setParameter("goods_tag", "");               //商品标记




//请求的URL
$reqUrl = $reqHandler->getRequestURL();

//获取debug信息,建议把请求和debug信息写入日志，方便定位问题
/*
$debugInfo = $reqHandler->getDebugInfo();
echo "<br/>" . $reqUrl . "<br/>";
echo "<br/>" . $debugInfo . "<br/>";
*/

?>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk">
	<title>财付通即时到帐程序演示</title>
</head>
<body>
<br/><a href="<?php echo $reqUrl ?>" target="_blank">财付通支付</a>
<form action="<?php echo $reqHandler->getGateUrl() ?>" method="post" target="_blank">
<?php
$params = $reqHandler->getAllParameters();
foreach($params as $k => $v) {
	echo "<input type=\"hidden\" name=\"{$k}\" value=\"{$v}\" />\n";
}
?>
<input type="submit" value="财付通支付">
</form>
</body>
</html>
