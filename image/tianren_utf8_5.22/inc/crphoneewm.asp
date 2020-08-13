 <span style="position:relative;z-index:100">　<div id="qrcodeyl" style="position:absolute;top:30px;z-index:100;left:-90px;display:none;border:5px solid #fff;padding:0;margin:0;width:200px;height:200px;"></div></span><span id="phonepageyl" onMouseOver="showHide0123(qrcodeyl)" onMouseOut="showHide0123(qrcodeyl)">手机版</span> | <script>var qrcode = new QRCode('qrcodeyl', {
  text: 'http://<%=Request.ServerVariables("SERVER_NAME")%>',
  width: 200,
  height: 200,
  colorDark : '#000000',
  colorLight : '#ffffff',
  correctLevel : QRCode.CorrectLevel.H
});</script>