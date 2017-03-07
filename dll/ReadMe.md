1. O2S.Components.PDFRender4NET.dll  
第三方DLL，可以实现PDF转图片，支持32位系统、64位系统  
官方试用版有红色水印，这个是没有水印的破解版，但还是希望大家支持正版。   
版本号4.5.1.2，目前在window7 64位下测试无问题  
2. Acrobat.dll  
Adobe官方接口，可以实现PDF转图片。该方法需要必须先安装Adobe Acrobat X Pro，再从安装目录下找到 Acrobat.dll 引用到项目中。 
Acrobat.dll 的转换效率要比其他第三方DLL 要快很多，而且更稳定，但是不支持多线程，所以在iis下会调用失败，有网友先用windows服务来调用Acrobat.dll，再用iis调用windows服务来解决该问题。  
如果您对转换速度、图片质量要求很高，该方法是最佳选择，但是实现过程最为麻烦。    
参考连接 http://blog.csdn.net/shi0090/article/details/7262199  
