热心朋友为64位excel改了代码，有需要的朋友可以尝试：
https://code.google.com/p/excelvba-googlehost/issues/detail?id=4



-


11.28更新了一个vba，专门针对部分网络商屏蔽ping包的网络环境，详情可在 Downloads 页面点入，或者https://code.google.com/p/excelvba-googlehost/downloads/detail?name=ExcelVBA_GoogleHost_DownPage.rar

如果网络ping地址功能正常可以完全不用理会这个专用包。


-

-

谷歌大部分服务可以通过自定义host来绕开墙，但是由于墙内各地网络封锁的差异，别人可用的host放在自己网络也许就不能用了。为了达到公开代码而且墙内大多数PC机的软件环境都可方便使用的目标，产生了这个利用微软Excel自带的VBA环境来测试、优化并且自动生成文本的解决办法。这个宏运行时读取 GoogleIp.txt文本里面保存的全部谷歌ip，在本机上自动进行ping测试，得到对应ip的延时并且排序以找到对于本机来说最快的ip，再根据此ip自动生成一组host文本，和一段特定格式的ip字符串，可以用在GoAgent的proxy.ini里面。

-

使用环境：有 32位Excel 的 Windows系统，ipv4网络环境。

-

项目网址：https://code.google.com/p/excelvba-googlehost/

下载网址：https://code.google.com/p/excelvba-googlehost/downloads/list

-

-

文件说明

压缩包有ExcelVBA\_GoogleHost.bas，GoogleHostName.txt，GoogleIp.txt，一份现成host.txt，说明.txt共五个文本文件。

ExcelVBA\_GoogleHost.bas：VBA代码，可以先用notepad打开，代码已经尽量加有注释了。

GoogleHostName.txt：收集到的一批谷歌的域名列表，格式是一行一个或数个谷歌域名，一行有多个域名的话域名之间用空格分开。
这份列表已经包含很多谷歌服务，也可以按照自己的实际需要来增加或者删除。

GoogleIp.txt：收集到的一批谷歌服务器ip，格式是一行一个ipv4地址，可以自行增减。
可以在just-ping.com，www.just-dnslookup.com或者其他网站查询谷歌多个不同域名，又或者使用命令行“nslookup -vc 谷歌域名 国外dns”这样的命令得到更多谷歌服务器ip，查询结果粘帖入excel整理成一行一个ip后添加进这个文本文件里以供使用。

一份现成host.txt：一份已经整理好的host文本，可以直接使用。

说明.txt：本文件。。。:P

-

-

使用方法

把压缩包解压到需要的目录，ExcelVBA\_GoogleHost.bas等文件应该先用notepad打开检查一次（VBA代码已经尽量添加注释了）。
在同一个目录里面右键新建一个空白工作簿，用Excel打开（如果是打开Excel自动新建的，记得先存盘到这个目录里面），确认Excel里面没有同时打开其他工作簿。Excel菜单点击“工具”->“宏”->“visual basic 编辑器”（或者快捷键 Alt + F11）打开VB编辑器界面，编辑器界面的菜单点击“文件”->“导入文件”（或者快捷键 Ctrl + M），选择ExcelVBA\_GoogleHost.bas，打开。然后回到Excel表格主界面，菜单选择“工具”->“宏”->“宏”（或者快捷键 Alt + F8），弹出的“宏”对话框里面选蓝GetGoogleHost，点击按钮“执行”，宏就开始自动运行。等到“整理完成”提示以后，表格的 H2 单元格就是可以让 GoAgent 使用的ip串，从 H5 单元格 # Google Begin 开始到下面 # Google End为止的就是整理得到的 host 文本，可以拷贝到系统host里面使用。

2.0版本开始，整理结果文本用蓝色字体显示，方便查看。

-

-

已知问题

1：如果全部ip的ping结果都是Error，有可能是
a、Excel被防火墙或者杀软拦截不能联网，需要放一次例外；
b、个别比较变态的网络商会屏蔽ping数据包，这时先在cmd窗口ping一下墙内大站看看有没有问题。

2：64位Excel需要更改变量和dll声明（参见： http://msdn.microsoft.com/zh-cn/library/office/ee691831.aspx ），我没有64位excel做测试，欢迎有心人帮忙更改代码。

-

-

其他说明

1：Office系列的宏病毒其实也就是VBA代码，所以对于只使用公式的同学，Excel菜单“工具”->“宏”->“安全性”里面应该设置为最高的安全性，而且不要运行来历不明的代码，如要运行应该事先检查代码内容。

安全性设置成最高以后，打开以前存盘的带宏工作簿会自动禁止运行，如果想既不更改安全性设置又运行本代码的话可以新建一个空白工作簿重新引入ExcelVBA\_GoogleHost.bas。

其实VBA也有代码可以设置安全性的，所以如果系统杀软没有宏病毒检测功能的话应该偶尔手动检查宏安全性。

2：那份现成host公开以后估计不久会用不了，或者某个时候想快速换一组host，这时可以使用notepad的 Ctrl+H 替换功能。先找到一个通的ip，可以在cmd里面ping，或者把GoogleIp.txt里面的ip逐个粘帖到浏览器地址栏回车，能看到谷歌页面的就是通的。然后把可用的ip前3段替换掉原来host的前3段（比方说host里面现在的a.b.c.d系列不能用了，找到一个e.f.g.h是通的，就可以试一下把host里面a.b.c. 全部替换为 e.f.g.），GoAgent的hosts串也可以使用类似技巧。

3：各系统的host文件修改方法请搜索教程，windows以外的其他系统的host某程度上格式差不多，可以利用Excle自动改一下适应其他系统的格式。
原host备份后清理掉全部旧的谷歌相关，用新的替换。如果谨慎的话应该先到类似 http://www.ip-adress.com/ip_tracer/ 这样的网站查询一下新ip，确认新ip段属于谷歌公司。
GoAgent的local目录的proxy.ini里面，[google\_cn](google_cn.md)和[google\_hk](google_hk.md)下面都有 hosts = 字段，如果使用默认的设置总是不能连通，可以尝试把 H2 单元格生成的ip串替换进去，另外尝试mode = https。

4：www.google.com/ncr可以避免跳转香港谷歌（用页面的齿轮可设置成使用简体中文）。谷歌服务尽可能使用https，谷歌搜索可以尝试使用encrypted.google.com。
chrome和firefox浏览器可以使用HTTPS Everwyhere插件（官网： https://www.eff.org/https-everywhere ），chrome也可以使用chrome://net-internals/#hsts。

5：2.0版本开始在GoogleHostName.txt末尾增加了YouTube的部分域名。这些域名添加进入host以后，可以使用 https://www.youtube.com 来查看YouTube网页和缩略图，但是不能直接观看视频，而且这部分YouTube的host还会造成GoAgent代理观看YouTube视频也出现故障。但是带来一个好处：墙内部分地区可以使用这部分YouTube的host直接向 YouTube 上传视频。

所以2.0版本的VBA代码会自动在这部分YouTube域名host的前面加#号屏蔽它们，这样使用其他翻墙方法看YouTube就不会受到影响。如果有需要向YouTube上传视频，可以尝试手动把行首的#号删除保存以后，清空浏览器缓存并使用https方式访问YouTube，上传视频和修改YouTube设置。

6：谷歌服务器的ip也是经常变化的。使用一份新的host之前应该用类似 http://www.ip-adress.com/ip_tracer/ 这样的网站查询一下新的这组ip是不是属于谷歌公司，特别是从别处获得的不明确的ip。