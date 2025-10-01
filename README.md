## 用途  
用于爬取uom系统的合格证，具体来说就是爬取中国民用航空局-民用无人驾驶航空器运营合格证的内容。  
后续还需要通过OCR技术对图片内容进行识别，整理成一个新版的Excel，以供网友方便查看。

最终目的是为了查看无人机培训机构在各个省份各个市级的数量以及成立趋势。         
运营合格证的链接参考：https://uom.caac.gov.cn/#/uav-sczs-show/BZSQ9142403001   

## 功能说明       
1. 合格证图片（将两张拼接成了一张）
3. 一个Excel文件（uav_cert_status_correct_range.xlsx）
保存的内容有：合格证编号、完成URL、状态、处理结果、图片合并结果、耗时（秒）。  
其中“合格证编号”为URL最终的编号，比如说上面的编号为：“BZSQ9142403001”。作为唯一项来确定爬取的内容。

处理模式有3种：  
1. 顺序获取（按编号范围批量生成，如2401001-2401010
2. 逐个获取（手动输入错误编号补充，用英文退号分隔）
3. 自动获取（自动补充xlsx表内错误）


合格证“已撤销”案例：[https://uom.caac.gov.cn/#/uav-sczs-show/BZSQ9142404300](https://uom.caac.gov.cn/#/uav-sczs-show/BZSQ9142404300)   
<img width="887" height="498" alt="PixPin_2025-10-01_08-51-00" src="https://github.com/user-attachments/assets/1e5c45ba-4a3c-4580-a288-361c9369d75a" />

##  最终获取内容     

合格证获取效果图片如下：   
<img width="892" height="2524" alt="BZSQ9142403001_merged" src="https://github.com/user-attachments/assets/cf369097-8911-4eb8-a292-085435e79f6d" />
表格内容效果参考：   
 <img width="1298" height="448" alt="image" src="https://github.com/user-attachments/assets/63b27540-8a1e-459b-a5c8-ec497949d1fb" />
   
    
## 后记    
这只是把图片获取到了图片，后续还需要通过OCR记录将所需要的内容提取出来，才能查看全国各个省份获得运营合格证的数量。     
UOM上完全走过手续的1000多个机构，而这合格证还没弄完总共都5000张证了（含撤销的）。为什么说正常启用的？因为代码“已启用”、“已撤销”部分还存在一些问题，还需要优化一下。
