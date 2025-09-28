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

合格证图片如下：   
<img width="892" height="2524" alt="BZSQ9142403001_merged" src="https://github.com/user-attachments/assets/cf369097-8911-4eb8-a292-085435e79f6d" />
