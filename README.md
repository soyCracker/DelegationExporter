# DelegationExporter
platform： .net core windows  
office lib： NPOI https://github.com/dotnetcore/NPOI   
pdf lib： itext7 https://github.com/itext/itext7-dotnet    
   
### 使用教學 2019/7/14:   
##### 松山:
1. 將欲處理的委派表excel檔案(xls或xlsx)放入file  
2. 用任何文字編輯器打開config.json，並將TargetXlsx欄位替換為你的委派表名稱並存檔，  
例如: "TargetXlsx": "2019年8月 傳道與生活聚會 學生委派表.xls"，注意雙引號  
3. 關閉excel後，執行DelegationExporter.exe，pdf會輸出至Export資料夾

##### 其他:
1.修改你的委派單excel與File資料夾裡的S89.xlsx檔名、位置、格式一樣
2.自行將S-89-CH.pdf放進File資料夾裡
3.將需要輸出的委派在export欄位打V
