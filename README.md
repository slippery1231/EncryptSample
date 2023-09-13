## 設計
利用excel模擬DB，作為存放key和value的地方。透過讀取Excel內的資料做比對來達到對不同的pdf做到不同組密碼的加密。
應用:薪資單

## 使用步驟
步驟一 : 打開壓縮檔裡面的 NameAndPassword.xlsx 的 excel 檔案，這個excel 的目的是為了要知道哪個人要對應到哪個密碼，這邊的設計身分證字號會是被加密的密碼。
E.g.林小美這個人最後被加密完要打開檔案必須輸入 B123456 才打得開檔案。如果之後有新進的人要記得再回來加這個 excel
 
步驟二:打開上圖紅框 appsettings 的檔案
 
步驟三 : 去更改裡面的設定值

- ExcelFilePath:excel 的路徑，是用來比對姓名和要加密的密碼
- PdfFilePath: 要被加密的 pdf 來源們，
- OutputFilePath : 加密完檔案要被放置的位置，意思為加密執行完檔案就會在這裡出現。(注意:這裡的檔案每次執行完都會被覆蓋，例如 9/1 和 9/2 各執行一次，那 9/1 的檔案就會被覆蓋，除非你有去更改輸出的路徑)
路徑的設定更改完，勇敢地按下存檔

步驟四: EncryptSample.exe 這個執行檔案
 
 
步驟五: 按下步驟四的檔案會跳出Console視窗，如果出現"按任意鍵關閉"代表已加密完成
