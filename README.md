# Selenium-台灣證券交易所資料抓取(爬蟲)
透過Selenium抓取網頁中Script設定的內容，以台灣證券交易所為例。

1.設計簡易輸入，依提供之規則輸入即可。

![1662621154926](https://user-images.githubusercontent.com/103618758/189059147-f0d020e5-26c1-4425-a8aa-6aa26c1943b0.jpg)

![1662621286445](https://user-images.githubusercontent.com/103618758/189059567-e2ed6a4c-fcfd-4722-bf63-c4bc43a228e8.jpg)

2.程式會自動執行Selenium自動化開啟「台灣證券交易所」的網頁操作後關閉。

※Selenium有版本差異，本程式所使用之module第50、51行在新版中有更動，但為執行美觀有參入warnings進行忽略!

![12354646](https://user-images.githubusercontent.com/103618758/189059429-3b3d4b21-70f6-4e87-b56a-47f1a9bad516.png)

3.網頁代碼抓取後，透過Beautiful Soup進行解析後，建立並寫入本地excel檔存取。

![1662621351363](https://user-images.githubusercontent.com/103618758/189059529-90199dc3-f9f6-415f-9bae-76b8b9622685.jpg)

※本程式為練習爬蟲使用，短時間內多次運行有極高機率遭阻擋，請勿泛用。
