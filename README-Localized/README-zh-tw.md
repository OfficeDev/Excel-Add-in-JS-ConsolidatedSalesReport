# Excel 2016 的綜合銷售報表工作窗格增益集範例

_適用版本：Excel 2016_

這個工作窗格增益集示範如何使用 Excel 2016 中的 JavaScript API，合併多個工作表中的資料。共有兩種型態︰程式碼編輯器和 Visual Studio。

![綜合銷售報表範例](../images/ConsolidatedSalesReport_report.PNG)

## 進行測試
### 程式碼編輯器版本

部署及測試增益集的最簡單方式，是將檔案複製到網路共用。

1.  在網路共用中建立資料夾 (例如，\\\MyShare\ConsolidatedSalesReport)，然後複製 [程式碼編輯器] 資料夾中的所有檔案。 
2.  編輯資訊清單檔案的 <SourceLocation> 元素，讓它指向步驟 1 的共用位置。 
3.  將資訊清單 (ConsolidatedSalesReportManifest.xml) 複製到網路共用 (例如 \\\MyShare\MyManifests)。
4.  在 Excel 中，將包含資訊清單的共用位置新增為受信任的應用程式目錄。

    a.啟動 Excel，並開啟空白的試算表。  
    
    b.選擇 **[檔案]** 索引標籤，然後選擇 **[選項]**。
    
    c.選擇 **[信任中心]**，然後選擇 **[信任中心設定]** 按鈕。
    
    d.選擇 **[受信任的增益集目錄]**。
    
    e.在 **[目錄 URL]** 方塊中，輸入您在步驟 3 建立的網路共用路徑，然後選擇 **[新增目錄]**。
    
   f. 選取 [在功能表中顯示]<e /> 核取方塊，然後選擇 [確定]<e />。接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。 
        
5.  測試並執行增益集。 

    a.在 Excel 2016 的 **[插入]** 索引標籤上，選擇 **[我的增益集]**。
    
    b.在 **[Office 增益集]** 對話方塊中，選擇 **[共用資料夾]**。
    
    c.選擇 **[綜合銷售報表範例]** > **[插入]** 增益集會在目前的工作表右側的工作窗格中開啟，如下圖所示。 
        
   ![綜合銷售報表範例](../images/ConsolidatedSalesReport_taskpane.PNG)

    d.選取 [美洲、亞洲及歐洲 ] 核取方塊，並按一下 **[合併!]** 按鈕。  這樣會建立新的儀表板工作表，其會顯示所有選定工作表的摘要檢視。
        
  ![綜合銷售報表範例](../images/ConsolidatedSalesReport_report.PNG)

### Visual Studio 版本
1.  將專案複製到本機資料夾，並在 Visual Studio 中開啟 Excel-Add-in-JS-ConsolidatedSalesReport.sln。
2.  按 F5 建置及部署範例增益集。Excel 會啟動，且增益集會在工作表右側空白部分的工作窗格中開啟，如下圖所示。 
        
   ![綜合銷售報表範例](../images/ConsolidatedSalesReport_taskpane.PNG)

    d.選取 [美洲、亞洲及歐洲 ] 核取方塊，並按一下 **[合併!]** 按鈕。  這樣會建立新的儀表板工作表，其會顯示所有選定工作表的摘要檢視。
        
  ![綜合銷售報表範例](../images/ConsolidatedSalesReport_report.PNG)


### 深入了解

1.  [Excel 增益集程式設計概觀](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Excel 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel 增益集程式碼範例](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Excel 增益集 JavaScript API 參考](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [建立第一個 Excel 增益集](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
