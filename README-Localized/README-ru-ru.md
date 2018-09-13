# <a name="consolidated-sales-report-task-pane-add-in-sample-for-excel-2016"></a>Пример надстройки области задач для Excel 2016, позволяющей получить сводный отчет о продажах

_Область применения: Excel 2016_

Этот пример надстройки области задач показывает, как объединять данные нескольких листов с помощью интерфейсов API JavaScript в Excel 2016. Надстройка представлена в двух вариантах — для редактора кода и для Visual Studio.

![Пример сводного отчета о продажах](../images/ConsolidatedSalesReport_report.PNG)

## <a name="try-it-out"></a>Проверка
### <a name="code-editor-version"></a>Версия для редактора кода

Самый простой способ развернуть и проверить надстройку — скопировать эти файлы в сетевую папку.

1.  Разместите файлы из папки Code Editor Proj (Проект редактора кода) на нужном сервере.
2.  Измените элементы \<SourceLocation\> и \<Url\> в файле манифеста (ConsolidatedSaleReportManifest.xml), чтобы он указывал на расположение, упоминавшееся на этапе 1 (например, https://localhost/consolidatedsalesreport/home.html).
3.  Скопируйте манифест (ConsolidatedSalesReportManifest.xml) в сетевую папку (например, \\\MyShare\MyManifests).
4.  Добавьте общую папку, содержащую этот манифест, в качестве доверенного каталога приложений в Excel.

    А.  Запустите Excel и откройте пустую электронную таблицу.

    Б.  Перейдите на вкладку **Файл**, а затем выберите **Параметры**.

    В.  Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.

    Г.  Выберите пункт **Доверенные каталоги надстроек**.

    Д.  В поле **URL-адрес каталога** введите путь к общему сетевому ресурсу, созданному на шаге 3, и выберите **Добавить каталог**.

   Установите флажок **Показывать в меню** и нажмите кнопку **ОК**. Появится сообщение о том, что параметры будут применены при следующем запуске Office.

5.  Проверьте и запустите надстройку.

    А.  На вкладке **Вставка** в Excel 2016 выберите элемент **Мои надстройки**.

    Б.  В диалоговом окне **Надстройки Office** выберите элемент **Общая папка**.

    В.  Выберите элементы **Consolidated Sales Report Sample** (Пример надстройки для создания сводного отчета о продажах) > **Вставить**. Надстройка откроется в области задач справа от текущего листа, как показано на рисунке ниже.

   ![Пример сводного отчета о продажах](../images/ConsolidatedSalesReport_taskpane.PNG)

    Г.  Установите флажки "America" (Америка), "Asia" (Азия) и "Europe" (Европа), а затем нажмите кнопку **Consolidate!** (Объединить).  Будет создан новый лист "Dashboard" (Панель мониторинга) со сводными данными всех выбранных листов.

  ![Пример сводного отчета о продажах](../images/ConsolidatedSalesReport_report.PNG)

### <a name="visual-studio-version"></a>Версия для Visual Studio
1.  Скопируйте проект в локальную папку и откройте файл Excel-Add-in-JS-ConsolidatedSalesReport.sln в Visual Studio.
2.  Нажмите клавишу F5, чтобы собрать и развернуть пример надстройки. Запустится Excel, а надстройка откроется в области задач справа от пустого листа, как показано на представленном ниже рисунке.

   ![Пример сводного отчета о продажах](../images/ConsolidatedSalesReport_taskpane.PNG)

    Г.  Установите флажки "America" (Америка), "Asia" (Азия) и "Europe" (Европа), а затем нажмите кнопку **Consolidate!** (Объединить).  Будет создан новый лист "Dashboard" (Панель мониторинга) со сводными данными всех выбранных листов.

  ![Пример сводного отчета о продажах](../images/ConsolidatedSalesReport_report.PNG)


### <a name="learn-more"></a>Подробнее

1.  [Общие сведения о программировании надстроек Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Обозреватель фрагментов кода для Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Примеры кода надстроек Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Справочник по API JavaScript для надстроек Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Создание первой надстройки Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)