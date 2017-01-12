# <a name="consolidated-sales-report-task-pane-add-in-sample-for-excel-2016"></a>Exemplo do Suplemento do Painel de Tarefas de Relatório de Vendas Consolidado para o Excel 2016

_Aplica-se a: Excel 2016_

Esse suplemento do painel de tarefas mostra como consolidar dados de várias planilhas usando as APIs JavaScript no Excel 2016. Há dois tipos: o editor de código e o Visual Studio.

![Exemplo de Relatório de Vendas Consolidado](../images/ConsolidatedSalesReport_report.PNG)

## <a name="try-it-out"></a>Experimente
### <a name="code-editor-version"></a>Versão do editor de código

A maneira mais fácil de implantar e testar o suplemento é copiar os arquivos para um compartilhamento de rede.

1.  Hospede os arquivos no Projeto do Editor de Códigos usando um servidor de sua escolha.
2.  Edite os elementos \<SourceLocation\> e \<URL\> do arquivo de manifesto (ConsolidatedSaleReportManifest.xml) para que ele aponte para o local hospedado da etapa 1. (por exemplo, https://localhost/consolidatedsalesreport/home.html)
3.  Copie o manifesto (ConsolidatedSalesReportManifest.xml) em um compartilhamento de rede (por exemplo, \\\MyShare\MyManifests).
4.  Adicione o local de compartilhamento que contém o manifesto como um catálogo de aplicativos confiáveis no Excel.

    a.  Inicie o Excel e abra uma planilha em branco.

    b.  Escolha a guia **Arquivo** e escolha **Opções**.

    c.  Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.

    d.  Escolha **Catálogos de Suplemento Confiáveis**.

    e.  Na caixa **URL de Catálogo**, insira o caminho para o compartilhamento de rede que você criou na etapa 3 e escolha **Adicionar Catálogo**.

   f.  Marque a caixa de seleção **Mostrar no Menu** e escolha **OK**. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office.

5.  Teste e execute o suplemento.

    a.  Na **guia Inserir** no Excel 2016, escolha **Meus Suplementos**.

    b.  Na caixa de diálogo **Suplementos do Office**, escolha **Pasta Compartilhada**.

    c.  Escolha **Exemplo de Relatório de Vendas Consolidado**>**Inserir**. O suplemento abre em um painel de tarefas à direita da planilha atual, conforme mostrado na figura a seguir.

   ![Exemplo de Relatório de Vendas Consolidado](../images/ConsolidatedSalesReport_taskpane.PNG)

    d.  Marque as caixas de seleção Américas, Ásia e Europa e clique no botão **Consolidar!**.  Isso criará a nova planilha Painel, que mostra um modo de exibição de resumo de todas as planilhas selecionadas.

  ![Exemplo de Relatório de Vendas Consolidado](../images/ConsolidatedSalesReport_report.PNG)

### <a name="visual-studio-version"></a>Versão do Visual Studio
1.  Copie o projeto para uma pasta local e abra o Excel-Add-in-JS-ConsolidatedSalesReport.sln no Visual Studio.
2.  Pressione F5 para criar e implantar o suplemento de exemplo. O Excel inicia e o suplemento abre em um painel de tarefas à direita da planilha em branco, conforme mostrado na figura a seguir.

   ![Exemplo de Relatório de Vendas Consolidado](../images/ConsolidatedSalesReport_taskpane.PNG)

    d.  Marque as caixas de seleção Américas, Ásia e Europa e clique no botão **Consolidar!**.  Isso criará a nova planilha Painel, que mostra um modo de exibição de resumo de todas as planilhas selecionadas.

  ![Exemplo de Relatório de Vendas Consolidado](../images/ConsolidatedSalesReport_report.PNG)


### <a name="learn-more"></a>Saiba mais

1.  [Visão geral da programação de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de Trechos para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemplos de código de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Referência da API JavaScript de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Criar seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)