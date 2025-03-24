const XLSX = require("xlsx");
const { Builder, By } = require("selenium-webdriver");
const chrome = require("selenium-webdriver/chrome");
const { esperarEClicar, limparEEscrever, verificarRetencao } = require("./utils");
const fs = require('fs');
require("chromedriver");


const workbook = XLSX.readFile("scamwkb0.csv");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
let dados = XLSX.utils.sheet_to_json(sheet);

dados = dados.filter(nf =>
    nf['Espec.Docum.'] === "NFS" &&
    ['andrey.cardoso', 'carlos.passos', 'ana.anjos', 'emily.gawski'].includes(nf['Nome Us.Incl'])
);

console.log(dados);


(async () => {
  
    let driver = await new Builder()
        .forBrowser("chrome")
        .setChromeOptions(new chrome.Options())
        .build();

    await driver.get("https://portal.eqsengenharia.com.br/login");


    await driver.findElement(By.xpath(
        "/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[1]/div[1]/po-login/po-field-container/div/div[2]/input"
        )).sendKeys("bot.contabil");
    
    await driver.findElement(By.xpath(
        "/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[2]/div[1]/po-password/po-field-container/div/div[2]/input"
        )).sendKeys("EQSeng852@");
    
    await driver.findElement(By.xpath(
        "/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/po-button/button"
        )).click();


    await esperarEClicar(driver, "/html/body/app-root/app-main/div/po-menu/div/div[1]/po-icon/i");

    await esperarEClicar(driver, "/html/body/app-root/app-main/div/po-menu/div[2]/div[2]/div/div[2]/div/div/nav/ul/li[2]/po-menu-item/div/div[1]/po-icon[2]");

    await esperarEClicar(driver, "/html/body/app-root/app-main/div/po-menu/div[2]/div[2]/div/div[2]/div/div/nav/ul/li[2]/po-menu-item/div/div[2]/ul/li[1]/po-menu-item/a/div");
    
    
    driver.manage().window().maximize();


    // EU TENHO QUE PERCORRER A LISTA DADOS
    
    dataDeEmissao = dados[0]['DT Emissao'];


    await limparEEscrever(
        driver,
        dataDeEmissao, 
        "/html/body/app-root/app-main/div/app-processo-pagamento-lista/po-modal[1]/div/div[2]/div/div/div[2]/form/po-datepicker[1]/po-field-container/div/div[2]/div/input"
        );

    await limparEEscrever(
        driver,
        dataDeEmissao, 
        "/html/body/app-root/app-main/div/app-processo-pagamento-lista/po-modal[1]/div/div[2]/div/div/div[2]/form/po-datepicker[2]/po-field-container/div/div[2]/div/input"
        );


    // EU TENHO QUE PERCORRER A LISTA DADOS

    let fornecedor = dados[0].Fornecedor

    let nf = dados[0].Numero
    nf = String(nf).padStart(9, '0');

    await limparEEscrever(
        driver,
        fornecedor, 
        "/html/body/app-root/app-main/div/app-processo-pagamento-lista/po-modal[1]/div/div[2]/div/div/div[2]/form/po-input[2]/po-field-container/div/div[2]/input"
        );

    await limparEEscrever(
        driver,
        nf, 
        "/html/body/app-root/app-main/div/app-processo-pagamento-lista/po-modal[1]/div/div[2]/div/div/div[2]/form/po-input[3]/po-field-container/div/div[2]/input"
        );

    await driver.findElement(By.xpath(
        "/html/body/app-root/app-main/div/app-processo-pagamento-lista/po-modal[1]/div/div[2]/div/div/po-modal-footer/div/div/po-button/button"
        )).click();


    await esperarEClicar(driver, "/html/body/app-root/app-main/div/app-processo-pagamento-lista/po-page-list/po-page/div/po-page-content/div/po-table/div[2]/div/div/div/div/cdk-virtual-scroll-viewport/div[1]/table/tbody/tr/td[14]/div/po-table-column-icon/po-table-icon[1]/po-icon");


    await esperarEClicar(driver, "/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[1]/div/div/po-tab-button[4]/div[1]");

    await driver.findElement(By.xpath(
        "/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[4]/po-table/po-container/div/div/div/div/div/table/tbody/tr/td[4]/div/po-table-column-icon/po-table-icon[1]/po-icon"
        )).click();


    if (verificarRetencao(dados[0])) {
        try {
            fs.mkdirSync('./NFS/COM RETENÇÃO', { recursive: true });
          } catch (err) {
            console.error('Erro ao criar pasta:', err);
          }
    } else {
        try {
            fs.mkdirSync('./NFS/SEM RETENÇÃO', { recursive: true });
          } catch (err) {
            console.error('Erro ao criar pasta:', err);
          }
    }
    
    console.log(fs.readdirSync("./NFS"));


})();

