const XLSX = require("xlsx");
const { Builder, By } = require("selenium-webdriver");
const chrome = require("selenium-webdriver/chrome");

const { esperarEClicar, limparEEscrever, verificarRetencao, esperarArquivoAparecer,
    direcionarArq, baixarPrimeiroArq, extrairTexto } = require("./utils");

const fs = require('fs');
const chromedriver = require("chromedriver");


const workbook = XLSX.readFile("scamwkb0.csv");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
let dados = XLSX.utils.sheet_to_json(sheet);

dados = dados.filter(nf =>
    ["NFS", "NFA"].includes(nf['Espec.Docum.']) &&
    ['andrey.cardoso', 'carlos.passos', 'ana.anjos', 'emily.gawski'].includes(nf['Nome Us.Incl'])
);


const extensoes = ['.pdf', '.png', '.jpg', '.jpeg'];


(async () => {

    const options = new chrome.Options();
    const service = new chrome.ServiceBuilder(chromedriver.path).build();
    options.addArguments(`user-data-dir=C:\\Users\\Usuário\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 4`);
   
    let driver = await new Builder()
        .forBrowser("chrome")
        .setChromeOptions(options, service)
        .build();

    await driver.get("https://portal.eqsengenharia.com.br/login");


    await driver.findElement(By.xpath(
        "/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[1]/div[1]/po-login/po-field-container/div/div[2]/input"
        )).sendKeys("**********");
    
    await driver.findElement(By.xpath(
        "/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[2]/div[1]/po-password/po-field-container/div/div[2]/input"
        )).sendKeys("*********");
    
    await driver.findElement(By.xpath(
        "/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/po-button/button"
        )).click();
    

    await driver.sleep(3000);
    
    //await esperarEClicar(driver, "/html/body/app-root/app-main/div/po-menu/div/div[1]/po-icon/i");

    await esperarEClicar(driver, "/html/body/app-root/app-main/div/po-menu/div/div[2]/div[1]/div[2]/div/div/nav/ul/li[2]/po-menu-item/div");

    await esperarEClicar(driver, "/html/body/app-root/app-main/div/po-menu/div/div[2]/div[1]/div[2]/div/div/nav/ul/li[2]/po-menu-item/div/div[2]/ul/li[1]/po-menu-item/a");
    

    for (const notaFiscal of dados){
    
        dataDeEmissao = notaFiscal['DT Emissao'];

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


        let fornecedor = notaFiscal.Fornecedor

        let nf = notaFiscal.Numero
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


        let temRetencao = verificarRetencao(notaFiscal)

        let nomeDoForn = notaFiscal['Nome Fornec.']

        let nomeDoArquivo = `NF ${nf} ${nomeDoForn}`


        if ( temRetencao ) {
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

        let aux = 1;
        await (async function cacarNFS() {

            try {

                let download = await driver.findElement(By.xpath(
                    "/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[4]/po-table/po-container/div/div/div/div/div/table/tbody/tr/td[4]/div/po-table-column-icon/po-table-icon[1]"
                    ));

                await driver.executeScript("arguments[0].click();", download);

                await new Promise(resolve => setTimeout(resolve, 600));

                let arquivo = await esperarArquivoAparecer();
                console.log(arquivo);

                await new Promise(resolve => setTimeout(resolve, 600));

                direcionarArq(temRetencao, arquivo, nomeDoArquivo);


            } catch (error) {
                console.error(error);

                try {

                    let tipoDeArq = await driver.findElement(By.xpath(
                        `/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[4]/po-table/po-container/div/div/div/div/div/table/tbody[${aux}]/tr/td[1]/div/span/po-table-column-label/po-tag/div/div/div/div/div/span`
                        )).getText();

                    if (tipoDeArq == 'Boleto') {

                        aux+=1;
                        return cacarNFS()


                    } else {

                        download = await driver.findElement(By.xpath(
                            `/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[4]/po-table/po-container/div/div/div/div/div/table/tbody[${aux}]/tr/td[4]/div/po-table-column-icon/po-table-icon[1]`
                            ));

                        await driver.executeScript("arguments[0].click();", download);
        
                        await new Promise(resolve => setTimeout(resolve, 600));
        
                        let arquivo = await esperarArquivoAparecer();

            
                        if (arquivo.toLowerCase().includes("nf")) {
                            aux+=1
                            if (!arquivo.toLowerCase().includes("bol"))
                                direcionarArq(temRetencao, arquivo, nomeDoArquivo);
                            else
                                return cacarNFS();

                        } else if (arquivo.toLowerCase().includes("bol")) {
                            console.log("carro do amongus");
                            aux+=1
                            try {
                                tipoDeArq = await driver.findElement(By.xpath(
                                    `/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[4]/po-table/po-container/div/div/div/div/div/table/tbody[${aux}]/tr/td[1]/div/span/po-table-column-label/po-tag/div/div/div/div/div/span`
                                    )).getText();

                                    return cacarNFS()
                                
                            } catch (error) {
                                console.error(error);
                                for (const extensao of extensoes){
                                    try {
                                        fs.unlink(`./NFS/${arquivo}${extensao}`);
                                        break;
                                    } catch (error) {
                                        console.error(error)};
                                }       
                            }

                            await baixarPrimeiroArq(temRetencao, nomeDoArquivo);

                        } else {
                            console.log("carro da fiona");
                            try{
                                let texto = await extrairTexto(`NFS/${arquivo}.pdf`);
                                aux+=1
                                if (!texto.toLowerCase().includes("beneficiário") && texto.toLowerCase().includes("nfs")) {
                                    direcionarArq(temRetencao, arquivo, nomeDoArquivo);
                                } else {
                                    for (const extensao of extensoes){
                                        try {
                                            fs.unlink(`./NFS/${arquivo}${extensao}`);
                                            break;
                                        } catch (error) {
                                            console.error(error)};
                                    }    
                                    return cacarNFS()};

                            } catch (error) {
                                console.error(error);

                                await baixarPrimeiroArq(driver, temRetencao, nomeDoArquivo);
                            
                            }
                        }
                    }
                } catch (error) {
                    console.error(error)
                }
            }
        })();

        await esperarEClicar(driver, "/html/body/app-root/app-main/div/po-menu/div/div[2]/div[1]/div[2]/div/div/nav/ul/li[2]/po-menu-item/div/div[2]/ul/li[1]/po-menu-item/a");
    }

})();

