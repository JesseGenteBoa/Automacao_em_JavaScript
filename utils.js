const { until, By } = require("selenium-webdriver");
const fs = require('fs');


exports.esperarEClicar = async (driver, xpath) => {
    while (true){
        try {
            await driver.findElement(By.xpath(xpath)).click()
            break
        } catch {
            await new Promise(resolve => setTimeout(resolve, 1000))
        }
    }
};


exports.limparEEscrever = async (driver, texto, xpath) => {
    const elementoWeb = await driver.findElement(By.xpath( xpath ));
        await driver.sleep(300);
        await elementoWeb.clear();
        await driver.sleep(300);
        await elementoWeb.sendKeys(texto);
};


exports.verificarRetencao = (dados) => {
    const temISS = dados['Rec Iss']
    
    const impostos = ['IRRF Ret', 'Valor INSS', 'Valor do PIS', 'Vlr.COFINS', 'Vlr.CSLL'];
    let temImposto = impostos.every(key => dados[key] != 0);

    if (temImposto == true || temISS == 'Sim')
        return temImposto = true;
    
    return temImposto = false
};


exports.esperarArquivoAparecer = async () => {
    while (true) {
        const arquivos = fs.readdirSync("./NFS")
            .filter(arquivo => arquivo !== 'COM RETENÇÃO' && arquivo !== 'SEM RETENÇÃO');
                
        if (arquivos.length > 0)
            return arquivos[0].split('.')[0];

        await new Promise(resolve => setTimeout(resolve, 200));
    }
};


exports.direcionarArq = (temRetencao, arquivo, nomeDoArquivo) => {
    const extensoes = ['.pdf', '.png', '.jpg', '.jpeg'];

    if ( temRetencao == true ) {
        for (const extensao of extensoes) {
            try {
                fs.renameSync(`NFS/${arquivo}${extensao}`, `NFS/COM RETENÇÃO/${nomeDoArquivo}${extensao}`);
                break;
            } catch (error){console.error(error)}
        }
    } else {
        for (const extensao of extensoes) {
            try {
                fs.renameSync(`NFS/${arquivo}${extensao}`, `NFS/SEM RETENÇÃO/${nomeDoArquivo}${extensao}`);
                break;
            } catch (error){console.error(error)}
        }
    }
};


exports.baixarPrimeiroArq = async (driver, temRetencao, nomeDoArquivo) => {
    let botao = await driver.findElement(By.xpath(
        "/html/body/app-root/app-main/div/app-processo-pagamento-nota-manutencao/po-page-default/po-page/div/po-page-content/div/div[2]/po-tabs/div[2]/po-tab[4]/po-table/po-container/div/div/div/div/div/table/tbody/tr/td[4]/div/po-table-column-icon/po-table-icon[1]"
        ));

    await driver.executeScript("arguments[0].click();", botao);

    await new Promise(resolve => setTimeout(resolve, 600));

    let arquivo = await esperarArquivoAparecer();

    await new Promise(resolve => setTimeout(resolve, 600));

    direcionarArq(temRetencao, arquivo, nomeDoArquivo);
};


exports.extrairTexto = async(caminho) => {
    const data = await fs.readFile(caminho);
    const { text } = await pdf(data);
    return text
};

