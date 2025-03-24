const { By } = require("selenium-webdriver");


exports.esperarEClicar = async (driver, xpath) => {
    for(let i = 1; i > 0; i--){
            try{
                await driver.findElement(By.xpath(xpath)).click()
                break
            }catch (erro){
                await new Promise(resolve => setTimeout(resolve, 1000))
                i+=1
            }
        }
};


exports.limparEEscrever = async (driver, texto, xpath) => {
    const elementoWeb = await driver.findElement(By.xpath( xpath ));
        await elementoWeb.clear();
        await elementoWeb.sendKeys(texto);
};


exports.verificarRetencao = (dados) => {
    const temISS = dados['Rec Iss']
    
    const impostos = ['IRRF Ret', 'Valor INSS', 'Valor do PIS', 'Vlr.COFINS', 'Vlr.CSLL'];
    const temImposto = impostos.every(key => dados[key] != 0);

    if (temImposto == true || temISS == 'Sim') {
        return temImposto = true
    }
    return temImposto = false
}
