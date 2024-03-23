function onEdit(e) {
  // Get the active spreadsheet and the sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = e.range;
  //var evalue = e.value;
  let cell = null;
  if (range.getRow() == 4 && range.getColumn() == 26 && sheet.getSheetName() == "Portifolio") {
    cell = sheet.getRange("Z4");
    cell.setValue(false);
    abrirHtmlExternal();// da erro mas funciona?
    //SpreadsheetApp.getUi().alert("Cadastro concluido");

    //abrirCRUDInvestment();

  }
  //Extract FII
  if (range.getRow() == 1 && range.getColumn() == 14 && sheet.getSheetName() == "Stocks Filter") {
    cell = sheet.getRange("N1");
    cell.setValue(false);
    mineStock(papel);
  }
}
function createStock(ticker, preco, qtd, data) {

  Logger.log("Ticker: " + ticker + " Preco: " + preco + " Quantidade: " + qtd + " Data: " + data);

}

function fecharModal() {
  // Fecha o modal
  google.script.host.close();
}
function abrirHtmlExternal() {
  Logger.log("Html modal");
  var htmlOutput = HtmlService.createHtmlOutputFromFile('CStocks')
    .setWidth(300).setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');

}
//'https://bvmf.bmfbovespa.com.br/sig/FormConsultaNegociacoes.asp?strTipoResumo=RES_NEGOCIACOES&strSocEmissora=CIEL&strDtReferencia=02/2024&strIdioma=P&intCodNivel=1&intCodCtrl=100'; possiveis sources de test

function mineStock(papel) {
  //abrirHtmlExternal();
  var url = 'https://gmc-eduardo.github.io/FundamentusStock/';

  var options = {
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    },
    muteHttpExceptions: true
  };
  try {
    var resposta = UrlFetchApp.fetch(url, options);
    var html = resposta.getContentText();
    Logger.log("HTML fetch: ");
    //Logger.log(html);
    // Extrai conteúdo das células da tabela
    var conteudoTabela = extrairConteudoTabela(html);
    //Logger.log('Conteúdo da Tabela: ' + JSON.stringify(conteudoTabela));
  } catch (erro) {
    // Em caso de erro, exibir uma mensagem de erro
    Logger.log('Erro na requisição: ' + erro);
    SpreadsheetApp.getUi().alert('Erro na requisição: ' + erro);
  }
}

function getContent_(url) {
  return UrlFetchApp.fetch(url).getContentText()
}

//Extraindo pela fundamentos - nova solução usando cheerios
function extrairConteudoTabela(html) {
  var content = getContent_('https://gmc-eduardo.github.io/FundamentusStock/');

  // Carrega o HTML usando Cheerio
  var $ = Cheerio.load(html);
  Logger.log($('tr').text());
  var novaLinha ;
}


function exibirTabelas() {

}




