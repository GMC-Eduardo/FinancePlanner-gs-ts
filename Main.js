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
  //Extract Stocks
  if (range.getRow() == 1 && range.getColumn() == 14 && sheet.getSheetName() == "Stocks Filter") {
    cell = sheet.getRange("N1");
    cell.setValue(false);
    mineStock(sheet);
  }
  //Extract Stocks
  if (range.getRow() == 1 && range.getColumn() == 14 && sheet.getSheetName() == "FII Filter") {
    cell = sheet.getRange("N1");
    cell.setValue(false);
    mineFII(sheet);
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

function mineStock(sheet) {
  //abrirHtmlExternal();
  var url = 'https://gmc-eduardo.github.io/FundamentusStock/';
  var icell = [0, 1, 2, 3, 5, 16, 17, 18, 19, 20];
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
    var conteudoTabela = extrairConteudoTabela(html, sheet);
    //Logger.log('Conteúdo da Tabela: ' + JSON.stringify(conteudoTabela));
  } catch (erro) {
    // Em caso de erro, exibir uma mensagem de erro
    Logger.log('Erro na requisição: ' + erro);
    SpreadsheetApp.getUi().alert('Erro na requisição: ' + erro);
  }
}
function mineFII(sheet) {
  //abrirHtmlExternal();
  var url = 'https://gmc-eduardo.github.io/FundamentusFII/';
  
  var icell = [0, 1, 2, 4, 5, 6];
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
    var conteudoTabela = extrairConteudoTabela(html, sheet,icell);
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
function extrairConteudoTabela(html, sheet,icell) {
  // Carrega o HTML usando Cheerio
  var $ = Cheerio.load(html);

  // Encontra todas as tabelas no HTML
  $('table').each(function (t, tabela) {
    var rowIndex = 4; // Começa na linha 4 da planilha

    // Itera sobre as linhas da tabela
    $(tabela).find('tr').each(function (j, linha) {
      var colIndex = 1; // Começa na coluna A da planilha
      var linhaArray = [];

      // Itera sobre as células da linha
      $(linha).find('td').each(function (k, celula) {
        // Verifica se o índice da célula está na lista de índices desejados
        if (icell.includes(k)) {
          // Adiciona o conteúdo da célula à linhaArray
          linhaArray.push($(celula).text().trim()); // Usando trim() para remover espaços em branco extras
        }
      });

      // Verifica se a linha contém dados antes de inserir na planilha
      if (linhaArray.length > 0) {
        // Itera sobre os valores da linhaArray e insere na planilha
        linhaArray.forEach(function(valor) {
          sheet.getRange(rowIndex, colIndex++).setValue(valor);
        });
        rowIndex++; // Move para a próxima linha da planilha
      }
    });

    Logger.log('Tabela ' + (t + 1) + ' inserida na planilha.');
  });
}


function exibirTabelas() {

}




