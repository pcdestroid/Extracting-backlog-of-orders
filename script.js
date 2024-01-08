//Required to use Drive API: "service" type settings - Version V2
//Finds an unread email based on the subject, saves the attached xlsx spreadsheet, extracts the data and saves it to a Google spreadsheet in Drive.


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Menu');
  menu.addItem('Carregar Pedidos em atraso', 'pedidos_prazo_expirado');
  menu.addToUi();
}

function pedidos_prazo_expirado() {

  const assuntoAProcurar = 'ORDENS DE COMPRA ENVIADAS E COM PRAZO DE ENTREGA EXPIDRADO';
  const idMyFolder = PropertiesService.getScriptProperties().getProperty('idMyFolder')
  const minhaPasta = DriveApp.getFolderById(idMyFolder);

  // Procurar apenas uma thread com o assunto específico.
  const threads = GmailApp.search(`subject:"${assuntoAProcurar}"`, 0, 1);

  if (threads.length > 0 && threads[0].isUnread()) { // Se houver pelo menos uma thread encontrada e não foi lido.

    threads[0].markRead() // Marca o email como lido

    const messages = threads[0].getMessages();
    const attachments = messages[messages.length - 1].getAttachments(); // Obter anexos da última mensagem da thread.
    for (let j = 0; j < attachments.length; j++) {
      const anexo = attachments[j];

      // remover o arquivo existente, se houver
      const arquivosAntigos = minhaPasta.getFilesByName('RELATORIO_107.XLSX');
      while (arquivosAntigos.hasNext()) {
        const arquivoAntigo = arquivosAntigos.next();
        arquivoAntigo.setTrashed(true);
      }

      // salvar o novo anexo na pasta especificada
      var arquivo = minhaPasta.createFile(anexo).setName('RELATORIO_107.XLSX');

      // Converta o arquivo para o formato Google Sheets
      var arquivoConvertido = Drive.Files.copy({}, arquivo.getId(), { convert: true });

      // Abra a planilha convertida
      var planilha = SpreadsheetApp.openById(arquivoConvertido.id);

      // Acesse a primeira guia (planilha ativa)
      var guia = planilha.getSheets()[0];

      // Obtenha os dados da planilha
      var dados = guia.getDataRange().getValues();

      //Planilha de pedidos com prazo expirado
      var spreadsheetId = PropertiesService.getScriptProperties().getProperty('idThisSheet');
      var planilhaRecebimento = SpreadsheetApp.openById(spreadsheetId);
      var guiaRecebimento = planilhaRecebimento.getSheetByName('Pedidos');
      var dadosRecebimento = guiaRecebimento.getDataRange().getValues();

      // Verifique cada linha na planilha do anexo
      for (let l = 1; l < dados.length; l++) {
        var empresa = dados[l][1]; // Coluna B

        var ordemCompra = dados[l][3]; // Coluna D (Ordem de Compra)

        let prazoEntrega = dados[l][4]; // Coluna E (Prazo de Entrega)

        var fornecedor = dados[l][5]; // Coluna F (Fornecedor)

        var comprador = dados[l][9]; // Coluna J (Comprador)

        let myName = PropertiesService.getScriptProperties().getProperty('myName')

        if (ordemCompra !== "" && ordemCompra !== "OC" && comprador == myName) {

          // Verifique se a Ordem de Compra não existe na planilha "Prazo de entrega"
          var existeOrdemCompra = false;

          for (var m = 0; m < dadosRecebimento.length; m++) {
            if (Number(dadosRecebimento[m][2]) === Number(ordemCompra)) { // Coluna 3
              existeOrdemCompra = true;
              break;
            }
          }

          // Se a solicitação não existe, insira os dados na planilha "Recebimento de Pedidos"
          if (!existeOrdemCompra) {
            guiaRecebimento.appendRow([new Date(), empresa, ordemCompra, prazoEntrega, fornecedor, comprador, 'Atraso', ""]);
          }

        }

      }

      Logger.log('Relatório atualizado')

      // Feche a planilha convertida
      DriveApp.getFileById(arquivoConvertido.id).setTrashed(true);

    }

  } else {
    Logger.log('Nenhum relatório encontrado no email')
    return false;
  }

}

function atualizar_situacao() {
  let spreadsheetId = PropertiesService.getScriptProperties().getProperty('idThisSpreadsheet')
  let spreadsheetDelayOrder = SpreadsheetApp.openById(spreadsheetId);
  let orderGuide = spreadsheetDelayOrder.getSheetByName('Pedidos');
  let backlogOrders = orderGuide.getDataRange().getValues();
  Logger.log(backlogOrders)
}

//##############################################################################################################

function obterEmpresaPeloNome(nome) {

  let idSheetCompanies = PropertiesService.getScriptProperties().getProperty('spreadsheetCompanies')
  let empresas = SpreadsheetApp.openById(idSheetCompanies).getSheetByName('empresas');
  let dados = empresas.getRange(2, 1, empresas.getLastRow(), 10).getValues();
  let empresa = dados.find(dado => dado[2] === nome);
  return empresa
}
//##############################################################################################################
