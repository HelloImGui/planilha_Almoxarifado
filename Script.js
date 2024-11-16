function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Cria o menu principal com submenus
  ui.createMenu('Opções') // Menu principal chamado "Opções"
    .addItem('Abrir Cadastro', 'abrirFormulario') // Opção do submenu "Cadastro"
    .addItem('Abrir Termo', 'JanelaSaida') // Opção do submenu "Termo"
    .addItem('Limpar Termo', 'limparTermo') // Opção do submenu "Limpar"
    .addToUi();  

  atualizarDataTermo(); // Atualiza a data na aba "Termo"
  buscarEquipamentosARetirar(); // Chama a função para atualizar a aba "Dashboard"
}

function atualizarDataTermo() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaManutencao = planilha.getSheetByName('Termo');
  var dataAtual = new Date();
  var dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy");

  abaManutencao.getRange('J4').setValue(dataFormatada);
}

function buscarEquipamentosARetirar() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaDashboard = planilha.getSheetByName('Dashboard');
  
  // Limpa a aba Dashboard antes de atualizar os dados
  if (abaDashboard.getLastRow() > 1) {
    abaDashboard.getRange(2, 1, abaDashboard.getLastRow() - 1, abaDashboard.getLastColumn()).clearContent();
  }
  
  Logger.log("Iniciando a busca de equipamentos com status 'À retirar'");

  // Abas para verificar os equipamentos com status "À retirar"
  var abasEquipamentos = ['D570', 'D580', 'Desktop Dell', 'Notebook HP', 'Notebook Dell', 'Notebook Daten', 'Monitores', 'Monitor Dell', 'Switch', 'Telefone', 'Access Point'];
  var equipamentosARetirarTotal = []; // Armazena todos os equipamentos "À retirar" de todas as abas

  abasEquipamentos.forEach(function(nomeAba) {
    var aba = planilha.getSheetByName(nomeAba);
    if (!aba) {
      Logger.log("Aba '" + nomeAba + "' não encontrada.");
      return;
    }
    
    var ultimaLinha = aba.getLastRow();
    Logger.log("Aba: " + nomeAba + ", Última Linha: " + ultimaLinha);

    if (ultimaLinha > 1) { // Verifica se há dados na aba
      var dados = aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn()).getValues();
      
      // Filtra os equipamentos com status "À retirar"
      var equipamentosARetirar = dados.filter(function(linha) {
        return linha[0] === "À retirar";
      });
      
      // Adiciona os equipamentos filtrados ao total se houver algum
      if (equipamentosARetirar.length > 0) {
        equipamentosARetirarTotal = equipamentosARetirarTotal.concat(equipamentosARetirar);
        Logger.log("Equipamentos 'À retirar' encontrados na aba '" + nomeAba + "': " + equipamentosARetirar.length);
      } else {
        Logger.log("Nenhum equipamento 'À retirar' encontrado na aba '" + nomeAba + "'");
      }
    } else {
      Logger.log("Aba '" + nomeAba + "' está vazia ou não contém linhas de dados suficientes.");
    }
  });

  // Verifica se há equipamentos para transferir antes de tentar copiar para evitar erro
  if (equipamentosARetirarTotal.length > 0) {
    abaDashboard.getRange(abaDashboard.getLastRow() + 1, 1, equipamentosARetirarTotal.length, equipamentosARetirarTotal[0].length).setValues(equipamentosARetirarTotal);
    Logger.log("Equipamentos 'À retirar' adicionados à aba 'Dashboard'.");
  } else {
    Logger.log("Nenhum equipamento com status 'À retirar' foi encontrado para atualizar na aba 'Dashboard'.");
  }
  
  Logger.log("Finalizada a atualização da aba 'Dashboard'.");
}


function onEdit(e) {
  const planilha = e.source.getActiveSheet();
  const range = e.range;
  const colunaEditada = range.getColumn();
  const valorAtual = range.getValue();
  const linhaAtual = range.getRow();

  // Verifica se o status foi alterado para "Entregue" na coluna A da aba "Saída"
  if (colunaEditada === 1 && valorAtual === "Entregue" && planilha.getName() === "Saída") {
    // Copia a linha para a aba "Implantados" com o status "Entregue"
    copiarLinhaParaImplantados(linhaAtual);

    // Exclui a linha da aba "Saída"
    planilha.deleteRow(linhaAtual);
  }
}

// Função para copiar a linha para a aba "Implantados"
function copiarLinhaParaImplantados(linhaSaida) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaSaida = planilha.getSheetByName("Saída");
  const abaImplantados = planilha.getSheetByName("Implantados");

  // Obtém os dados da linha na aba "Saída"
  const dadosLinha = abaSaida.getRange(linhaSaida, 1, 1, abaSaida.getLastColumn()).getValues()[0];

  // Atualiza o status para "Entregue" antes de copiar para "Implantados"
  dadosLinha[0] = "Entregue"; // Coluna A (Status) definida como "Entregue"

  // Adiciona a linha na aba "Implantados"
  abaImplantados.appendRow(dadosLinha);
}










//--------------------------------------------------------------MENU---------------------------------------------

function abrirMenu() {
  var htmlForm = HtmlService.createHtmlOutputFromFile('Menu')
    .setWidth(1700)
    .setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(htmlForm, 'Menu');

}
//--------------------------------------------------------------MENU---------------------------------------------
//------------------------------------------------------------LIMPAR TERMO------------------------------------------------------------

function limparTermo(){
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaTermo = planilha.getSheetByName('Termo');
  abaTermo.getRange('B11:K27').clearContent();
}

//------------------------------------------------------------LIMPAR TERMO------------------------------------------------------------
//--------------------------------------------------------------SAÍDA---------------------------------------------

// Função para buscar o SF em abas específicas apenas com Status "Em estoque"
function buscarSF(sfItem, abasPermitidas) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  for (var j = 0; j < abasPermitidas.length; j++) {
    var aba = planilha.getSheetByName(abasPermitidas[j]);
    var dados = aba.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {  // Começar da linha 2, ignorando cabeçalho
      if (dados[i][0] === "Em estoque" && dados[i][4] == sfItem) {  // Verifica o Status na Coluna A e o SF na Coluna E
        return {
          aba: aba,
          linha: i + 1,  // Ajusta o índice para coincidir com a interface da planilha
          dados: dados[i]
        };
      }
    }
  }
  return null;  // SF não encontrado com Status "Em estoque"
}



function atualizarHora() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName("Termo");
  // Aplica a hora atual na célula J5
  var horaAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
  guia.getRange('J5').setValue(horaAtual);
}

function limparHora() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName("Termo");
  guia.getRange('J5').clear();
}

// Função para obter SFs e Modelos de abas específicas apenas com Status "Em estoque"
function getSFsFromAbas() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var sfDesktopAbas = ['D570', 'D580', 'Desktop Dell'];
  var sfNotebookAbas = ['Notebook Daten', 'Notebook HP', 'Notebook Dell'];
  var sfMonitorAbas = ['Monitores', 'Monitor Dell'];
  var sfSwitchAbas = ['Switch'];
  var sfTelefoneAbas = ['Telefone'];
  var sfAccessPointAbas = ['Access Point'];
  var sfMiniPcAbas = ['Mini PC'];
  var sfCameraRallyAbas = ['Camera Rally'];
  var sfTapAbas = ['TAP'];
  var sfExtTelefoneAbas = ['Extensor Telefone'];

  var sfData = {
    desktop: [],
    notebook: [],
    monitor: [],
    switch: [],
    telefone: [],
    accessPoint: [],
    miniPc: [],
    cameraRally: [],
    tap: [],
    extTelefone: []
  };

  // Função auxiliar para buscar SFs de uma aba específica com status "Em estoque"
  function buscarSFsComStatus(abaNome, categoria) {
    var aba = planilha.getSheetByName(abaNome);
    if (aba) {
      var dados = aba.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) { // Começar da linha 2, ignorando cabeçalho
        if (dados[i][0] === "Em estoque") { // Verifica o Status na Coluna A
          sfData[categoria].push({
            sf: dados[i][4], // Coluna E - SF do item
            modelo: dados[i][3] // Coluna D - Modelo do item
          });
        }
      }
    }
  }

  // Buscar SFs e Modelos para cada categoria apenas com Status "Em estoque"
  sfDesktopAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'desktop');
  });

  sfNotebookAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'notebook');
  });

  sfMonitorAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'monitor');
  });

  sfSwitchAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'switch');
  });

  sfTelefoneAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'telefone');
  });

  sfAccessPointAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'accessPoint');
  });

  sfMiniPcAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'miniPc');
  });

  sfCameraRallyAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'cameraRally');
  });

  sfTapAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'tap');
  });

  sfExtTelefoneAbas.forEach(function(abaNome) {
    buscarSFsComStatus(abaNome, 'extTelefone');
  });

  return sfData;
}


// Função para abrir o formulário de termo
function JanelaSaida() {
  var htmlForm = HtmlService.createHtmlOutputFromFile('Saida')
    .setHeight(700)
    .setWidth(1500);

  SpreadsheetApp.getUi().showModalDialog(htmlForm, 'Saída');
}

// Função para validar os campos obrigatórios
function validarCamposObrigatorios(incidente, ua, responsavel) {
  if (!incidente || !ua || !responsavel) {
    SpreadsheetApp.getUi().alert("Preencha os campos obrigatórios e tente novamente.");
    return false;
  }
  return true;
}

// Função para formatar o SF
function formatarSF(sfItem) {
  if (!sfItem) return null; // Se o campo SF não for preenchido, retornar nulo e continuar o processo

  sfItem = sfItem.toUpperCase().trim();
  if (!/^SF\d{6}$/.test(sfItem)) {
    return "SF" + sfItem.padStart(6, '0');
  }

  if (!/^SF\d{6}$/.test(sfItem)) {
    SpreadsheetApp.getUi().alert("O SF não segue o padrão, tente novamente.");
    return null;
  }
  return sfItem;
}

// Função para gerar PDF do Termo
function gerarPdfTermoSaida(email) {
  var PRINT_OPTIONS = { 
    'size': 7,               // Tamanho do papel A4
    'fzr': false,            // Repetir cabeçalhos de linha
    'portrait': false,       // Modo retrato
    'fitw': true,            // Ajustar à página
    'gridlines': false,      // Não mostrar linhas de grade
    'printtitle': false,
    'sheetnames': false,
    'pagenum': 'UNDEFINED',  // Não mostrar número de página
    'attachment': false
  };

  var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

  SpreadsheetApp.flush();
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName("Termo");
  var range = guia.getRange("B2:K42").activate();
  var incidente = String(guia.getRange('E11').getValue()); // Converte o valor para string
  var login = String(email.split('@')[0]); // Extrai o login do email

  var gid = guia.getSheetId();
  
  var printRange = objectToQueryString({
    'c1': range.getColumn() - 1,
    'r1': range.getRow() - 1,
    'c2': range.getColumn() + range.getWidth() - 1,
    'r2': range.getRow() + range.getHeight() - 1
  });
  
  var url = planilha.getUrl().replace(/edit$/, '') + 'export?format=pdf' + PDF_OPTS + printRange + "&gid=" + gid;

  try {
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });

    // Define o nome do arquivo com "incidente + _ + login"
    var blob = response.getBlob().setName(incidente + '_' + login + '.pdf'); 

    if (!blob) {
      Logger.log('Erro ao gerar o PDF: o arquivo blob está vazio.');
      throw new Error('Erro ao gerar o PDF.');
    }

    var folderId = '1YFXFftQcv4v4PIvVc43W-WHSAAwE4vmV';
    var folder = DriveApp.getFolderById(folderId);

    Logger.log('Tentando criar arquivo na pasta do Drive com o ID: ' + folderId);

    if (!folder) {
      Logger.log('Erro ao encontrar a pasta no Google Drive.');
      throw new Error('Pasta não encontrada no Google Drive.');
    }

    var file = folder.createFile(blob);
    Logger.log('Arquivo criado com sucesso: ' + file.getUrl());

    return file.getUrl();
  } catch (error) {
    Logger.log('Erro ao criar o arquivo no Google Drive: ' + error.message);
    throw new Error('Erro ao tentar criar o arquivo no Google Drive.');
  }
}




function processarDadosTermo(sfDesktop, sfNotebook, sfMonitor1, sfMonitor2, sfSwitch, sfTelefone, sfAccessPoint, sfMiniPc, sfCameraRally, sfTap, extTelefone, incidente, ua, responsavel, email, patrimonioRelacionado, termoTipo, ...perifericosSelecionados) {
  limparTermo();
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaTermo = planilha.getSheetByName('Termo');
  var abaPerifericos = planilha.getSheetByName('Periféricos');
  var abaSaidaPerifericos = planilha.getSheetByName('Saída_Periféricos');
  var linhaTermo = 11; // Primeira linha para preencher na aba Termo
  var dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Preenche o valor de "Substituição" ou "Empréstimo" em E26, conforme o checkbox selecionado
  if (termoTipo) {
    abaTermo.getRange("E26").setValue(termoTipo);
  }

  if (patrimonioRelacionado) {
    abaTermo.getRange(linhaTermo, 2).setValue(patrimonioRelacionado);  // Coluna B para patrimônio
  }

  // Definir as abas permitidas para cada tipo de SF
  var abasDesktop = ['D570', 'D580', 'Desktop Dell'];
  var abasNotebook = ['Notebook Daten', 'Notebook HP', 'Notebook Dell'];
  var abasMonitores = ['Monitores', 'Monitor Dell'];
  var abasSwitch = ['Switch'];
  var abasTelefone = ['Telefone'];
  var abasAccessPoint = ['Access Point'];
  var abasMiniPc = ['Mini PC'];
  var abasCameraRally = ['Camera Rally'];
  var abasTap = ['TAP'];
  var abasExtTelefone = ['Extensor Telefone'];

  // Variáveis para armazenar dados dos SFs encontrados
  var desktopData = sfDesktop ? buscarSF(formatarSF(sfDesktop), abasDesktop) : null;
  var notebookData = sfNotebook ? buscarSF(formatarSF(sfNotebook), abasNotebook) : null;
  var monitor1Data = sfMonitor1 ? buscarSF(formatarSF(sfMonitor1), abasMonitores) : null;
  var monitor2Data = sfMonitor2 ? buscarSF(formatarSF(sfMonitor2), abasMonitores) : null;
  var switchData = sfSwitch ? buscarSF(formatarSF(sfSwitch), abasSwitch) : null;
  var telefoneData = sfTelefone ? buscarSF(formatarSF(sfTelefone), abasTelefone) : null;
  var accessPointData = sfAccessPoint ? buscarSF(formatarSF(sfAccessPoint), abasAccessPoint) : null;
  var miniPcData = sfMiniPc ? buscarSF(formatarSF(sfMiniPc), abasMiniPc) : null;
  var cameraRallyData = sfCameraRally ? buscarSF(formatarSF(sfCameraRally), abasCameraRally) : null;
  var tapData = sfTap ? buscarSF(formatarSF(sfTap), abasTap) : null;
  var extTelefoneData = extTelefone ? buscarSF(formatarSF(extTelefone), abasExtTelefone) : null;

  // Função para atualizar dados na própria aba de origem e mover para a aba "Termo" e "Saída"
  function atualizarNaAbaOrigem(dadosSF) {
    var aba = dadosSF.aba;
    var linha = dadosSF.linha;
    var dados = dadosSF.dados;
    var tipoConcatenado = dados[1] + " " + dados[2] + " " + dados[3];  // Concatena tipo, marca e modelo

    // Atualizar status para "À retirar", incidente, data de alteração, usuário e e-mail
    aba.getRange(linha, 1).setValue("À retirar");
    aba.getRange(linha, 8).setValue(incidente);
    aba.getRange(linha, 9).setValue(dataAtual);
    aba.getRange(linha, 10).setValue(responsavel);
    aba.getRange(linha, 11).setValue(email);

    // Preencher os dados na aba "Termo"
    abaTermo.getRange(linhaTermo, 2).setValue(dados[4]);  // Coluna B: Patrimônio (SF)
    abaTermo.getRange(linhaTermo, 3).setValue(dados[5]);  // Coluna C: Número de Série
    abaTermo.getRange(linhaTermo, 4).setValue(tipoConcatenado);  // Coluna D: Descrição (Tipo + Marca + Modelo)
    abaTermo.getRange(linhaTermo, 5).setValue(incidente);  // Coluna E: Incidente
    abaTermo.getRange(linhaTermo, 6).setValue(ua);  // Coluna F: U.A.
    abaTermo.getRange(linhaTermo, 9).setValue(responsavel);  // Coluna I: Responsável
    linhaTermo++;

    // Copiar dados para a aba "Saída" antes de excluir a linha da aba de origem
    var abaSaida = planilha.getSheetByName("Saída");
    abaSaida.appendRow([
      "À retirar",         // Status atualizado
      dados[1],            // Tipo
      dados[2],            // Marca
      dados[3],            // Modelo
      dados[4],            // SF do item
      dados[5],            // S/N
      dados[6],            // MAC
      incidente,           // Observações (Incidente)
      dataAtual,           // Data de Alteração
      responsavel,         // Usuário
      email                // E-mail
      // Link do PDF será adicionado após geração
    ]);

    // Excluir a linha da aba de origem após copiar para "Saída"
    aba.deleteRow(linha);
  }

  // Processar os dados dos SFs encontrados
  [desktopData, notebookData, monitor1Data, monitor2Data, switchData, telefoneData, accessPointData, miniPcData, cameraRallyData, tapData, extTelefoneData].forEach(function(dadosSF) {
    if (dadosSF) atualizarNaAbaOrigem(dadosSF);
  });

  // Processar periféricos
  perifericosSelecionados.forEach(function(periferico) {
    if (periferico) {
      var dadosPerifericos = abaPerifericos.getRange('A:C').getValues();
      for (var i = 1; i < dadosPerifericos.length; i++) {
        if (dadosPerifericos[i][0] === periferico) {  // Se o periférico estiver na aba
          var quantidadeAtual = dadosPerifericos[i][2]; // Coluna C - Quantidade
          if (quantidadeAtual > 0) {
            abaPerifericos.getRange(i + 1, 3).setValue(quantidadeAtual - 1);
            abaTermo.getRange(linhaTermo, 4).setValue(periferico);
            abaTermo.getRange(linhaTermo, 5).setValue(incidente);
            abaTermo.getRange(linhaTermo, 6).setValue(ua);
            abaTermo.getRange(linhaTermo, 9).setValue(responsavel);
            linhaTermo++;

            // Adicionar linha na aba "Saída_Periféricos"
            abaSaidaPerifericos.appendRow([
              periferico,             // Nome do periférico
              1,                      // Quantidade
              responsavel,            // Responsável
              incidente,              // Incidente
              patrimonioRelacionado,  // Patrimônio relacionado
              dataAtual               // Data
              // Link do PDF será adicionado após geração
            ]);
          }
        }
      }
    }
  });

  // Gerar o PDF e obter o link
  var linkPdf = gerarPdfTermoSaida(email);

  // Inserir o link do PDF nas abas "Saída" e "Saída_Periféricos"
  var ultimaLinhaSaida = planilha.getSheetByName("Saída").getLastRow();
  planilha.getSheetByName("Saída").getRange(ultimaLinhaSaida, 12).setValue(linkPdf);

  var ultimaLinhaSaidaPerifericos = abaSaidaPerifericos.getLastRow();
  abaSaidaPerifericos.getRange(ultimaLinhaSaidaPerifericos, 7).setValue(linkPdf);  // Coluna G para o link na "Saída_Periféricos"

  atualizarHora();
  limparHora();

  SpreadsheetApp.getUi().alert("Termo criado e PDF gerado com sucesso.");
}

// Função para inserir dados no termo, garantindo validação
function inserirDadosTermo(sfDesktop, sfNotebook, sfMonitor1, sfMonitor2, sfSwitch, sfTelefone, sfAccessPoint, sfMiniPc, sfCameraRally, sfTap, sfExtTelefone, incidente, ua, responsavel, email, patrimonioRelacionado, ...perifericosSelecionados) {
  if (!validarCamposObrigatorios(incidente, ua, responsavel)) return;
  processarDadosTermo(sfDesktop, sfNotebook, sfMonitor1, sfMonitor2, sfSwitch, sfTelefone, sfAccessPoint, sfMiniPc, sfCameraRally, sfTap, sfExtTelefone, incidente, ua, responsavel, email, patrimonioRelacionado, ...perifericosSelecionados);
}



//--------------------------------------------------------------SAÍDA---------------------------------------------





//------------------------------------------------------------ENTRADA------------------------------------------------------------

// Abre o formulário de entrada de equipamentos
function abrirFormularioEntrada() {
  var htmlForm = HtmlService.createHtmlOutputFromFile('Entrada')
    .setWidth(2200)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(htmlForm, 'Entrada de Equipamentos');
}

// Gera o termo de entrada e recolhimento, incluindo a assinatura do funcionário e registrando dados do usuário
function gerarTermoEntradaERecolhimento(wo, nomeUsuario, emailUsuario, desktop, notebook, monitor, telefone, switchs, accessPoints, perifericosSelecionados, funcionario) {
  try {
    // Depuração inicial
    Logger.log("Iniciando a geração do termo...");
    Logger.log("Dados coletados: WO - " + wo + ", Funcionário - " + funcionario + ", Nome do Usuário - " + nomeUsuario + ", E-mail do Usuário - " + emailUsuario);

    // Armazena os dados temporariamente no cache
    var cache = CacheService.getUserCache();
    cache.put('wo', JSON.stringify(wo), 300); // Armazena por 5 minutos (300 segundos)
    cache.put('desktop', JSON.stringify(desktop), 300);
    cache.put('notebook', JSON.stringify(notebook), 300);
    cache.put('monitor', JSON.stringify(monitor), 300);
    cache.put('telefone', JSON.stringify(telefone), 300);
    cache.put('switchs', JSON.stringify(switchs), 300);
    cache.put('accessPoints', JSON.stringify(accessPoints), 300);
    cache.put('perifericosSelecionados', JSON.stringify(perifericosSelecionados), 300);
    cache.put('nomeUsuario', nomeUsuario, 300);
    cache.put('emailUsuario', emailUsuario, 300);

    Logger.log("Dados armazenados no cache com sucesso!");

    // Autentica e insere a assinatura com base no funcionário autenticado
    autenticarAssinatura(funcionario);
    Logger.log("Assinatura autenticada com sucesso.");

    // Depuração final
    Logger.log("Termo de entrada e recolhimento gerado com sucesso.");
  } catch (error) {
    Logger.log("Erro ao gerar termo: " + error.message);
    throw new Error("Erro ao gerar termo: " + error.message);
  }
}


// Função chamada para autenticar e inserir a assinatura do funcionário
function autenticarAssinatura(funcionario) {
  try {
    Logger.log("Iniciando autenticação de assinatura...");

    // Recupera os dados do cache
    var cache = CacheService.getUserCache();
    var wo = JSON.parse(cache.get('wo'));
    var desktop = JSON.parse(cache.get('desktop'));
    var notebook = JSON.parse(cache.get('notebook'));
    var monitor = JSON.parse(cache.get('monitor'));
    var telefone = JSON.parse(cache.get('telefone'));
    var switchs = JSON.parse(cache.get('switchs'));
    var accessPoints = JSON.parse(cache.get('accessPoints'));
    var perifericosSelecionados = JSON.parse(cache.get('perifericosSelecionados')) || [];
    var nomeUsuario = cache.get('nomeUsuario');
    var emailUsuario = cache.get('emailUsuario');

    Logger.log("Dados recuperados do cache com sucesso.");

    // Preenche a aba "Termo" com os dados dos equipamentos
    preencherAbaTermo(wo, desktop, notebook, monitor, telefone, switchs, accessPoints, perifericosSelecionados);

    // Insere a assinatura do funcionário
    inserirAssinatura(funcionario);
    Logger.log("Assinatura inserida.");

    // Gera o PDF e salva no Drive
    var linkPdf = gerarPdfTermoEntrada(emailUsuario);

    // Insere os dados com o link do PDF nas abas correspondentes, incluindo nome e e-mail do usuário
    inserirDadosComLink(wo, nomeUsuario, emailUsuario, desktop, notebook, monitor, telefone, switchs, accessPoints, perifericosSelecionados, linkPdf);

    SpreadsheetApp.getUi().alert('Termo gerado, PDF criado e dados de entrada cadastrados com sucesso. Link do PDF inserido.');
  } catch (error) {
    Logger.log("Erro na autenticação de assinatura: " + error.message);
    throw new Error("Erro na autenticação de assinatura: " + error.message);
  }
}



// Ajuste na função para preencher a aba "Termo" com os periféricos em uma única linha
function preencherAbaTermo(wo, desktop, notebook, monitor, telefone, switchs, accessPoints, perifericosSelecionados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var termoAba = planilha.getSheetByName('Termo');

  limparTermo()  

  var linhaInicial = 11;
  var responsavel = "Erik Mesel Ferreira Pires";
  var ua = "35241";

  // Função auxiliar para preencher a aba Termo
  function preencherTermo(itens, tipo) {
    itens.forEach(function(item) {
      termoAba.getRange(linhaInicial, 2).setValue(item.sf); // Coluna B: SF
      termoAba.getRange(linhaInicial, 3).setValue(item.sn); // Coluna C: S/N
      
      // Verifica se o item é Telefone ou Switch para definir o conteúdo da coluna D
      if (tipo === "Telefone" || tipo === "Switch" || tipo === "Access Point") {
        termoAba.getRange(linhaInicial, 4).setValue(tipo + " " + item.modelo); // Apenas o modelo
      } else {
        termoAba.getRange(linhaInicial, 4).setValue(tipo + " " + item.marca + " " + item.modelo); // Coluna D: Descrição
      }

      termoAba.getRange(linhaInicial, 5).setValue(item.wo || wo); // Coluna E: WO, individual ou geral
      termoAba.getRange(linhaInicial, 6).setValue(ua); // Coluna F: U.A.
      termoAba.getRange(linhaInicial, 9).setValue(responsavel); // Coluna I: Responsável
      linhaInicial++;
    });
  }

  preencherTermo(desktop, "Desktop");
  preencherTermo(notebook, "Notebook");
  preencherTermo(monitor, "Monitor");
  preencherTermo(telefone, "Telefone");
  preencherTermo(switchs, "Switch");
  preencherTermo(accessPoints, "Access Point");

  // Preenche o termo para periféricos em uma única linha por periférico
  perifericosSelecionados.forEach(function(periferico) {
    termoAba.getRange(linhaInicial, 2).setValue('N/A'); 
    termoAba.getRange(linhaInicial, 3).setValue('N/A');
    termoAba.getRange(linhaInicial, 4).setValue(periferico.quantidade + " " + periferico.nome); // Única linha com quantidade e nome
    termoAba.getRange(linhaInicial, 5).setValue(wo);
    termoAba.getRange(linhaInicial, 6).setValue(ua);
    termoAba.getRange(linhaInicial, 9).setValue(responsavel);
    linhaInicial++;
  });
  
  termoAba.getRange("E26").setValue("Recolhimento").setFontWeight("bold");
}



// Função para inserir a assinatura de acordo com o funcionário autenticado
function inserirAssinatura(nome) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaTermo = planilha.getSheetByName('Termo');
  
  var assinaturaId;

  // Define o ID da assinatura de acordo com o funcionário escolhido
  switch (nome) {
    case 'Guilherme':
      assinaturaId = '15Tg9eEag9aiOVBa0MoCZtNerkhOOyvSa'; // Substitua com o ID real
      break;
    case 'Kaique':
      assinaturaId = '1mCKh6baX2Dc7BSN9b0ynINthqT9oHsuP'; // Substitua com o ID real
      break;
    case 'William':
      assinaturaId = '1SN5Oobc6ayjZmbAhVDnQwfCNemUyqxMa'; // Substitua com o ID real
      break;
    case 'Habner':
      assinaturaId = '19Od663A0sjWXv8Ekw6Qosmx3ea61kEK9'; // Substitua com o ID real
      break;
    default:
      SpreadsheetApp.getUi().alert("Funcionário não encontrado.");
      return;
  }

  try {
    var assinatura = DriveApp.getFileById(assinaturaId);
    Logger.log("Assinatura encontrada: " + assinatura.getName());

    // Insere a assinatura na aba "Termo"
    var range = abaTermo.getRange('C35');
    var imagemBlob = assinatura.getBlob();

    // Insere a imagem da assinatura no termo
    abaTermo.insertImage(imagemBlob, range.getColumn(), range.getRow());

    SpreadsheetApp.getUi().alert("Assinatura inserida com sucesso.");
    
  } catch (e) {
    Logger.log("Erro ao buscar assinatura: " + e.message);
    SpreadsheetApp.getUi().alert("Erro: Assinatura não encontrada.");
  }
}

// Função para gerar o PDF do termo e salvar no Google Drive
function gerarPdfTermoEntrada(emailUsuario) {
  var PRINT_OPTIONS = { 
    'size': 7,               // Tamanho do papel A4
    'fzr': false,            // Repetir cabeçalhos de linha
    'portrait': false,        // Modo retrato
    'fitw': true,            // Ajustar à página
    'gridlines': false,      // Não mostrar linhas de grade
    'printtitle': false,
    'sheetnames': false,
    'pagenum': 'UNDEFINED',  // Não mostrar número de página
    'attachment': false
  };

  var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

  SpreadsheetApp.flush();
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName("Termo");
  var range = guia.getRange("B2:K43").activate();
  var incidente = String(guia.getRange('E11').getValue()); // Converte o valor para string
  
  var login = String(emailUsuario.split('@')[0]); // Extrai o login do email

  var gid = guia.getSheetId();
  
  var printRange = objectToQueryString({
    'c1': range.getColumn() - 1,
    'r1': range.getRow() - 1,
    'c2': range.getColumn() + range.getWidth() - 1,
    'r2': range.getRow() + range.getHeight() - 1
  });
  
  var url = planilha.getUrl().replace(/edit$/, '') + 'export?format=pdf' + PDF_OPTS + printRange + "&gid=" + gid;

  try {
    // Faz a requisição para gerar o PDF
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });

    var blob = response.getBlob().setName(incidente + '_' + login + '.pdf');

    // Verifica se o arquivo foi criado corretamente
    if (!blob) {
      Logger.log('Erro ao gerar o PDF: o arquivo blob está vazio.');
      throw new Error('Erro ao gerar o PDF.');
    }

    // Obtém a pasta pelo ID fornecido e salva o PDF nela
    var folderId = '1r96tIdBdSb7jnYZyl2zykeXgG8K0tDQG'; // Substitua com o ID da sua pasta no Google Drive
    var folder = DriveApp.getFolderById(folderId);

    Logger.log('Tentando criar arquivo na pasta do Drive com o ID: ' + folderId);

    // Verifica se a pasta foi encontrada corretamente
    if (!folder) {
      Logger.log('Erro ao encontrar a pasta no Google Drive.');
      throw new Error('Pasta não encontrada no Google Drive.');
    }

    var file = folder.createFile(blob);
    Logger.log('Arquivo criado com sucesso: ' + file.getUrl());

    apagarAssinaturaAbaTermo();
    
    // Retorna o link do arquivo salvo
    return file.getUrl();
  } catch (error) {
    Logger.log('Erro ao criar o arquivo no Google Drive: ' + error.message);
    throw new Error('Erro ao tentar criar o arquivo no Google Drive.');
  }
}

// Função para apagar a assinatura (imagem) da aba Termo
function apagarAssinaturaAbaTermo() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaTermo = planilha.getSheetByName('Termo');
  
  // Obtém todas as imagens na aba Termo
  var imagens = abaTermo.getImages();
  
  // Verifica se há imagens e apaga a imagem na posição desejada
  imagens.forEach(function(imagem) {
    var posicaoImagem = imagem.getAnchorCell().getA1Notation();
    
    // Verifica se a imagem está na célula D36 (ou qualquer célula onde a assinatura foi inserida)
    if (posicaoImagem === 'C35') {
      imagem.remove();
      Logger.log("Assinatura apagada da aba Termo.");
    }
  });
}



// Função auxiliar para criar query string de parâmetros de impressão
function objectToQueryString(obj) {
  return Object.keys(obj).map(function(key) {
    return Utilities.formatString('&%s=%s', key, obj[key]);
  }).join('');
}

// Insere os dados do termo nas abas correspondentes e registra o link do PDF, incluindo nome e e-mail do usuário
function inserirDadosComLink(wo, nomeUsuario, emailUsuario, desktop, notebook, monitor, telefone, switchs, accessPoints, perifericosSelecionados, linkPdf) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaSucata = planilha.getSheetByName('Sucata'); // Referência à aba Sucata

  // Insere desktops
  desktop.forEach(function(desktop) {
    var woDesktop = desktop.wo || wo;
    if (desktop.sucata) {
      abaSucata.appendRow(['Sucata', 'Desktop', desktop.marca, desktop.modelo, desktop.sf, desktop.sn, '', woDesktop, new Date(), nomeUsuario, emailUsuario, linkPdf]); // MAC vazio
    } else {
      var aba;
      if (desktop.modelo === 'D570') aba = planilha.getSheetByName('D570');
      else if (desktop.modelo === 'D580') aba = planilha.getSheetByName('D580');
      else if (desktop.modelo === 'Optiplex 9020') aba = planilha.getSheetByName('Desktop Dell');

      aba.appendRow(['Em estoque', 'Desktop', desktop.marca, desktop.modelo, desktop.sf, desktop.sn, '', woDesktop, new Date(), nomeUsuario, emailUsuario, linkPdf]); // MAC vazio
    }
  });

  // Insere notebooks
  notebook.forEach(function(notebook) {
    var woNotebook = notebook.wo || wo;
    if (notebook.sucata) {
      abaSucata.appendRow(['Sucata', 'Notebook', notebook.marca, notebook.modelo, notebook.sf, notebook.sn, '', woNotebook, new Date(), nomeUsuario, emailUsuario, linkPdf]); // MAC vazio
    } else {
      var aba;
      if (notebook.modelo === 'DT02-M4') aba = planilha.getSheetByName('Notebook Daten');
      else if (notebook.modelo === 'EB640G9') aba = planilha.getSheetByName('Notebook HP');
      else if (notebook.modelo === 'Latitude 14 5420') aba = planilha.getSheetByName('Notebook Dell');

      aba.appendRow(['Em estoque', 'Notebook', notebook.marca, notebook.modelo, notebook.sf, notebook.sn, '', woNotebook, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    }
  });

  // Insere monitores
  monitor.forEach(function(monitor) {
    var woMonitor = monitor.wo || wo;
    if (monitor.sucata) {
      abaSucata.appendRow(['Sucata', 'Monitor', monitor.marca, monitor.modelo, monitor.sf, monitor.sn, '', woMonitor, new Date(), nomeUsuario, emailUsuario, linkPdf]); // MAC vazio
    } else {
      var aba = monitor.marca === 'Dell' ? planilha.getSheetByName('Monitor Dell') : planilha.getSheetByName('Monitores');
      aba.appendRow(['Em estoque', 'Monitor', monitor.marca, monitor.modelo, monitor.sf, monitor.sn, '', woMonitor, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    }
  });

  // Insere telefones
  telefone.forEach(function(telefone) {
    var woTelefone = telefone.wo || wo;
    if (telefone.sucata) {
      abaSucata.appendRow(['Sucata', 'Telefone', 'Cisco', telefone.modelo, telefone.sf, telefone.sn, telefone.mac, woTelefone, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    } else {
      var aba = planilha.getSheetByName('Telefone');
      aba.appendRow(['Em estoque', 'Telefone', 'Cisco', telefone.modelo, telefone.sf, telefone.sn, telefone.mac, woTelefone, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    }
  });

  // Insere switchs
  switchs.forEach(function(switchItem) {
    var woSwitch = switchItem.wo || wo;
    if (switchItem.sucata) {
      abaSucata.appendRow(['Sucata', 'Switch', 'Cisco', switchItem.modelo, switchItem.sf, switchItem.sn, switchItem.mac, woSwitch, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    } else {
      var aba = planilha.getSheetByName('Switch');
      aba.appendRow(['Em estoque', 'Switch', 'Cisco', switchItem.modelo, switchItem.sf, switchItem.sn, switchItem.mac, woSwitch, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    }
  });

  // Insere Access Points
  accessPoints.forEach(function(ap) {
    var woAP = ap.wo || wo;
    if (ap.sucata) {
      abaSucata.appendRow(['Sucata', 'Access Point', 'Cisco', ap.modelo, ap.sf, ap.sn, ap.mac, woAP, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    } else {
      var aba = planilha.getSheetByName('Access Point');
      aba.appendRow(['Em estoque', 'Access Point', 'Cisco', ap.modelo, ap.sf, ap.sn, ap.mac, woAP, new Date(), nomeUsuario, emailUsuario, linkPdf]);
    }
  });

  // Insere periféricos
  var abaPerifericos = planilha.getSheetByName('Entrada_Periféricos');
  perifericosSelecionados.forEach(function(periferico) {
    if (periferico.nome && periferico.quantidade > 0) {
      abaPerifericos.appendRow([periferico.nome, periferico.quantidade, wo, nomeUsuario, emailUsuario, new Date(), linkPdf]); 
    }
  });
}

//------------------------------------------------------------ENTRADA------------------------------------------------------------


//------------------------------------------------------------MANUTENÇÃO------------------------------------------------------------

// Função para pegar os SFs e Modelos das abas limitadas
function getSFsFromSpecificAbas() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abas = ['D570', 'D580', 'Desktop Dell', 'Notebook Daten', 'Notebook HP', 'Notebook Dell'];
  var todosSFs = [];

  for (var j = 0; j < abas.length; j++) {
    var aba = planilha.getSheetByName(abas[j]);
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {  // Começa na linha 2 para ignorar o cabeçalho
      todosSFs.push({
        sf: dados[i][4],  // Coluna A - SF do item
        modelo: dados[i][3] // Coluna C - Modelo do item
      });
    }
  }
  return todosSFs;
}


// Função para abrir o formulário de manutenção
function abrirFormularioManutencao() {
  var htmlForm = HtmlService.createHtmlOutputFromFile('Manutencao')
    .setHeight(600)
    .setWidth(850);
  SpreadsheetApp.getUi().showModalDialog(htmlForm, 'Manutenção Preventiva');
}

// Função para buscar dados do SF na planilha
function buscarDadosSF(sfItem) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abas = planilha.getSheets();  // Obtem todas as abas

  // Loop por todas as abas para encontrar o SF
  for (var j = 0; j < abas.length; j++) {
    var aba = abas[j];
    var nomeAba = aba.getName();

    if (nomeAba === "Implantados" || nomeAba === "Sucata" || nomeAba === "Manutenção" || nomeAba === "Periféricos" || nomeAba === "Termo" || nomeAba === "Menu" || nomeAba === "Saída_Periféricos") {
      continue; // Ignora essas abas
    }

    var dados = aba.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {  // Começar da linha 2, ignorando cabeçalho
      if (dados[i][4] === sfItem) {
        var tipoModelo = dados[i][1] + " " + dados[i][3];  // Concatena Tipo + Modelo
        aba.getRange(i + 1, 1).setValue("Manutenção");  // Atualiza o status para "Manutenção"
        return { sfItem: dados[i][4], descricao: tipoModelo };
      }
    }
  }
  return null;  // Retorna null se o SF não for encontrado
}

// Função para enviar os dados para a aba "Manutenção"
function enviarDadosManutencao(sfTable) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaManutencao = planilha.getSheetByName('Manutenção');
  var linha = 10;  // A partir da linha 9 para baixo

  for (var i = 0; i < sfTable.length; i++) {
    abaManutencao.getRange(linha, 2).setValue(sfTable[i].patrimonio);  // Coluna A
    abaManutencao.getRange(linha, 3).setValue(sfTable[i].descricao);   // Coluna B
    linha++;
  }
}

// Função para atualizar a data do termo ao abrir a planilha
function atualizarDataAtualTermo() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaManutencao = planilha.getSheetByName('Manutenção');
  var dataAtual = new Date();
  var dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy");

  abaManutencao.getRange('D4').setValue(dataFormatada);  // Atualiza a célula C3
}

function getSFsFromPatrimonios() {
  var abas = ['D570', 'D580', 'Desktop Dell', 'Notebook Daten', 'Notebook Dell', 'Notebook HP'];
  var sfs = [];

  var planilha = SpreadsheetApp.getActiveSpreadsheet();

  // Percorre todas as abas de patrimonios
  abas.forEach(function(nomeAba) {
    var sheet = planilha.getSheetByName(nomeAba);
    if (!sheet) return;

    var data = sheet.getRange('A:E').getValues(); // Obtém as colunas A a E (SF, Status, Modelo, Tipo)

    // Percorre as linhas para verificar o status "Manutenção"
    for (var i = 1; i < data.length; i++) { // Começa na linha 1 para ignorar o cabeçalho
      var status = data[i][0]; // Coluna E: Status
      if (status === 'Manutenção') {
        var sf = data[i][4]; // Coluna A: SF
        var modelo = data[i][3]; // Coluna C: Modelo
        var tipo = data[i][1]; // Coluna D: Tipo

        // Adiciona à lista de SFs com descrição completa incluindo o SF
        sfs.push(sf + ' - ' + tipo + ' ' + modelo);
      }
    }
  });

  return sfs; // Retorna a lista de SFs encontrados com a descrição
}

// Função para gerar o PDF do termo e retornar o link
function gerarPdfManutencao() {
  var PRINT_OPTIONS = { 
    'size': 7,               // Tamanho do papel. 0 = carta, 1 = tablóide, 2 = Ofício, 3 = declaração, 4 = executivo, 5 = fólio, 6 = A3, 7 = A4, 8 = A5, 9 = B4, 10 = B
    'fzr': false,            // repetir cabeçalhos de linha
    'portrait': true,        // false = paisagem
    'fitw': true,            // ajustar a janela ou tamanho real
    'gridlines': false,      // mostrar linhas de grade
    'printtitle': false,
    'sheetnames': false,
    'pagenum': 'UNDEFINED',  // CENTRO = mostrar números de página / UNDEFINED = não mostrar
    'attachment': false,
    'margins': {
      'top': 0.25,          // Margem superior estreita (0.25 polegadas)
      'bottom': 0.25,       // Margem inferior estreita (0.25 polegadas)
      'left': 0.25,         // Margem esquerda estreita (0.25 polegadas)
      'right': 0.25         // Margem direita estreita (0.25 polegadas)
    }
  };

  var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

  SpreadsheetApp.flush();
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName("Manutenção");
  var range = guia.getRange("B2:D48").activate();
  // Obtém a data atual da célula J4 e formata para o padrão dd/mm/yyyy
  var data = guia.getRange('D4').getValue();
  var dataFormatada = Utilities.formatDate(new Date(data), Session.getScriptTimeZone(), "dd-MM-yyyy"); // Formata a data

  var gid = guia.getSheetId();
  
  var printRange = objectToQueryString({
    'c1': range.getColumn() - 1,
    'r1': range.getRow() - 1,
    'c2': range.getColumn() + range.getWidth() - 1,
    'r2': range.getRow() + range.getHeight() - 1
  });
  
  var url = planilha.getUrl().replace(/edit$/, '') + 'export?format=pdf' + PDF_OPTS + printRange + "&gid=" + gid;

  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  var blob = response.getBlob().setName('Manutenção' + '_' + dataFormatada + '.pdf');
  
  // Obtém a pasta pelo ID fornecido e salva o PDF nela
  var folder = DriveApp.getFolderById('1KQPFnNfxyjnaDs7LL5zgmUxtxCKgazXV');
  var file = folder.createFile(blob);
  
  // Retorna o link do arquivo salvo
  var fileUrl = file.getUrl();
  
  return fileUrl; // Retorna o link do PDF gerado
}

function inserirLinkPdfNaManutencao() {
  // Gera o PDF e obtém o link
  var fileUrl = gerarPdfManutencao();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Manutenção');
  
  // Procura a última linha preenchida na coluna G
  var colunaG = sheet.getRange('H:H').getValues(); // Obtém todos os valores da coluna G
  var ultimaLinhaPreenchida = 0;
  
  // Encontra a última linha preenchida na coluna G
  for (var i = colunaG.length - 1; i >= 0; i--) {
    if (colunaG[i][0] !== '') {
      ultimaLinhaPreenchida = i + 2; // Linha encontrada (add +1 pois array é zero-based)
      break;
    }
  }

  // Insere o link na última linha preenchida da coluna H
  sheet.getRange(ultimaLinhaPreenchida, 8).setValue(fileUrl);
  
  SpreadsheetApp.getUi().alert('PDF gerado e link inserido com sucesso na última linha preenchida da coluna H.');
}
//-------------------------------------------------------------------MANUTENÇÃO-------------------------------------------------------------------

//-------------------------------------------------------------------PEÇAS MANUTENÇÃO-------------------------------------------------------------------

function abrirFormularioPecas() {
  var htmlForm = HtmlService.createHtmlOutputFromFile('PecasManutencao')
    .setHeight(900)
    .setWidth(1000);

  SpreadsheetApp.getUi().showModalDialog(htmlForm, 'Seleção de Peças e Equipamentos');
}

function getPecasFromPlanilha() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Peças');
  var data = sheet.getRange('A2:C').getValues(); // Obtém os dados a partir da linha 2 (ignora o cabeçalho)
  var pecas = [];

  // Itera pelos dados para pegar nome da peça e quantidade
  for (var i = 0; i < data.length; i++) {
    var nomePeca = data[i][0]; // Coluna A: Nome do item
    var quantidade = data[i][2]; // Coluna C: Quantidade

    if (nomePeca && quantidade > 0) { // Adiciona apenas se a quantidade for maior que 0
      pecas.push({ nome: nomePeca, quantidade: quantidade });
    }
  }
  return pecas; // Retorna a lista de peças com nome e quantidade
}

// Função para processar o envio de múltiplas peças selecionadas
function processarPecasSelecionadas(pecasSelecionadas, incidente, responsavel, ua, sfSelecionado) {
  var linhaTermo = 11;

  // Limpa o conteúdo das linhas na aba "Termo" antes de inserir novas peças
  limparTermo();

  for (var i = 0; i < pecasSelecionadas.length; i++) {
    var peca = pecasSelecionadas[i];

    // Monta os dados da peça para enviar ao Termo
    var dados = {
      sf: sfSelecionado,
      item: peca.item,
      sn: peca.sn, // Inclui o S/N da peça
      incidente: incidente,
      ua: ua,
      responsavel: responsavel
    };

    // Mover a peça para o Termo na linha correspondente
    moverParaTermo(dados, linhaTermo);

    // Incrementa a linha para a próxima peça
    linhaTermo++;
  }
}

// Função para enviar os dados de várias peças para o Termo
function enviarVariasPecas(pecasSelecionadas, incidente, responsavel, ua, sfSelecionado) {
  if (pecasSelecionadas.length > 0) {
    processarPecasSelecionadas(pecasSelecionadas, incidente, responsavel, ua, sfSelecionado);
  }
}

// Função para mover os dados da peça para a aba "Termo"
function moverParaTermo(dados, linha) {
  var sheetTermo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Termo');

  // Aloca os dados corretamente nas colunas da aba "Termo"
  sheetTermo.getRange(linha, 2).setValue(dados.sf);           // Coluna B: SF
  sheetTermo.getRange(linha, 3).setValue(dados.sn);           // Coluna C: S/N (S/N da peça)
  sheetTermo.getRange(linha, 4).setValue(dados.item);         // Coluna D: Nome do item
  sheetTermo.getRange(linha, 5).setValue(dados.incidente);    // Coluna E: Incidente
  sheetTermo.getRange(linha, 6).setValue(dados.ua);           // Coluna F: U.A.
  sheetTermo.getRange(linha, 9).setValue(dados.responsavel);  // Coluna I: Responsável

  // Atualiza "Manutenção Preventiva" em E32 e aplica negrito
  var cellE32 = sheetTermo.getRange("E26");
  cellE32.setValue("Manutenção Preventiva");
  cellE32.setFontWeight("bold");  // Aplica formatação em negrito

  // Realiza a baixa na aba "Peças"
  baixarPeca(dados.item);

  // Registrar a saída na aba "Saída_Peças", passando também o S/N
  registrarSaidaPecasManutencao(dados.item, dados.responsavel, dados.incidente, dados.sf, dados.sn);
}

// Função para realizar a baixa na aba "Peças"
function baixarPeca(item) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Peças');
  var data = sheet.getRange('A:C').getValues(); // Obtém dados das colunas A e C

  for (var i = 1; i < data.length; i++) { // Começa na linha 2 para ignorar o cabeçalho
    if (data[i][0] === item) {
      var quantidade = data[i][2];
      if (quantidade > 0) {
        sheet.getRange(i + 1, 3).setValue(quantidade - 1); // Subtrai 1 da coluna C
      } else {
        SpreadsheetApp.getUi().alert("Quantidade insuficiente para " + item);
      }
      break;
    }
  }
}

// Função para registrar a saída na aba "Saída_Peças"
function registrarSaidaPecasManutencao(item, responsavel, incidente, sfRelacionado, sn) {
  var sheetSaidaPecas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Saída_Peças');
  var dataAtual = new Date();
  var dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Adiciona as informações na aba "Saída_Peças"
  sheetSaidaPecas.appendRow([item, sn, 1, responsavel, incidente, sfRelacionado, dataFormatada]); 
  // Coluna A: item, Coluna B: S/N, Coluna C: quantidade (1), Coluna D: responsável, Coluna E: incidente, Coluna F: SF relacionado, Coluna G: data
}

//--------------------------------------------------------------------------------------------------------------------------------------PEÇASMANUTENÇÃO-----------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------SAÍDA PEÇAS------------------------------------------------------------


// Função para abrir o formulário de saída de peças
function abrirFormularioSaidaPecas() {
  var htmlForm = HtmlService.createHtmlOutputFromFile('Pecas')
    .setHeight(700)  // Defina a altura conforme necessário
    .setWidth(700);  // Defina a largura conforme necessário

  SpreadsheetApp.getUi().showModalDialog(htmlForm, 'Saída de Peças para Upgrade');
}

// Função para buscar as peças da aba "Peças"
function getPecasFromPlanilha() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Peças');
  var data = sheet.getRange('A2:C').getValues(); // Obtém os dados a partir da linha 2 (ignora o cabeçalho)
  var pecas = [];

  // Itera pelos dados para pegar nome da peça e quantidade
  for (var i = 0; i < data.length; i++) {
    var nomePeca = data[i][0]; // Coluna A: Nome do item
    var quantidade = data[i][2]; // Coluna C: Quantidade

    if (nomePeca && quantidade > 0) { // Adiciona apenas se a quantidade for maior que 0
      pecas.push({ nome: nomePeca, quantidade: quantidade });
    }
  }

  return pecas; // Retorna a lista de peças com nome e quantidade
}

// Função para processar o envio de peças e realizar a baixa
function enviarVariasPecasUpgrade(pecasSelecionadas, incidente, responsavel, ua, sfManual) {
  var linhaTermo = 11;

  // Limpa a aba "Termo" antes de começar a adicionar as peças
  limparTermo();

  pecasSelecionadas.forEach(function(peca) {
    // Monta os dados da peça para enviar ao Termo
    var dados = {
      sf: sfManual,
      item: peca.item,
      sn: peca.sn,
      incidente: incidente,
      ua: ua,
      responsavel: responsavel
    };

    // Mover a peça para o Termo na linha correspondente
    termoUpgrade(dados, linhaTermo);

    // Baixa a peça na aba "Peças" e registra a saída na aba "Saída_Peças"
    baixarPeca(dados.item);
    registrarSaidaPecas(dados.item, responsavel, incidente, sfManual, peca.sn);

    // Incrementa a linha para a próxima peça
    linhaTermo++;
  });
}

// Função para mover os dados da peça para a aba "Termo"
function termoUpgrade(dados, linha) {
  var sheetTermo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Termo');

  // Aloca os dados corretamente nas colunas da aba "Termo"
  sheetTermo.getRange(linha, 2).setValue(dados.sf);           // Coluna B: SF
  sheetTermo.getRange(linha, 3).setValue(dados.sn);           // Coluna C: S/N (S/N da peça)
  sheetTermo.getRange(linha, 4).setValue(dados.item);         // Coluna D: Nome do item
  sheetTermo.getRange(linha, 5).setValue(dados.incidente);    // Coluna E: Incidente
  sheetTermo.getRange(linha, 6).setValue(dados.ua);           // Coluna F: U.A.
  sheetTermo.getRange(linha, 9).setValue(dados.responsavel);  // Coluna I: Responsável

  // Atualiza "Upgrade" em E30 e aplica negrito
  var cellE30 = sheetTermo.getRange("E26");
  cellE30.setValue("Upgrade");
  cellE30.setFontWeight("bold");  // Aplica formatação em negrito
}

// Função para baixar a peça na aba "Peças"
function baixarPeca(item) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Peças');
  var data = sheet.getRange('A2:C').getValues(); // Obtém dados a partir da linha 2 (ignora o cabeçalho)

  for (var i = 0; i < data.length; i++) { // Itera sobre as linhas (sem o cabeçalho)
    if (data[i][0] === item) { // Verifica se o nome da peça corresponde
      var quantidade = data[i][2]; // Coluna C: Quantidade
      if (quantidade > 0) {
        sheet.getRange(i + 2, 3).setValue(quantidade - 1); // Subtrai 1 da coluna C, i+2 pois estamos usando A2:C no getRange
      } else {
        SpreadsheetApp.getUi().alert("Quantidade insuficiente para " + item);
      }
      break; // Para o loop após encontrar a peça
    }
  }
}

// Função para registrar a saída da peça na aba "Saída_Peças"
function registrarSaidaPecas(item, responsavel, incidente, sfManual, sn) {
  var sheetSaidaPecas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Saída_Peças');
  var dataAtual = new Date();
  var dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Adiciona as informações na aba "Saída_Peças"
  sheetSaidaPecas.appendRow([item, sn, 1, responsavel, incidente, sfManual, dataFormatada]); 
  // Coluna A: item, Coluna B: S/N, Coluna C: quantidade (1), Coluna D: responsável, Coluna E: incidente, Coluna F: SF relacionado, Coluna G: data
}

//------------------------------------------------------------------------------------------------------------------------SAÍDA PEÇAS------------------------------------------------------------


//---------------------------------------------------------------------------------------------------------------------------RETORNO MANUTENÇÃO-------------------------------------------------------

// Função para abrir a janela de confirmação do retorno da manutenção
function abrirRetornoManutencao() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('RetornoManutencao')
    .setWidth(400)
    .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Retorno da Manutenção');
}

// Função para alterar o status dos equipamentos para "Em estoque" nas abas especificadas
function atualizarStatusRetornoManutencao() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaManutencao = planilha.getSheetByName('Manutenção');
  var patrimoniosManutencao = abaManutencao.getRange('B10:B42').getValues(); // Pega os SFs da coluna B da aba Manutenção
  var abasPermitidas = ['D570', 'D580', 'Desktop Dell', 'Notebook Daten', 'Notebook HP', 'Notebook Dell'];

  // Loop pelos patrimônios da aba Manutenção
  for (var i = 0; i < patrimoniosManutencao.length; i++) {
    var sf = patrimoniosManutencao[i][0]; // Patrimônio (SF) da coluna B

    if (sf) { // Verifica se o SF não está vazio
      // Loop pelas abas permitidas
      for (var j = 0; j < abasPermitidas.length; j++) {
        var aba = planilha.getSheetByName(abasPermitidas[j]);
        var dados = aba.getRange('A:J').getValues(); // Pega os dados das colunas A até J

        // Loop pelos dados para verificar se o SF é encontrado
        for (var k = 1; k < dados.length; k++) { // Começa na linha 2, ignorando o cabeçalho
          if (dados[k][4] === sf) { // Se o SF for encontrado na coluna E
            aba.getRange(k + 1, 1).setValue('Em estoque'); // Atualiza a coluna A (status) para "Em estoque"
            break; // Sai do loop se o SF foi encontrado
          }
        }
      }
    }
  }

  // Limpa o intervalo B10:D42 da aba Manutenção
  abaManutencao.getRange('B10:D42').clearContent();

  SpreadsheetApp.getUi().alert('Status atualizado para "Em estoque" para todos os SFs conferidos.');
}



//---------------------------------------------------------------------------------------------------------------------------RETORNO MANUTENÇÃO-------------------------------------------------------


//------------------------------------------------------------IMPRIMIR TERMO------------------------------------------------------------

function imprimirTermo() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaTermo = planilha.getSheetByName("Termo");
  var intervalo = abaTermo.getRange("B2:K42");

  // Define as opções de impressão
  var options = {
    size: 7,                  // A4
    portrait: false,           // Modo paisagem
    fitw: true,               // Ajustar ao tamanho da página
    gridlines: false,         // Sem linhas de grade
    printtitle: false,
    sheetnames: false,
    pagenum: 'UNDEFINED',
    attachment: false
  };

  // Constrói a URL para impressão
  var sheetId = abaTermo.getSheetId();
  var url = planilha.getUrl().replace(/edit$/, '') + 'export?format=pdf' +
            '&gid=' + sheetId +
            '&range=' + intervalo.getA1Notation() +
            '&size=' + options.size +
            '&portrait=' + options.portrait +
            '&fitw=' + options.fitw +
            '&gridlines=' + options.gridlines +
            '&printtitle=' + options.printtitle +
            '&sheetnames=' + options.sheetnames +
            '&pagenum=' + options.pagenum +
            '&attachment=' + options.attachment;

  // Abre a janela de impressão
  var htmlOutput = HtmlService.createHtmlOutput('<html><script>window.open("' + url + '");google.script.host.close();</script></html>')
    .setWidth(100)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Imprimir Termo');
}




//------------------------------------------------------------IMPRIMIR TERMO------------------------------------------------------------




