function doGet(e) {

  const contaId = e.parameter.conta;

  if (contaId) {
    const template = HtmlService.createTemplateFromFile("Extrato");
    template.contaId = contaId;

    return template.evaluate()
      .setTitle("Extrato - " + contaId.toUpperCase())
  } else {
    return HtmlService.createHtmlOutputFromFile("Menu")
      .setTitle("Menu Principal")
  }

}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function cadastrarMovimentacao(dadosForm) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movimentacoes = ss.getSheetByName("movimentacoes");

  let totalRepeticoes = 1;
  const idVInculoGrupo = "GRP-" + new Date().getTime();

  if (dadosForm.repetir === "parcelado"){
    totalRepeticoes = parseInt(dadosForm.parcelas);
  } else if (dadosForm.repetir === "fixa"){
    totalRepeticoes = 12;
  };

  const partesData = dadosForm.dataVencimento.split('/');
  let dataBase = new Date(partesData[2], partesData[1] - 1, partesData[0]);

  let fatura = '';
  if (dadosForm.conta.toString().toLowerCase().startsWith('cartao')){
    const dia = partesData[0];
    const abaCartoes = ss.getSheetByName('cartoes');
    const dadosCartoes = abaCartoes.getDataRange().getValues();

    const linhaCartao = dadosCartoes.find(linha => linha[0] === dadosForm.conta);

    let diaFechamento = 25;

    if (linhaCartao) {
      diaFechamento = linhaCartao[2];
    }

    if(parseInt(dia) >= diaFechamento){
      dataBase.setMonth(dataBase.getMonth() + 1);
    };
  };

  for (let i=0; i < totalRepeticoes; i++){    
    let novaData = new Date(dataBase);

    let mesAlvo = dataBase.getMonth() + i;

    novaData.setMonth(mesAlvo);

    if (novaData.getDate() !== dataBase.getDate()){
      novaData.setDate(0);
    };

    let descricaoFinal = dadosForm.descricao;
    if (dadosForm.repetir === "parcelado"){
      descricaoFinal = `${dadosForm.descricao} (${i + 1}/${totalRepeticoes})`;
    }

    let dataFormatada = Utilities.formatDate(novaData, "GMT-3", "dd/MM/yyyy");

    if (dadosForm.conta.toString().toLowerCase().startsWith('cartao')){
      fatura = Utilities.formatDate(novaData, "GMT-3", "MM/yyyy" );
    }

    movimentacoes.appendRow([
      dataFormatada,
      descricaoFinal,
      dadosForm.categoria,
      dadosForm.valor,
      dadosForm.conta,
      dadosForm.tipo,
      dadosForm.status,
      dadosForm.dataPagamento,
      dadosForm.repetir,
      fatura,
      idVInculoGrupo,
    ]);
  }

  return true;
}

function filtrarExtratoPorConta(contaId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("movimentacoes");
  const dados = aba.getDataRange().getValues();
  const headers = dados[0];

  const contaIdBusca = contaId.toString().trim().toLowerCase();
  let resultado = [];

  for (let i = 1; i < dados.length; i++) {
    const linhaAtual = dados[i];
    const contaBate = linhaAtual[4].toString().trim().toLowerCase();

    if (contaBate === contaIdBusca) {
      let obj = {};
      headers.forEach((header, index) => {
        let val = linhaAtual[index];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT-3", "dd/MM/yyyy");
        }
        obj[header] = val;
      });

      obj.linha = i + 1; 

      resultado.push(obj);
    }
  }
  
  return resultado;
}

function buscarDadosIniciais(contaId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resumo = ss.getSheetByName("resumo");
  
  const movimentacoes = filtrarExtratoPorConta(contaId);

  const dadosResumo = resumo.getRange("A2:B7").getValues();
  let saldoInicial = 0;

  for (let i = 0; i < dadosResumo.length; i++) {
    if (dadosResumo[i][0].toString().toLowerCase().trim() === contaId.toLowerCase().trim()) {
      saldoInicial = dadosResumo[i][1];
      break;
    }
  }

  return {
    extrato: movimentacoes,
    saldo: saldoInicial
  };
}

function buscarDadosIniciaisMenu(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resumo = ss.getSheetByName("resumo");
  const saldos = resumo.getDataRange().getValues();
  const saldoInter = saldos[1][1];
  const saldoNubank = saldos[2][1];
  const saldoPicpay = saldos[3][1];
  const saldoPluxeeAlimentacao = saldos[4][1];
  const saldoPluxeeRefeicao = saldos[5][1];
  const saldoClear = saldos[6][1];

  return{
    saldoInter: saldoInter,
    saldoNubank: saldoNubank,
    saldoPicpay: saldoPicpay,
    saldoPluxeeAlimentacao: saldoPluxeeAlimentacao,
    saldoPluxeeRefeicao: saldoPluxeeRefeicao,
    saldoClear: saldoClear,
  }
}

function gatilhoMensalContasFixas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("movimentacoes");
  const dados = aba.getDataRange().getValues();
  
  const hoje = new Date();
  const limiteFuturo = new Date();
  limiteFuturo.setMonth(hoje.getMonth() + 12);

  // 1. Mapear as contas fixas e encontrar a última data de cada uma
  let ultimosLancamentos = {};
  let dataItem = {};

  for (let i = 1; i < dados.length; i++) {
    const [dataStr, descricao, , , , , , , tipoRepeticao] = dados[i];
    
    if (tipoRepeticao === "fixa") {
      let dataItem;

      if (dataStr instanceof Date){
        dataItem = dataStr;
      } else {
        const partes = dataStr.split("/");
        const dataItem = new Date(partes[2], partes[1] - 1, partes[0]);
      }

      if (!ultimosLancamentos[descricao] || dataItem > ultimosLancamentos[descricao].data) {
        ultimosLancamentos[descricao] = {
          data: dataItem,
          linhaOriginal: dados[i]
        };
      }
    }
  }

  // 2. Verificar quem precisa de novos lançamentos
  for (let desc in ultimosLancamentos) {
    let ultimaData = ultimosLancamentos[desc].data;
    let dadosOriginais = ultimosLancamentos[desc].linhaOriginal;

    
    while (ultimaData < limiteFuturo) {
      // Adiciona 1 mês à data
      ultimaData.setMonth(ultimaData.getMonth() + 1);
      
      // Correção do problema do dia 31 (Month Rollover)
      if (ultimaData.getDate() !== ultimosLancamentos[desc].data.getDate()) {
        ultimaData.setDate(0);
      }

      let novadataStr = Utilities.formatDate(ultimaData, "GMT-3", "dd/MM/yyyy");

      
      aba.appendRow([
        novadataStr,
        dadosOriginais[1], 
        dadosOriginais[2], 
        dadosOriginais[3],
        dadosOriginais[4],
        dadosOriginais[5],
        "pendente",        
        "",                
        "fixa"             
      ]);
    }
  }
}

function cadastrarTransferencia(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movimentacoes = ss.getSheetByName('movimentacoes');

  const idVinculo = "TR-" + new Date().getTime();  

  let totalRepeticoes = 1;

  if (dados.recorrenciaTransferencia === "parcelado"){
    totalRepeticoes = dados.quantidadeParcelasTransferencia;
  } else if (dados.recorrenciaTransferencia === "fixa"){
    totalRepeticoes = 12;
  };

  const partesData = dados.dataTransferencia.split('/');
  const dataBase = new Date(partesData[2], partesData[1] - 1, partesData[0]);

  for (let i=0; i < totalRepeticoes; i++){
    let novaData = new Date(dataBase);

    let mesAlvo = dataBase.getMonth() + i;

    novaData.setMonth(mesAlvo);

    let descricaoFinal = dados.descricaoTransferencia
    if (dados.recorrenciaTransferencia === "parcelado"){
      descricaoFinal = `${dados.descricaoTransferencia} (${i + 1}/${totalRepeticoes})`;
    }

    let dataFormatada = Utilities.formatDate(novaData, "GMT-3", "dd/MM/yyyy");

    //saida da transferencia
    movimentacoes.appendRow([
      dataFormatada,
      descricaoFinal,
      dados.categoriaTransferencia,
      -Math.abs(dados.valorTransferencia),
      dados.contaOrigem,
      "saida",
      dados.statusTransferencia,
      dados.dataEfetivaTransferencia,
      dados.recorrenciaTransferencia,
      "",
      idVinculo,
    ]);

    //entrada da transferencia
    movimentacoes.appendRow([
      dataFormatada,
      descricaoFinal,
      dados.categoriaTransferencia,
      dados.valorTransferencia,
      dados.contaDestino,
      "entrada",
      dados.statusTransferencia,
      dados.dataEfetivaTransferencia,
      dados.recorrenciaTransferencia,
      "",
      idVinculo,
    ]);
  }

  return true;
}

function atualizarMovimentacao(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("movimentacoes");

  
  if (!dados.editarTodos) {
    const idExistente = aba.getRange(dados.linha, 11).getValue();
    
    const novaLinha = [
      dados.dataVencimento, 
      dados.descricao, 
      dados.categoria,
      dados.valor, 
      dados.conta, 
      dados.tipo, 
      dados.status,
      dados.dataPagamento, 
      dados.repetir, 
      "", 
      idExistente
    ];
    aba.getRange(dados.linha, 1, 1, novaLinha.length).setValues([novaLinha]);
    return true;
  }

  const partesData = dados.dataVencimento.split('/');
  const novoDia = parseInt(partesData[0]);

  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    if (valores[i][10] === dados.idVinculo && valores[i][6] !== "pago") {
      const linhaPlanilha = i + 1;

      let dataOriginalLinha = valores[i][0];
      if (!(dataOriginalLinha instanceof Date)) {
        const p = dataOriginalLinha.split('/');
        dataOriginalLinha = new Date(p[2], p[1] - 1, p[0]);
      }

      // CRIAR A NOVA DATA mantendo o Mês/Ano da linha, mas alterando o DIA
      let novaDataParaEstaLinha = new Date(dataOriginalLinha.getFullYear(), dataOriginalLinha.getMonth(), novoDia);
      
      // Formata para DD/MM/YYYY
      let dataFormatada = Utilities.formatDate(novaDataParaEstaLinha, "GMT-3", "dd/MM/yyyy");

      aba.getRange(linhaPlanilha, 1).setValue(dataFormatada);
      aba.getRange(linhaPlanilha, 3).setValue(dados.categoria);
      aba.getRange(linhaPlanilha, 4).setValue(dados.valor);
      aba.getRange(linhaPlanilha, 6).setValue(dados.tipo);
      aba.getRange(linhaPlanilha, 7).setValue(dados.dataPagamento);
      aba.getRange(linhaPlanilha, 9).setValue(dados.repetir);

      }
  }
  return true;
}

function excluirMovimentacao(linha, excluirTudo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("movimentacoes");
  
  
  const idVinculo = aba.getRange(linha, 11).getValue();

  if (excluirTudo && idVinculo) {
    // EXCLUSÃO EM MASSA
    const valores = aba.getDataRange().getValues();
    // Deletamos de baixo para cima para não bagunçar os índices das linhas
    for (let i = valores.length - 1; i >= 1; i--) {
      // Verifica se o ID bate (coluna 11 é índice 10)
      if (valores[i][10] === idVinculo) {
        aba.deleteRow(i + 1);
      }
    }
  } else {
    // EXCLUSÃO INDIVIDUAL
    aba.deleteRow(linha);
  }
  
  return true;
}
