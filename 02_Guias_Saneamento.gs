/**
 * 02_Guias_Saneamento.gs
 * Estrutura específica para Saneamento (11 colunas).
 */
function criarGuiaSaneador(nomeGuia) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!nomeGuia) return;

  if (ss.getSheetByName(nomeGuia)) return ss.getSheetByName(nomeGuia);

  const novaGuia = ss.insertSheet(nomeGuia);
  
  const cabecalhos = [
    "PROCESSO", "Data de Chegada", "PROTOCOLO", "PARECER/ NOTA/ COTA", "OBJETO",
    "CÉLULA", "MODALIDADE", "DATA DO STATUS", "SANEAMENTO ENCERRADO?", "LOCALIZAÇÃO", "STATUS"
  ];

  const rangeCabecalho = novaGuia.getRange(1, 1, 1, cabecalhos.length);
  rangeCabecalho.setValues([cabecalhos]);
  rangeCabecalho.setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  
  // Estilo Visual
  novaGuia.setRowHeight(1, 45);
  novaGuia.getRange(1, 1, 1, 11).setBackground("#FCE5CD"); // Laranja Claro
  novaGuia.getRange(1, 1, 1, 1).setBackground("#E69138").setFontColor("white"); // Destaque Processo
  novaGuia.getRange(1, 9, 1, 1).setBackground("#EA9999"); // Encerrado?

  novaGuia.setFrozenRows(1);
  novaGuia.setFrozenColumns(2);

  // Validações
  const regraSimNao = SpreadsheetApp.newDataValidation().requireValueInList(["SIM", "NÃO"], true).build();
  novaGuia.getRange(2, 9, 999, 1).setDataValidation(regraSimNao);

  const regraData = SpreadsheetApp.newDataValidation().requireDate().build();
  novaGuia.getRange(2, 2, 999, 1).setDataValidation(regraData);
  novaGuia.getRange(2, 8, 999, 1).setDataValidation(regraData);

  novaGuia.autoResizeColumns(1, cabecalhos.length);
  novaGuia.setColumnWidth(5, 250); // Objeto maior
  
  return novaGuia;
}