/**
 * 01_Menu_Saneamento.gs
 * Menu exclusivo da Planilha de Saneamento.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Saneamento')
    .addItem('➕ Nova Guia (Saneador)', 'acionarNovaGuiaSaneador')
    .addSeparator()
    .addItem('📥 Buscar Processos (ToFor)', 'executarDistribuicaoSaneamento')
    .addToUi();
}

function acionarNovaGuiaSaneador() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Novo Saneador', 'Digite o nome:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    criarGuiaSaneador(response.getResponseText().trim());
  }
}