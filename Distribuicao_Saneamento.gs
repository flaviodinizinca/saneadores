/**
 * Distribuicao_Saneamento.gs
 * Busca dados da ToFor (Controle) e distribui AQUI se o usuário for Saneador.
 */
function executarDistribuicaoSaneamento() {
  const ssSaneamento = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. CONFIGURAÇÕES
  // ID da Planilha de Controle (Onde está a ToFor)
  const ID_PLANILHA_CONTROLE = "1n6l2ofxEvQTrZ49IY7b30U_dcUqb-MuAbVaW890S6ng"; 
  // ID da Planilha de Usuários (Para pegar nomes)
  const ID_PLANILHA_USUARIOS = "1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA";

  // 2. Carrega Lista de Saneadores (Desta planilha, guia Config_Saneamento)
  const guiaConfig = ssSaneamento.getSheetByName("Config_Saneamento");
  if (!guiaConfig) {
    SpreadsheetApp.getUi().alert("Erro: Crie a guia 'Config_Saneamento' com os logins na Coluna A.");
    return;
  }
  const listaSaneadores = guiaConfig.getRange("A2:A").getValues().flat().map(String).filter(String);

  // 3. Acessa ToFor na Planilha de Controle
  const ssControle = SpreadsheetApp.openById(ID_PLANILHA_CONTROLE);
  const guiaToFor = ssControle.getSheetByName("ToFor");
  const dadosToFor = guiaToFor.getDataRange().getValues().slice(1); // Remove cabeçalho

  // 4. Mapeia Logins -> Nomes (Planilha Usuários)
  const ssUsers = SpreadsheetApp.openById(ID_PLANILHA_USUARIOS);
  const dadosUsers = ssUsers.getSheetByName("User_SEI").getDataRange().getValues();
  const mapaNomes = {};
  dadosUsers.slice(1).forEach(r => mapaNomes[r[1]] = r[0]); // Login -> Nome

  let contador = 0;

  // 5. Distribui
  dadosToFor.forEach(linha => {
    const processo = linha[0];
    const login = String(linha[1]).trim();
    const especificacao = linha[3];

    // Verifica se o login está na lista de Config_Saneamento
    if (listaSaneadores.includes(login)) {
      const nomeCompleto = mapaNomes[login] || login;
      // Formata nome (Primeiro Nome)
      const nomeGuia = nomeCompleto.split(" ")[0]; 

      let guiaDestino = criarGuiaSaneador(nomeGuia); // Cria ou pega existente

      // Monta linha
      const novaLinha = [
        processo, 
        new Date(), // Data Chegada
        "", "", 
        especificacao, // Objeto
        "", "", "", 
        "NÃO", "", "A Iniciar"
      ];
      
      guiaDestino.appendRow(novaLinha);
      contador++;
    }
  });

  SpreadsheetApp.getUi().alert(`Importação Concluída! ${contador} processos distribuídos para Saneamento.`);
}