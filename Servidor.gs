/**
 * CONFIGURAÇÕES GERAIS E NAVEGAÇÃO
 */
const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA";

function doGet(e) {
  // Se o seu arquivo na esquerda for "index.html", coloque "index" abaixo:
  return HtmlService.createTemplateFromFile('index') 
      .evaluate()
      .setTitle('Sistema Clínico - Psicologia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAba(nome){
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(nome);
  if(!aba) throw new Error("Aba não encontrada: " + nome);
  return aba;
}

function uid(){ return Utilities.getUuid(); }

/**
 * MÓDULO: DASHBOARD (NOVO)
 * Integração com js_dashboard e processamento de indicadores
 */

function carregarDashboard() {
  return HtmlService.createHtmlOutputFromFile('js_dashboard').getContent();
}

// --- FUNÇÃO PARA PEGAR DADOS REAIS PARA O DASHBOARD ---
function getDadosDash() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const abaF = ss.getSheetByName("FINANCEIRO");
  const abaP = ss.getSheetByName("PRONTUARIO");
  const abaS = ss.getSheetByName("SESSOES");
  const abaPacientes = ss.getSheetByName("PACIENTES");
  
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  
  // 1. FATURAMENTO REAL
  let faturamento = 0;
  let pendente = 0;
  let valoresMensais = new Array(12).fill(0);
  const dadosF = abaF.getDataRange().getValues();
  
  for(let i = 1; i < dadosF.length; i++){
    let dt = dadosF[i][1];
    if(dt instanceof Date){
      let m = dt.getMonth();
      if(dadosF[i][3] === "Receita"){
        valoresMensais[m] += parseFloat(dadosF[i][4]) || 0;
        if(m === mesAtual) faturamento += parseFloat(dadosF[i][4]) || 0;
      } 
    }
  }

  // 2. SESSÕES REALIZADAS
  const sessoeRealizadas = abaP.getLastRow() - 1;

  // 3. CONTAGEM SEPARADA POR STATUS (Coluna K - Índice 10)
  let ativosCount = 0;
  let inativosCount = 0;
  let altaCount = 0;

  if (abaPacientes) {
    const dadosPacientes = abaPacientes.getDataRange().getValues();
    for (let i = 1; i < dadosPacientes.length; i++) {
      if (dadosPacientes[i].length > 10) {
        let status = String(dadosPacientes[i][10]).trim().toUpperCase();
        
        if (status === "ATIVO") {
          ativosCount++;
        } else if (status === "INATIVO") {
          inativosCount++;
        } else if (status === "ALTA" || status === "DE ALTA") {
          altaCount++;
        }
      }
    }
  }

  // 4. PACIENTES HOJE
  const hojeF = Utilities.formatDate(hoje, "America/Sao_Paulo", "dd/MM/yyyy");
  const dadosS = abaS.getDataRange().getValues();
  let hojeCount = 0;
  for(let i = 1; i < dadosS.length; i++){
    let dF = dadosS[i][2] instanceof Date ? Utilities.formatDate(dadosS[i][2], "America/Sao_Paulo", "dd/MM/yyyy") : String(dadosS[i][2]);
    if(dF === hojeF) hojeCount++;
  }

  // 5. RETORNO DOS DADOS ESTRUTURADOS
  return {
    cards: {
      faturamento: "R$ " + faturamento.toLocaleString('pt-BR', {minimumFractionDigits:2}),
      pendente: "R$ " + pendente.toLocaleString('pt-BR', {minimumFractionDigits:2}),
      realizadas: sessoeRealizadas.toString(),
      ativos:  ativosCount.toString(),
      inativos: inativosCount.toString(),
      alta: altaCount.toString(),
      hoje: hojeCount.toString()
    },
    grafico: {
      labels: ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"],
      valores: valoresMensais
    },
    aniversariantes: []
  };
}

/**
 * MÓDULO: LOGIN E SEGURANÇA
 */
function login(email, senha) {
  try {
    const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA";
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("USUÁRIOS"); // Verifique se o nome está idêntico na planilha
    const dados = aba.getDataRange().getValues();

    const emailDigitado = String(email).trim().toLowerCase();
    const senhaDigitada = String(senha).trim();

    for (let i = 1; i < dados.length; i++) {
      const emailPlanilha = String(dados[i][2]).trim().toLowerCase(); // Coluna C
      const senhaPlanilha = String(dados[i][3]).trim(); // Coluna D
      
      if (emailPlanilha === emailDigitado && senhaPlanilha === senhaDigitada) {
        return {
          sucesso: true,
          nome: String(dados[i][1]), // Coluna B
          usuario: emailPlanilha,
          perfil: String(dados[i][4])  // Coluna E
        };
      }
    }
    return { sucesso: false };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}


function verificarLogin() {
  const sessao = PropertiesService.getUserProperties();
  const user = sessao.getProperty('usuarioLogado');
  
  if (user) {
    // Se achar o e-mail na memória técnica, devolve os dados para o js_main abrir o sistema
    return {
      sucesso: true,
      nome: sessao.getProperty('nomeLogado'),
      usuario: user,
      perfil: sessao.getProperty('perfilLogado')
    };
  }
  return { sucesso: false };
}

function logout() {
  PropertiesService.getUserProperties().deleteAllProperties();
  return true;
}

function alterarSenhaNoServidor(email, novaSenha) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("USUÁRIOS"); 
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    
    // Converte o e-mail digitado para minúsculas e remove espaços soltos
    const emailBuscado = email.toString().trim().toLowerCase();
    
    if (emailBuscado === "") return false;

    // Varre a planilha à procura do e-mail (Coluna C é o índice 2)
    for (let i = 1; i < data.length; i++) {
      if (!data[i][2]) continue; // Se a célula estiver vazia, pula para a próxima
      
      // Limpa totalmente o e-mail da planilha para a comparação ser perfeita
      let emailPlanilha = data[i][2].toString().trim().toLowerCase();
      
      if (emailPlanilha === emailBuscado) {
        // Altera a senha na Coluna D (coluna 4) na linha correta (i + 1)
        sheet.getRange(i + 1, 4).setValue(novaSenha.toString().trim()); 
        SpreadsheetApp.flush(); // Força o Google Sheets a gravar imediatamente
        return true; // Sucesso! Retorna para o HTML
      }
    }
    
    return false; // Se percorrer tudo e não encontrar o e-mail
  } catch (erro) {
    Logger.log("Erro ao redefinir senha: " + erro.toString());
    return false;
  }
}

//* Função para salvar Pacientes e Editar

function salvarNovoPaciente(dados) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PACIENTES");
    const valores = aba.getDataRange().getValues();
    
    let linhaDestino = -1;
    let idFinal = dados.id;

    // 1. Verifica se é EDIÇÃO (procura o UUID na coluna A)
    if (idFinal && idFinal.trim() !== "") {
      for (let i = 0; i < valores.length; i++) {
        if (String(valores[i][0]).trim() === String(idFinal).trim()) {
          linhaDestino = i + 1;
          break;
        }
      }
    }

    // 2. Se for NOVO cadastro (não achou ID), gera o UUID
    if (linhaDestino === -1) {
      linhaDestino = aba.getLastRow() + 1;
      idFinal = Utilities.getUuid(); // <--- AQUI GERA O UUID PROFISSIONAL
    }

    // 3. Monta a linha exatamente como na sua planilha (A até M)
    const linhaDados = [
      idFinal,               // A - ID (UUID)
      dados.nome,            // B - Nome
      dados.rg || "",        // C - RG
      dados.cpf,             // D - CPF
      dados.telefone,        // E - Telefone
      dados.email || "",     // F - Email
      dados.nascimento,      // G - Data Nascimento
      dados.sexo,            // H - Sexo
      dados.endereco,        // I - Endereço
      dados.valor,           // J - Valor Sessão
      dados.status || "Ativo", // K - Status
      dados.obs || "",       // L - Observação
      dados.responsavel || "" // M - Responsável
    ];

    // 4. Grava na planilha
    aba.getRange(linhaDestino, 1, 1, linhaDados.length).setValues([linhaDados]);

    // 5. Registra o Log
    salvarNoBancoDeLogs({
      modulo: "PACIENTES",
      acao: idFinal === dados.id ? "EDIÇÃO" : "CRIAÇÃO",
      detalhes: "Paciente: " + dados.nome + " | ID: " + idFinal,
      usuario: "ADM"
    });

    return { sucesso: true, id: idFinal };

  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

//* Função para excluir o Paciente

function excluirPacienteSistema(id, nome) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PACIENTES");
    const dados = aba.getDataRange().getValues();
    
    for (let i = 0; i < dados.length; i++) {
      // Procura o UUID na Coluna A
      if (String(dados[i][0]).trim() === String(id).trim()) {
        aba.deleteRow(i + 1);
        
        // Registra a exclusão no Log
        salvarNoBancoDeLogs({
          modulo: "PACIENTES",
          acao: "EXCLUSÃO",
          detalhes: "Paciente removido: " + nome + " | ID: " + id,
          usuario: "ADM"
        });
        
        return { sucesso: true };
      }
    }
    return { sucesso: false, erro: "Paciente não encontrado." };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

/**
 * MÓDULO: AGENDA / SESSÕES
 */
function getSessoesCalendario() {
  try {
    const aba = getAba("SESSOES");
    // Voltamos a ler estritamente o TEXTO VISÍVEL na tela (Ex: "16:00", "18:30")
    // Isso ignora qualquer fuso horário invisível do objeto Date do Google
    const dados = aba.getDataRange().getDisplayValues(); 
    const eventos = [];
    
    for (let i = 1; i < dados.length; i++) {
      const r = dados[i];
      if (!r[0] || !r[1] || !r[2]) continue;
      
      let dataStr = String(r[2]).trim(); 
      let horaStr = String(r[3]).trim(); 
      
      // 1. AJUSTE DA DATA (Se estiver no padrão brasileiro DD/MM/YYYY)
      if (dataStr.includes('/')) {
        const partes = dataStr.split('/');
        dataStr = partes[2] + '-' + partes[1] + '-' + partes[0];
      } else if (dataStr.includes('T')) {
        dataStr = dataStr.split('T')[0];
      }
      
      // 2. EXTRAÇÃO E PURIFICAÇÃO DO TEXTO DA HORA
      if (horaStr.includes(' ')) {
        // Se vier algo como "30/12/1899 16:00:00", pega apenas a parte do horário
        const partesEspaco = horaStr.split(' ');
        horaStr = partesEspaco[partesEspaco.length - 1];
      }
      
      // Agora garantimos que temos apenas "HH:MM", ignorando os segundos se existirem
      if (horaStr.includes(':')) {
        const partesHora = horaStr.split(':');
        let h = partesHora[0].trim().padStart(2, '0');
        let m = partesHora[1].trim().padStart(2, '0');
        horaStr = h + ":" + m;
      } else if (horaStr !== "" && !isNaN(horaStr)) {
        // Caso a célula mostre apenas o número "16"
        horaStr = horaStr.padStart(2, '0') + ":00";
      } else {
        // Se estiver vazio ou inválido, define um padrão para não quebrar a agenda
        horaStr = "00:00";
      }
      
      let status = String(r[5]).trim().toUpperCase();
      let cor = status === "REALIZADO" || status === "REALIZADA" ? "#16a34a" : (status === "CANCELADO" ? "#dc2626" : "#3b82f6");
      
      // Envia para o FullCalendar juntando os textos de forma literal
      eventos.push({ 
        id: String(r[0]), 
        title: String(r[1]), 
        start: dataStr + "T" + horaStr + ":00", 
        color: cor 
      });
    }
    return eventos;
  } catch (e) { 
    console.error("Erro na agenda: " + e.message);
    return []; 
  }
}


/**
 * MÓDULO: PRONTUÁRIOS
 */

function salvarEvolucao(d) {
  try {
    const ss = SpreadsheetApp.openById("COLE_AQUI_O_ID_DA_SUA_PLANILHA");
    const aba = ss.getSheetByName("PRONTUARIO");
    
    // d.texto aqui deve vir do formulário já formatado ou ser enviado em partes
    // Vamos salvar seguindo sua ordem: ID, Paciente, Data, Evolucao, Humor, Tema...
    aba.appendRow([
      Utilities.getUuid(), // A: ID
      d.paciente,          // B: Paciente
      new Date(),          // C: Data
      d.texto,             // D: Evolucao
      "",                  // E: Humor (vazio ou pegue do objeto d se enviar)
      "Sessão"             // F: Tema
    ]);
    return true;
  } catch (e) {
    throw new Error(e.message);
  }
}


function getSessaoPorId(idEnviado) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("SESSOES");
    const dados = aba.getDataRange().getDisplayValues(); // Importante: getDisplayValues pega o texto visível
    
    // Limpa o ID de qualquer aspa ou espaço
    const idBusca = String(idEnviado).replace(/['"]+/g, '').trim().toLowerCase();

    for (var i = 1; i < dados.length; i++) {
      var idPlanilha = String(dados[i][0]).trim().toLowerCase();
      
      if (idPlanilha === idBusca) {
        var r = dados[i];
        
        // Converte a data DD/MM/YYYY da planilha para YYYY-MM-DD do formulário
        var dOriginal = r[2]; 
        var dFormatada = dOriginal;
        if (dOriginal.includes('/')) {
           var partes = dOriginal.split('/');
           dFormatada = partes[2] + '-' + partes[1] + '-' + partes[0];
        }

        return {
          id: r[0],
          paciente: r[1],
          data: dFormatada,
          hora: r[3],
          tipo: r[4],
          status: r[5],
          valor: r[6],
          pagamento: r[7]
        };
      }
    }
    return null; // Se chegar aqui, não achou
  } catch (e) {
    throw new Error("Erro no servidor: " + e.message);
  }
}

/**
 * SALVAR SESSÃO (CORREÇÃO DE ACESSO)
 */
function salvarSessaoCompleta(dados) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("SESSOES");
    const data = aba.getDataRange().getValues();
    
    // Garante que a hora seja apenas texto limpo (ex: "14:00")
    let horaLimpa = dados.hora ? String(dados.hora).trim().substring(0, 5) : "00:00";
    
    const linhaParaSalvar = [
      dados.id || Utilities.getUuid(),
      dados.paciente,
      String(dados.data).trim(), // Salva estritamente como texto plano
      horaLimpa,                 // Salva estritamente como texto plano
      dados.tipo || "Consulta",
      dados.status,
      dados.valor,
      dados.pagamento || "Pendente" 
    ];

    if (dados.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(dados.id).trim()) {
          aba.getRange(i + 1, 1, 1, 8).setValues([linhaParaSalvar]);
          return "Sessão updated!";
        }
      }
    }
    
    aba.appendRow(linhaParaSalvar);
    return "Novo agendamento salvo!";
  } catch (e) {
    throw new Error("Falha ao salvar: " + e.message);
  }
}

/**
 * Busca os dados de um paciente específico pelo ID para edição
 */
function getPacientePorId(idEnviado) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PACIENTES");
    const dados = aba.getDataRange().getValues();
    
    const idBusca = String(idEnviado).trim();

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === idBusca) {
        const r = dados[i];
        
        // Formata a data para o input do HTML (yyyy-MM-dd)
        let dataNasc = "";
        if (r[6] instanceof Date) {
          dataNasc = Utilities.formatDate(r[6], "GMT-3", "yyyy-MM-dd");
        }

        // MAPEAMENTO REAL BASEADO NA TUA FOTO:
        return {
          id: r[0],          // A
          nome: r[1],        // B
          rg: r[2],          // C
          cpf: r[3],         // D
          telefone: r[4],    // E
          email: r[5],       // F
          nascimento: dataNasc, // G (Índice 6)
          sexo: r[7],        // H
          endereco: r[8],    // I
          valorSessao: r[9], // J
          status: r[10],     // K
          obs: r[11],        // L
          responsavel: r[12] // M (Índice 12)
        };
      }
    }
    return null;
  } catch (e) {
    throw new Error("Erro no servidor: " + e.toString());
  }
}

function buscarPorId(nomeAba, id) {
  const aba = getAba(nomeAba);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) return dados[i];
  }
  return null;
}

function buscarDadosPacienteParaFiscal(nomePaciente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPacientes = ss.getSheetByName("PACIENTES");
    const dados = sheetPacientes.getDataRange().getValues();
    
    // Procura o paciente na coluna B (índice 1) e pega o CPF na coluna C (índice 2)
    // Ajuste os índices [1] e [2] se suas colunas forem diferentes
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][1] === nomePaciente) {
        return {
          cpf: dados[i][2] || "Não cadastrado",
          valorPadrao: dados[i][5] || "" // Ex: busca valor da sessão se houver
        };
      }
    }
    return { cpf: "Não encontrado", valorPadrao: "" };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * BUSCA O HISTÓRICO REAL DA ABA PRONTUARIO
 */
function buscarHistorico(nomePaciente) {
  try {
    const ss = SpreadsheetApp.openById("COLE_AQUI_O_ID_DA_SUA_PLANILHA");
    const aba = ss.getSheetByName("PRONTUARIO"); // Nome exato da sua aba
    
    if (!aba) return [];

    const dados = aba.getDataRange().getValues();
    const historico = [];
    const nomeBusca = String(nomePaciente).trim().toLowerCase();

    // Começa em 1 para pular o cabeçalho
    for (let i = 1; i < dados.length; i++) {
      // Coluna B (índice 1) é o Paciente
      const nomeNaPlanilha = String(dados[i][1]).trim().toLowerCase();

      if (nomeNaPlanilha === nomeBusca) {
        // Coluna C (índice 2) é a Data
        let dataRaw = dados[i][2]; 
        let dataFormatada = (dataRaw instanceof Date) 
          ? Utilities.formatDate(dataRaw, "GMT-3", "dd/MM/yyyy") 
          : String(dataRaw);

        historico.push({
          data: dataFormatada,
          evolucao: dados[i][3] || "Sem evolução preenchida", // Coluna D
          tema: dados[i][5] || "Geral"                         // Coluna F
        });
      }
    }
    
    // Inverte para os mais recentes aparecerem primeiro no topo da tabela
    return historico.reverse();

  } catch (e) {
    console.error("Erro no servidor: " + e.message);
    return [];
  }
}

// EXCLUIR REGISTRO
function excluirRegistroProntuario(id) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PRONTUARIO");
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] == id) { // Coluna A é o ID
        aba.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch (e) {
    throw new Error("Erro ao excluir: " + e.message);
  }
}

// ATUALIZAR REGISTRO EXISTENTE
function atualizarRegistroProntuario(dados) {
  try {
    // 1. Use o ID da planilha onde aparecem os dados na imagem
    const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA"; 
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PRONTUARIO");
    const valores = aba.getDataRange().getValues();
    
    // ID que vem do formulário de edição
    const idProcurado = String(dados.id).trim();

    for (let i = 1; i < valores.length; i++) {
      // Compara com o ID na Coluna A (índice 0)
      if (String(valores[i][0]).trim() === idProcurado) {
        
        // --- ALTERA APENAS AS DUAS COLUNAS DESEJADAS ---
        
        // Coluna D (Evolução) é a 4ª coluna
        aba.getRange(i + 1, 4).setValue(dados.evolucao); 
        
        // Coluna F (Tema) é a 6ª coluna
        aba.getRange(i + 1, 6).setValue(dados.tema);     
        
        return true; 
      }
    }
    return false; // Retorna falso se não achar o ID na Coluna A
  } catch (e) {
    throw new Error("Erro ao salvar: " + e.message);
  }
}

/**
 * SALVA OS DADOS NAS COLUNAS 
 */
// --- FUNÇÃO PARA SALVAR TUDO NAS GAVETAS CERTAS ---
function salvarNovaSessao(d) {
  try {
    const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA";
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const hoje = new Date();
    const hojeF = Utilities.formatDate(hoje, "GMT-3", "dd/MM/yyyy");
    
    // 1. GRAVA NO PRONTUÁRIO
    const abaP = ss.getSheetByName("PRONTUARIO");
    abaP.appendRow([
      Utilities.getUuid(), d.paciente, hoje, d.evolucao, d.humor, d.tema, "", d.duracao, "Realizada"
    ]);

    // 2. BUSCA O VALOR NA ABA 'SESSOES' (AGENDA)
    const abaS = ss.getSheetByName("SESSOES");
    const dadosS = abaS.getDataRange().getValues();
    let valorEncontrado = 0;

    for (let i = 1; i < dadosS.length; i++) {
      let dataSessaoF = "";
      if (dadosS[i][2] instanceof Date) {
        dataSessaoF = Utilities.formatDate(dadosS[i][2], "GMT-3", "dd/MM/yyyy");
      } else {
        dataSessaoF = String(dadosS[i][2]);
      }
      
      // Verifica Paciente e Data (Coluna B e Coluna C)
      if (String(dadosS[i][1]).trim() === String(d.paciente).trim() && dataSessaoF === hojeF) {
        valorEncontrado = dadosS[i][6]; // Pega o valor da Coluna G (índice 6)
        
        // Marca como Realizada na Agenda (Coluna F e H)
        abaS.getRange(i + 1, 6).setValue("Realizada"); 
        abaS.getRange(i + 1, 8).setValue("Realizada");
        break;
      }
    }

    // 3. SE NÃO ACHOU VALOR NA AGENDA, tenta pegar o valor padrão do paciente ou usa 0
    // Isso evita que o financeiro fique vazio ou dê erro
    if (!valorEncontrado || valorEncontrado == 0) {
       console.log("Aviso: Valor não encontrado na agenda para hoje. Gravando com valor zero.");
    }

    // 4. GRAVA NO FINANCEIRO (Agora com proteção para não travar)
    const abaF = ss.getSheetByName("FINANCEIRO");
    abaF.appendRow([
      Utilities.getUuid(), 
      hoje, 
      "Sessão: " + d.paciente, 
      "Receita", 
      valorEncontrado, 
      "Atendimento", 
      "Pix/Dinheiro", 
      "Recebido"
    ]);

    return { sucesso: true };
  } catch (e) { 
    console.error("Erro ao salvar sessão completa: " + e.message);
    return { sucesso: false, erro: e.message }; 
  }
}

function buscarHistoricoProntuario(nomePaciente) {
  try {
    const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA"; 
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PRONTUARIO");
    const dados = aba.getDataRange().getValues();
    const nomeBusca = String(nomePaciente).trim().toLowerCase();

    return dados.slice(1)
      .filter(linha => String(linha[1]).trim().toLowerCase() === nomeBusca)
      .map(linha => ({
        id: String(linha[0]),        // ID da Coluna A
        data: linha[2] instanceof Date ? Utilities.formatDate(linha[2], "GMT-3", "dd/MM/yyyy HH:mm") : linha[2],
        evolucao: linha[3],          // Coluna D
        tema: linha[5]               // Coluna F
      }))
      .reverse();
  } catch (e) { 
    return []; 
  }
}

// Garanta que a busca também pegue o ID corretamente

// Função auxiliar para mudar o status na aba SESSOES automaticamente
function atualizarStatusParaRealizado(nomePaciente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaSessoes = ss.getSheetByName("SESSOES");
  const dados = abaSessoes.getDataRange().getValues();
  const hoje = new Date().toLocaleDateString('en-CA');

  for (let i = 1; i < dados.length; i++) {
    let dataPlanilha = "";
    if (dados[i][2] instanceof Date) {
      dataPlanilha = dados[i][2].toISOString().split('T')[0];
    } else {
      dataPlanilha = String(dados[i][2]);
    }

    if (dados[i][1] === nomePaciente && dataPlanilha === hoje) {
      abaSessoes.getRange(i + 1, 6).setValue("Realizado"); // Coluna F
      break;
    }
  }
}

// BUSCA O HISTÓRICO DAS SESSÕES PARA O PRONTUARIO
function buscarHistoricoPaciente(idPaciente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("PRONTUARIO");
    
    if (!sheet) return [];

    // Pegamos todos os dados transformando tudo em TEXTO (DisplayValues)
    const dados = sheet.getDataRange().getDisplayValues();
    
    // O filtro: Compara a Coluna A (índice 0) com o idPaciente
    const historico = dados.filter(linha => {
      // Trim remove espaços vazios que podem estar escondidos
      return linha[0].toString().trim() === idPaciente.toString().trim();
    });

    console.log("ID Buscado: " + idPaciente);
    console.log("Linhas encontradas: " + historico.length);
    
    return historico.reverse(); 
  } catch (e) {
    return [];
  }
}

/**
 * MÓDULO: FINANCEIRO E PDF
 */

function listarPacientes() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PACIENTES");
    const dados = aba.getDataRange().getValues();

    // Filtramos para garantir que só venham linhas onde o NOME (Coluna B / Índice 1) exista
    const listaFiltrada = dados.filter((linha, index) => {
      return index > 0 && linha[1] && String(linha[1]).trim() !== "";
    });

    // Tratamos cada linha para garantir que datas e tipos complexos não quebrem o envio
    return listaFiltrada.map(r => {
      return r.map(celula => {
        if (celula instanceof Date) {
          return Utilities.formatDate(celula, "GMT-3", "dd/MM/yyyy");
        }
        return celula === null || celula === undefined ? "" : celula;
      });
    });

  } catch (e) {
    console.error("Erro ao listar pacientes: " + e.message);
    return [];
  }
}

function listarHistoricoFiscal() {
  try {
    const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA";
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("HISTORICO_FISCAL");
    const dados = aba.getDataRange().getValues();
    
    let logs = [];
    // Começa do 1 para pular o cabeçalho
    for (let i = 1; i < dados.length; i++) {
      logs.push({
        id: dados[i][0],              // Coluna A
        paciente: dados[i][1],        // Coluna B
        cpfPagador: dados[i][3],      // Coluna D
        valor: dados[i][4],           // Coluna E
        dataPagamento: dados[i][5],   // Coluna F
        dataRegistro: dados[i][6]     // Coluna G
      });
    }
    
    // Inverte para mostrar os mais recentes primeiro
    return logs.reverse(); 
  } catch (e) {
    console.log("Erro ao listar histórico: " + e.message);
    return [];
  }
}

// FUNÇÃO PARA LISTAR O FINANCEIRO
function listarFinanceiro() {
  try {
    const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA";
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const abaSessoes = ss.getSheetByName("SESSOES");
    const dados = abaSessoes.getDataRange().getValues();
    
    // Pegamos também o que já foi registrado para não mostrar duplicado
    const abaFin = ss.getSheetByName("FINANCEIRO");
    const registrosExistentes = abaFin.getDataRange().getValues().map(r => r[0]); // Pega os IDs/UUIDs se você usa
    
    let pendentes = [];
    
    // Percorre a aba SESSOES (Coluna F é o Status, índice 5)
    for (let i = 1; i < dados.length; i++) {
      const status = dados[i][5]; // Coluna F
      const paciente = dados[i][1]; // Coluna B
      const data = Utilities.formatDate(new Date(dados[i][2]), "GMT-3", "dd/MM/yyyy"); // Coluna C
      const valor = dados[i][6]; // Coluna G
      const idSessao = dados[i][0]; // Coluna A (ID Único)

      // CRITÉRIO: Status "Realizada" e que não esteja no Financeiro (opcional se você deletar da aba sessoes)
      if (status === "Realizada") {
        pendentes.push({
          id: idSessao,
          paciente: paciente,
          data: data,
          valor: valor
        });
      }
    }
    
    return pendentes;
  } catch (e) {
    return [];
  }
}

// Busca CPF e dados fiscais do paciente pelo nome
function getDadosFiscaisPaciente(nome) {
  try {
    const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA";
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("PACIENTES");
    const dados = aba.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][1] == nome) { // Coluna B é o Nome
        return { 
          cpf: dados[i][3] // MUDANÇA AQUI: Índice 3 é a Coluna D (CPF)
        };
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}


// SALVA A NOVA EVOLUÇÃO
function salvarProntuario(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPacientes = ss.getSheetByName("PACIENTES");
    const sheetProntuario = ss.getSheetByName("PRONTUARIO");
    
    if (!sheetPacientes || !sheetProntuario) {
      throw new Error("Aba PACIENTES ou PRONTUARIO não encontrada.");
    }

    // 1. BUSCAR O UUID REAL NA ABA PACIENTES
    const dadosPacientes = sheetPacientes.getDataRange().getValues();
    let uuidReal = "";
    
    // Procura o nome do paciente (Coluna B) para pegar o ID (Coluna A)
    for (let i = 1; i < dadosPacientes.length; i++) {
      if (dadosPacientes[i][1] === dados.paciente) {
        uuidReal = dadosPacientes[i][0]; // Pega o UUID da Coluna A
        break;
      }
    }

    // Se por algum motivo não achar o UUID, usa o que veio do front ou gera um erro
    if (!uuidReal) {
      uuidReal = dados.id; 
    }

    // 2. PREPARAR A LINHA PARA GRAVAR NO PRONTUARIO
    // Ordem: ID, Paciente, Data, Evolução, Humor, Tema, Tarefa, Tempo, Status
    const novaLinha = [
      uuidReal,           // Coluna A (UUID correto)
      dados.paciente,     // Coluna B
      dados.data,         // Coluna C
      dados.evolucao,     // Coluna D
      dados.humor,        // Coluna E
      dados.tema,         // Coluna F
      dados.tarefa,       // Coluna G
      dados.tempo,        // Coluna H
      "Finalizado"        // Coluna I (Status)
    ];

    sheetProntuario.appendRow(novaLinha);
    return true;

  } catch (e) {
    console.error("Erro ao salvar: " + e.message);
    return "Erro: " + e.message;
  }
}

/**
 * EXECUTE ESTA FUNÇÃO MANUALMENTE NO EDITOR DE SCRIPT
 * Ela serve apenas para forçar o Google a pedir autorização total.
 */
function forcarPermissoes() {
  // Tenta abrir a planilha pelo ID (Troque pelo seu ID se não for esse)
  const ID_REAL = "COLE_AQUI_O_ID_DA_SUA_PLANILHA"; 
  const ss = SpreadsheetApp.openById(ID_REAL);
  
  // Pega a primeira aba disponível
  const aba = ss.getSheets()[0];
  
  // Realiza uma operação de escrita "fantasma" na célula Z1000
  // Isso obriga o Google a pedir autorização de EDIÇÃO
  aba.getRange("Z1000").setValue("Teste de Conexão");
  aba.getRange("Z1000").clearContent();
  
  console.log("Conexão estabelecida com sucesso para: " + ss.getName());
}

// FUNÇÃO PARA TESTAR O ACESSO (Execute esta primeiro no editor)
function testarConexao() {
  const ID_REAL = "COLE_AQUI_O_ID_DA_SUA_PLANILHA"; 
  const ss = SpreadsheetApp.openById(ID_REAL);
  Logger.log("Conectado com sucesso à planilha: " + ss.getName());
}

function listarAgendadosHoje() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const abaS = ss.getSheetByName("SESSOES");
    const dados = abaS.getDataRange().getValues();
    
    const hoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
    let listaHoje = [];

    for (let i = 1; i < dados.length; i++) {
      let dataSessao = dados[i][2]; 
      let dataFormatada = dataSessao instanceof Date ? Utilities.formatDate(dataSessao, "GMT-3", "dd/MM/yyyy") : String(dataSessao);
      
      // Se a data for hoje e não estiver cancelado
      if (dataFormatada === hoje && String(dados[i][5]).toUpperCase() !== "CANCELADO") {
        listaHoje.push(dados[i][1]); // Pega o nome do paciente (Coluna B)
      }
    }
    return listaHoje;
  } catch (e) {
    return [];
  }
}

/**
 * Busca as sessões realizadas e cruza com os dados do responsável (Coluna M)
 */

function getPlanilha() {
  // Troque o texto abaixo pelo ID que está na URL da sua planilha
  const ID_PLANILHA = "COLE_AQUI_O_ID_DA_SUA_PLANILHA"; 
  
  try {
    return SpreadsheetApp.openById(ID_PLANILHA);
  } catch (e) {
    // Se falhar pelo ID, tenta pegar a que estiver aberta
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function listarSessoesComResponsavel() {
  try {
    const ss = getPlanilha(); 
    const abaSessoes = ss.getSheetByName("SESSOES");
    const abaPacientes = ss.getSheetByName("PACIENTES");
    
    const dadosSessao = abaSessoes.getDataRange().getValues();
    const dadosPac = abaPacientes.getDataRange().getValues();
    
    // Cria um mapa de CPFs e Responsáveis usando o Nome do Paciente como chave
    // Nome: Col B(1) | CPF: Col D(3) | Responsável: Col M(12)
    let infoPacientes = {};
    for(let i = 1; i < dadosPac.length; i++) {
      let nome = String(dadosPac[i][1]).trim();
      if (nome) {
        infoPacientes[nome] = {
          cpf: String(dadosPac[i][3] || "").trim(), // Coluna D
          responsavel: String(dadosPac[i][12] || "").trim() // Coluna M
        };
      }
    }

    let listaFinal = [];
    for (let i = 1; i < dadosSessao.length; i++) {
      let status = String(dadosSessao[i][5]).trim().toLowerCase();
      
      // Só processa se o status for "realizada"
      if (status === "realizada") {
        let nomePac = String(dadosSessao[i][1]).trim();
        let pacienteDados = infoPacientes[nomePac] || { cpf: "Não Encontrado", responsavel: "" };

        listaFinal.push({
          paciente: nomePac,
          cpfPaciente: pacienteDados.cpf, // CPF da Coluna D mapeado aqui
          responsavel: pacienteDados.responsavel,
          data: dadosSessao[i][2] instanceof Date ? 
                Utilities.formatDate(dadosSessao[i][2], "GMT-3", "dd/MM/yyyy") : 
                dadosSessao[i][2],
          valor: dadosSessao[i][6]
        });
      }
    }
    return listaFinal;
  } catch (e) {
    return { erro: "Erro ao listar: " + e.message };
  }
}

function salvarRegistroFiscalCompleto(dados) {
  try {
    const ss = getPlanilha();
    const abaHist = ss.getSheetByName("HISTORICO_FISCAL");
    const abaSessoes = ss.getSheetByName("SESSOES");
    
    if (!abaHist || !abaSessoes) return { sucesso: false, erro: "Abas não encontradas." };

    // --- PASSO 1: Gravar no Histórico Fiscal ---
    abaHist.appendRow([
      Utilities.getUuid(),
      dados.paciente,
      dados.cpfPaciente,
      dados.nomePagador,
      dados.cpfPagador,
      dados.valor,
      dados.dataSessao,
      new Date()
    ]);

    // --- PASSO 2: Mudar Status na aba SESSOES para não repetir ---
    const dadosSessao = abaSessoes.getDataRange().getValues();
    
    // Procuramos a linha correta: Nome do Paciente (Col B) e Data (Col C)
    for (let i = 1; i < dadosSessao.length; i++) {
      let dataPlanilha = dadosSessao[i][2] instanceof Date ? 
                         Utilities.formatDate(dadosSessao[i][2], "GMT-3", "dd/MM/yyyy") : 
                         dadosSessao[i][2];

      if (dadosSessao[i][1] == dados.paciente && dataPlanilha == dados.dataSessao) {
        // Coluna F é o índice 5 (Status)
        abaSessoes.getRange(i + 1, 6).setValue("Faturado"); 
        break; 
      }
    }

    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * BUSCA FILTRADA NO HISTÓRICO FISCAL
 * Filtra por nome do paciente e intervalo de datas
 */
function buscarHistoricoFiltrado(paciente, dataInicio, dataFim) {
  try {
    const ss = getPlanilha();
    const aba = ss.getSheetByName("HISTORICO_FISCAL");
    const dados = aba.getDataRange().getValues();
    
    const dInicio = dataInicio ? new Date(dataInicio + "T00:00:00") : null;
    const dFim = dataFim ? new Date(dataFim + "T23:59:59") : null;

    let resultado = [];

    for (let i = 1; i < dados.length; i++) {
      const nomePac = dados[i][1];
      const dataSessaoStr = dados[i][6]; 
      
      if (!dataSessaoStr) continue; // Pula linhas totalmente vazias se houver

      const partes = String(dataSessaoStr).split("/");
      if (partes.length < 3) continue;
      const dataSessaoObj = new Date(partes[2], partes[1] - 1, partes[0]);

      const bateNome = !paciente || nomePac.toLowerCase().includes(paciente.toLowerCase());
      const bateDataInic = !dInicio || dataSessaoObj >= dInicio;
      const bateDataFim = !dFim || dataSessaoObj <= dFim;

      if (bateNome && bateDataInic && bateDataFim) {
        resultado.push({
          id: String(dados[i][0]),
          paciente: nomePac,
          cpfPac: dados[i][2] || "Não Informado", // Blindagem para não vir undefined se estiver vazio
          cpfPaciso: dados[i][2] || "Não Informado", // Segurança caso o front use este nome
          pagador: dados[i][3] || "O Próprio",
          cpfPag: dados[i][4] || "Não Informado",
          valor: dados[i][5],
          dataSessao: dataSessaoStr
        });
      }
    }
    return resultado.reverse();
  } catch (e) {
    return { erro: e.message };
  }
}
// ==========================================
// MÓDULO DE GERAÇÃO DE PDF (PODE COLAR NO FINAL DO SERVIDOR.GS)
// ==========================================

function gerarPDFRecibo(dadosFiscais) {
  try {
    let r = null;

    // SE RECEBER UM OBJETO COMPLETO (Estratégia nova e rápida do Botão Vermelho)
    if (dadosFiscais && typeof dadosFiscais === 'object') {
      r = {
        paciente: dadosFiscais.paciente || "",
        cpfPac: dadosFiscais.cpfBeneficiario || "Não informado",
        resp: dadosFiscais.responsavelPagamento || dadosFiscais.paciente,
        cpfResp: dadosFiscais.cpfPagador || "",
        valor: dadosFiscais.valor || "0,00",
        data: dadosFiscais.dataPagamento || ""
      };
    } 
    // SE RECEBER APENAS UM ID (Mantém a compatibilidade caso outro botão do seu sistema use)
    else if (dadosFiscais) {
      const ss = typeof getPlanilha === "function" ? getPlanilha() : SpreadsheetApp.getActiveSpreadsheet();
      const aba = ss.getSheetByName("HISTORICO_FISCAL");
      if (!aba) return "Erro: Aba HISTORICO_FISCAL não encontrada.";
      
      const dadosPlanilha = aba.getDataRange().getValues();
      for (let i = 1; i < dadosPlanilha.length; i++) {
        if (dadosPlanilha[i][0] == dadosFiscais) { // Procura pelo ID na coluna A
          r = { 
            paciente: dadosPlanilha[i][1], 
            cpfPac: dadosPlanilha[i][2], 
            resp: dadosPlanilha[i][3] || dadosPlanilha[i][1], 
            cpfResp: dadosPlanilha[i][4], 
            valor: dadosPlanilha[i][5], 
            data: dadosPlanilha[i][6] 
          };
          break;
        }
      }
    }

    // Se mesmo após as checagens não encontrar dados válidos
    if (!r || !r.paciente) return "Erro: Registro ou informações fiscais não encontradas.";

    // O texto HTML que vira PDF (Mantive seu layout elegante original)
    const html = `
      <div style="padding: 40px; font-family: 'Helvetica', sans-serif; color: #333; border: 1px solid #ddd; max-width: 600px; margin: 0 auto;">
        <div style="text-align: center; margin-bottom: 30px;">
          <h1 style="margin: 0; color: #1e293b; letter-spacing: 2px;">RECIBO</h1>
          <p style="font-size: 12px; color: #64748b; text-transform: uppercase; margin-top: 5px;">Psicoterapia Especializada</p>
        </div>
        
        <p style="font-size: 16px; line-height: 1.8; text-align: justify; color: #334155;">
          Recebi de <b>${r.resp}</b>, inscrito(a) no CPF sob o nº <b>${r.cpfResp || '___________________'}</b>, 
          a importância de <b>R$ ${r.valor}</b> referente a serviços de atendimento psicoterápico 
          prestados ao beneficiário(a) <b>${r.paciente}</b> ${r.cpfPac ? '(CPF: ' + r.cpfPac + ')' : ''}, 
          realizado em <b>${r.data}</b>.
        </p>
        
        <div style="margin-top: 80px; text-align: center;">
          <p style="color: #64748b; font-size: 14px;">${Utilities.formatDate(new Date(), "GMT-3", "dd 'de' MMMM 'de' yyyy")}</p>
          <br><br>
          <div style="border-top: 1px solid #94a3b8; width: 280px; margin: 0 auto;"></div>
          <p style="margin-top: 8px; font-size: 14px; color: #1e293b;">
            <b>Seu Nome Completo</b><br>
            <span style="color: #64748b; font-size: 12px;">Psicólogo(a) - CRP: 00/00000</span>
          </p>
        </div>
      </div>
    `;

    const blob = Utilities.newBlob(html, "text/html", "recibo.html");
    const pdf = blob.getAs("application/pdf").setName("Recibo_" + r.paciente.replace(/\s+/g, "_") + ".pdf");
    
    return "data:application/pdf;base64," + Utilities.base64Encode(pdf.getBytes());

  } catch (erro) {
    return "Erro ao processar PDF: " + erro.toString();
  }
}

function gerarPDFDoHistoricoGeral(mesFiltrado, anoFiltrado) {
  try {
    // ABRE A PLANILHA DIRETAMENTE PELO ID SEGURO
    const ss = SpreadsheetApp.openById("COLE_AQUI_O_ID_DA_SUA_PLANILHA");
    const sheet = ss.getSheetByName("HISTORICO_FISCAL");
    if (!sheet) return "Erro: Aba HISTORICO_FISCAL não encontrada.";
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "Erro: Nenhum registro encontrado no Histórico Fiscal.";
    
    let linhasHtml = "";
    let totalGeral = 0;
    let contador = 0;
    
    // Varre a planilha a partir da linha 2 (índice 1)
    for (let i = 1; i < data.length; i++) {
      const paciente = data[i][1];             // Coluna B
      const cpfBeneficiario = data[i][2];      // Coluna C
      const responsavelPagamento = data[i][3]; // Coluna D
      const cpfPagador = data[i][4];           // Coluna E
      const valorRaw = data[i][5];             // Coluna F
      const dataPagamento = String(data[i][6]);// Coluna G
      
      // Filtros de Mês e Ano
      let corresponde = true;
      if (mesFiltrado && !dataPagamento.includes(mesFiltrado)) corresponde = false;
      if (anoFiltrado && !dataPagamento.includes(anoFiltrado)) corresponde = false;
      
      if (corresponde) {
        contador++;
        const valorNumerico = parseFloat(String(valorRaw).replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
        totalGeral += valorNumerico;

        // Monta a linha da tabela em HTML
        linhasHtml += `
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 10px; font-size: 12px; color: #334155;">${paciente}</td>
            <td style="padding: 10px; font-size: 12px; color: #334155;">${cpfBeneficiario || '-'}</td>
            <td style="padding: 10px; font-size: 12px; color: #334155;">${responsavelPagamento || paciente}</td>
            <td style="padding: 10px; font-size: 12px; color: #334155;">${cpfPagador || '-'}</td>
            <td style="padding: 10px; font-size: 12px; color: #334155; text-align: center;">${dataPagamento}</td>
            <td style="padding: 10px; font-size: 12px; color: #1e293b; font-weight: bold; text-align: right;">R$ ${valorNumerico.toFixed(2).replace('.', ',')}</td>
          </tr>
        `;
      }
    }
    
    if (contador === 0) {
      return "Erro: Nenhum registro localizado no HISTORICO_FISCAL para o período filtrado.";
    }
    
    // Estrutura do documento em formato de Relatório/Lista Comercial
    const htmlCompleto = `
      <div style="padding: 30px; font-family: 'Helvetica', sans-serif;">
        <div style="border-bottom: 2px solid #1e293b; padding-bottom: 15px; margin-bottom: 20px;">
          <h1 style="margin: 0; color: #1e293b; font-size: 22px;">Extrato do Histórico Fiscal</h1>
          <p style="margin: 5px 0 0 0; color: #64748b; font-size: 12px;">Período de Referência: ${mesFiltrado || 'Todos'} / ${anoFiltrado || 'Todos'}</p>
        </div>
        
        <table style="width: 100%; border-collapse: collapse; text-align: left; margin-bottom: 20px;">
          <thead>
            <tr style="background: #f1f5f9; border-bottom: 2px solid #cbd5e1;">
              <th style="padding: 12px 10px; font-size: 11px; color: #475569; text-transform: uppercase;">Paciente</th>
              <th style="padding: 12px 10px; font-size: 11px; color: #475569; text-transform: uppercase;">CPF Paciente</th>
              <th style="padding: 12px 10px; font-size: 11px; color: #475569; text-transform: uppercase;">Resp. Pagamento</th>
              <th style="padding: 12px 10px; font-size: 11px; color: #475569; text-transform: uppercase;">CPF Pagador</th>
              <th style="padding: 12px 10px; font-size: 11px; color: #475569; text-transform: uppercase; text-align: center;">Data Pgto</th>
              <th style="padding: 12px 10px; font-size: 11px; color: #475569; text-transform: uppercase; text-align: right;">Valor</th>
            </tr>
          </thead>
          <tbody>
            ${linhasHtml}
          </tbody>
        </table>
        
        <div style="display: flex; justify-content: flex-end; margin-top: 30px; text-align: right;">
          <div style="display: inline-block; background: #f8fafc; border: 1px solid #e2e8f0; padding: 15px 25px; border-radius: 8px;">
            <span style="font-size: 12px; color: #64748b; text-transform: uppercase; display: block; margin-bottom: 5px;">Total do Período</span>
            <strong style="font-size: 20px; color: #0f172a;">R$ ${totalGeral.toFixed(2).replace('.', ',')}</strong>
          </div>
        </div>
      </div>
    `;
    
    // Converte o relatório consolidado em PDF e retorna a Base64
    const blob = Utilities.newBlob(htmlCompleto, "text/html", "relatorio.html");
    const pdf = blob.getAs("application/pdf").setName("Relatorio_Fiscal_" + (mesFiltrado || "Geral").replace(/\//g, "") + "_" + anoFiltrado + ".pdf");
    
    return "data:application/pdf;base64," + Utilities.base64Encode(pdf.getBytes());
    
  } catch (e) {
    return "Erro interno no servidor: " + e.toString();
  }
}

//* MODOLO DE CONFIGURAÇÕES 
// Lista todos os usuários para a tela de Configurações
function listarUsuariosConfig() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("USUÁRIOS"); // Verifique se o nome tem acento como na aba
    if (!aba) return [];

    const dados = aba.getDataRange().getValues();
    
    // Remove o cabeçalho (linha 1) e filtra linhas vazias
    return dados.slice(1).filter(linha => linha[0] !== "");
  } catch (e) {
    console.error("Erro ao listar usuários: " + e.toString());
    return [];
  }
}

/**
 * Atualiza Senha e Perfil de um usuário específico
 */
function atualizarAcessoUsuario(dados) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("USUÁRIOS");
    const valores = aba.getDataRange().getValues();
    
    let linhaDestino = -1;

    // Localiza a linha pelo ID (Coluna A)
    for (let i = 0; i < valores.length; i++) {
      if (String(valores[i][0]).trim() === String(dados.id).trim()) {
        linhaDestino = i + 1; // Soma 1 porque o array começa em 0 e a planilha em 1
        break;
      }
    }

    if (linhaDestino !== -1) {
      // De acordo com sua planilha:
      // Coluna D (índice 4) é Senha
      // Coluna E (índice 5) é PERFIL
      aba.getRange(linhaDestino, 4).setValue(dados.senha);
      aba.getRange(linhaDestino, 5).setValue(dados.perfil);
      
      return { sucesso: true };
    }
    
    return { sucesso: false, erro: "Usuário não encontrado no sistema." };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

// Cria um novo usuário com ID automático
function criarNovoUsuario(u) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("USUÁRIOS");
    
    // Gera um ID baseado na última linha (ex: 1, 2, 3...)
    const proxId = aba.getLastRow(); 
    
    // Salva na ordem da sua planilha: ID, Nome, Email, Senha, Perfil
    aba.appendRow([proxId, u.nome, u.email, u.senha, "PSICOLOGO"]);
    
    return { sucesso: true };
  } catch (e) { 
    return { sucesso: false, erro: e.toString() }; 
  }
}

// Deleta a linha do usuário pelo ID
function excluirUsuarioSistema(id) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("USUÁRIOS");
    const dados = aba.getDataRange().getValues();
    for (let i = 0; i < dados.length; i++) {
      if (String(dados[i][0]) === String(id)) {
        aba.deleteRow(i + 1);
        return { sucesso: true };
      }
    }
  } catch (e) { return { sucesso: false, erro: e.toString() }; }
}



/**
 * FUNÇÃO CENTRAL DE LOG - Executada no Servidor
 */
function salvarNoBancoDeLogs(logObjeto) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("LOGS");
    if (!aba) return;

    // Gera um ID simples baseado na última linha
    const ultimoId = aba.getLastRow();
    const novoId = ultimoId === 0 ? 1 : ultimoId;

    // Ordem exata da sua planilha: 
    // A: Id | B: DataHora | C: Usuario | D: Acao | E: Tabela | F: Detalhes
    aba.appendRow([
      novoId,           // Coluna A (Id)
      new Date(),       // Coluna B (DataHora)
      logObjeto.usuario,// Coluna C (Usuario)
      logObjeto.acao,   // Coluna D (Acao)
      logObjeto.modulo, // Coluna E (Tabela - mapeado como seu módulo)
      logObjeto.detalhes// Coluna F (Detalhes)
    ]);
  } catch (e) {
    console.error("Erro ao gravar log: " + e.toString());
  }
}