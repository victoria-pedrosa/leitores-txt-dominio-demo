// Substitua pela sua chave de API do Gemini
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

// Tipos de arquivo aceitos ao varrer uma pasta
const MIME_ACEITOS = ['application/pdf', 'image/jpeg', 'image/png', 'image/jpg'];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Realizar Leitura')
    .addItem('1. Importar Nota (1 link)', 'importarNotaTomada')
    .addItem('2. Importar Vários Links', 'importarVariosLinks')
    .addItem('3. Importar Pasta Inteira (Drive)', 'importarPastaTomada')
    .addSeparator()
    .addItem('4. Gerar Arquivo TXT (Domínio)', 'gerarTxtDominio')
    .addToUi();
}

// ------------------------------------------------------------------
// 1. Um único link/ID
// ------------------------------------------------------------------
function importarNotaTomada() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Importar NFS-e de Serviços Tomados',
    'Cole o link ou o ID do arquivo PDF/imagem do Google Drive:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.CANCEL) return;
  const fileId = extrairIdDrive(response.getResponseText());
  if (!fileId) {
    ui.alert('Erro', 'Não consegui identificar o ID do arquivo nesse link.', ui.ButtonSet.OK);
    return;
  }

  const resultado = processarNota(fileId);
  if (resultado.ok) {
    ui.alert('Sucesso', 'NFS-e de Serviços Tomados importada com sucesso!', ui.ButtonSet.OK);
  } else {
    ui.alert('Erro', 'Ocorreu um erro: ' + resultado.erro, ui.ButtonSet.OK);
  }
}

// ------------------------------------------------------------------
// 2. Vários links/IDs colados de uma vez (um por linha)
// ------------------------------------------------------------------
function importarVariosLinks() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Importar Vários Links',
    'Cole os links ou IDs dos arquivos, um por linha:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.CANCEL) return;

  const linhas = response.getResponseText()
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l !== '');

  if (linhas.length === 0) {
    ui.alert('Aviso', 'Nenhum link informado.', ui.ButtonSet.OK);
    return;
  }

  const ids = linhas.map(extrairIdDrive).filter(id => id);
  processarLoteEIds(ids, ui);
}

// ------------------------------------------------------------------
// 3. Varrer uma pasta inteira do Drive
// ------------------------------------------------------------------
function importarPastaTomada() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Importar Pasta Inteira',
    'Cole o link ou o ID da pasta do Google Drive com os PDFs/imagens das notas:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.CANCEL) return;
  const folderId = extrairIdDrive(response.getResponseText());
  if (!folderId) {
    ui.alert('Erro', 'Não consegui identificar o ID da pasta nesse link.', ui.ButtonSet.OK);
    return;
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    ui.alert('Erro', 'Pasta não encontrada ou sem permissão de acesso.', ui.ButtonSet.OK);
    return;
  }

  const ids = [];
  MIME_ACEITOS.forEach(mime => {
    const arquivos = folder.getFilesByType(mime);
    while (arquivos.hasNext()) {
      ids.push(arquivos.next().getId());
    }
  });

  if (ids.length === 0) {
    ui.alert('Aviso', 'Nenhum PDF ou imagem encontrado nessa pasta.', ui.ButtonSet.OK);
    return;
  }

  processarLoteEIds(ids, ui);
}

// ------------------------------------------------------------------
// Núcleo comum: processa uma lista de IDs e mostra um resumo no final
// ------------------------------------------------------------------
function processarLoteEIds(ids, ui) {
  let sucesso = 0;
  const falhas = [];

  ids.forEach((fileId) => {
    const resultado = processarNota(fileId);
    if (resultado.ok) {
      sucesso++;
    } else {
      falhas.push(`- ${fileId}: ${resultado.erro}`);
    }
    Utilities.sleep(1200); // evita estourar a cota da API em lotes grandes
  });

  let msg = `Notas processadas com sucesso: ${sucesso} de ${ids.length}.`;
  if (falhas.length > 0) {
    msg += `\n\nFalharam (revisar manualmente):\n${falhas.join('\n')}`;
  }
  ui.alert(falhas.length > 0 ? 'Concluído com pendências' : 'Concluído', msg, ui.ButtonSet.OK);
}

// Extrai o ID de um link do Drive em qualquer formato comum, ou retorna o texto
// já se ele mesmo for o ID.
function extrairIdDrive(texto) {
  if (!texto) return null;
  const t = texto.trim();
  const padroes = [/\/d\/([a-zA-Z0-9_-]{10,})/, /[?&]id=([a-zA-Z0-9_-]{10,})/];
  for (const p of padroes) {
    const m = t.match(p);
    if (m) return m[1];
  }
  if (/^[a-zA-Z0-9_-]{10,}$/.test(t)) return t;
  return null;
}

// ------------------------------------------------------------------
// Lê 1 arquivo (PDF ou imagem), chama o Gemini e grava a linha na planilha.
// ------------------------------------------------------------------
function processarNota(fileId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType();
    const base64Data = Utilities.base64Encode(blob.getBytes());

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${GEMINI_API_KEY}`;

    // Leiaute "Notas Fiscais de Serviços Tomados (Leiaute Domínio - Excel)" - 28 campos
    const promptTexto = `Leia esta nota fiscal de serviço tomado (NFS-e).
    Retorne APENAS um objeto JSON válido, sem formatação markdown ou crases.
    Para valores monetários vazios ou não encontrados, use "0,00". Para textos vazios, use "".
    Sempre use vírgula para casas decimais (ex: "1500,50").

    O JSON deve conter EXATAMENTE as seguintes chaves:
    "cnpj_prestador", "razao_prestador", "uf_prestador", "municipio_prestador", "endereco_prestador",
    "numero_nf", "serie" (se não tiver, retorne "1"), "data_emissao" (formato DD/MM/AAAA),
    "valor_servicos", "valor_descontos", "valor_contabil", "base_calculo",
    "aliquota_iss", "valor_iss_normal", "valor_iss_retido",
    "valor_irrf", "valor_pis", "valor_cofins", "valor_csll", "valor_crf", "valor_inss"`;

    const payload = {
      "contents": [{
        "parts": [
          { "text": promptTexto },
          { "inline_data": { "mime_type": mimeType, "data": base64Data } }
        ]
      }]
    };

    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const res = UrlFetchApp.fetch(url, options);
    const resText = res.getContentText();

    if (res.getResponseCode() !== 200) {
      return { ok: false, erro: 'Erro na API do Gemini: ' + resText };
    }

    const resultJson = JSON.parse(resText);
    let textoResposta = resultJson.candidates[0].content.parts[0].text;
    textoResposta = textoResposta.replace(/```json/g, '').replace(/```/g, '').trim();
    const dados = JSON.parse(textoResposta);

    // MAPEAMENTO DAS 28 POSIÇÕES DO LEIAUTE DOMÍNIO - SERVIÇOS TOMADOS
    // Validado byte-a-byte contra a linha de exemplo oficial do template (solução 9264)
    const linhaNova = [
      dados.cnpj_prestador || '',      // 1. CPF/CNPJ
      dados.razao_prestador || '',     // 2. Razão Social
      dados.uf_prestador || '',        // 3. UF
      dados.municipio_prestador || '', // 4. Município
      dados.endereco_prestador || '',  // 5. Endereço
      dados.numero_nf || '',           // 6. Número Documento
      dados.serie || '1',              // 7. Série
      dados.data_emissao || '',        // 8. Data de Emissão
      dados.data_emissao || '',        // 9. Data de Entrada (repete a emissão)
      '0',                             // 10. Situação (0-Regular)
      '800',                           // 11. Acumulador
      '1933',                          // 12. CFOP
      dados.valor_servicos,            // 13. Valor Serviços
      dados.valor_descontos,           // 14. Valor Descontos
      dados.valor_contabil,            // 15. Valor Contábil
      dados.base_calculo,              // 16. Base de Cálculo
      dados.aliquota_iss,              // 17. Alíquota ISS
      dados.valor_iss_normal,          // 18. Valor ISS Normal
      dados.valor_iss_retido,          // 19. Valor ISS Retido
      dados.valor_irrf,                // 20. Valor IRRF
      dados.valor_pis,                 // 21. Valor PIS
      dados.valor_cofins,              // 22. Valor COFINS
      dados.valor_csll,                // 23. Valor CSLL
      dados.valor_crf,                 // 24. Valor CRF
      dados.valor_inss,                // 25. Valor INSS
      '',                              // 26. Código do Item
      '',                              // 27. Quantidade
      ''                               // 28. Valor Unitário
    ];

    sheet.appendRow(linhaNova);
    return { ok: true };

  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

// Substitui a Macro do Excel: lê a planilha e gera o TXT (separador ";", decimal ",", quebra CRLF)
// Localiza a linha em que os dados de fato começam: procura a linha cujo
// primeiro valor é exatamente "CPF/CNPJ" (o cabeçalho real da coluna) e devolve
// a linha seguinte. O template oficial do Domínio tem várias linhas de
// instrução e um cabeçalho mesclado ANTES do cabeçalho real — sem isso, essas
// linhas eram exportadas como se fossem notas de verdade.
function localizarInicioDados(data) {
  for (let i = 0; i < data.length; i++) {
    const primeiraCelula = (data[i][0] || '').toString().trim().toUpperCase();
    if (primeiraCelula === 'CPF/CNPJ') {
      return i + 1;
    }
  }
  return 1; // fallback: comportamento antigo, caso não encontre o cabeçalho
}

// Tipo de cada uma das 28 colunas do leiaute (Serviços Tomados), na ordem exata.
const TIPOS_COLUNA = [
  'TEXTO', 'TEXTO', 'TEXTO', 'TEXTO', 'TEXTO',
  'INTEIRO', 'INTEIRO', 'DATA', 'DATA', 'INTEIRO', 'INTEIRO', 'INTEIRO',
  'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL',
  'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL',
  'DECIMAL', 'DECIMAL', 'DECIMAL',
  'INTEIRO', 'INTEIRO', 'DECIMAL'
];

// Formata uma célula para o TXT de acordo com o tipo esperado pelo Domínio.
// O Google Sheets converte sozinho texto de data/número digitado pelo script
// em valores de Data/Número reais — sem tratar isso, "43" (Número Documento)
// virava "43,00" (formato de valor monetário) e a importação rejeitava o
// campo por não ser do tipo numérico esperado.
function formatarCelula(valor, tipo) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  if (valor === '' || valor === null || valor === undefined) {
    return tipo === 'DECIMAL' ? '0,00' : '';
  }
  if (tipo === 'INTEIRO') {
    const numero = typeof valor === 'number' ? valor : parseFloat(valor.toString().replace(',', '.'));
    return isNaN(numero) ? valor.toString().trim() : Math.trunc(numero).toString();
  }
  if (tipo === 'DECIMAL') {
    const numero = typeof valor === 'number' ? valor : parseFloat(valor.toString().replace(',', '.'));
    return isNaN(numero) ? valor.toString().trim() : numero.toFixed(2).replace('.', ',');
  }
  return valor.toString();
}

function gerarTxtDominio() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const data = sheet.getDataRange().getValues();
  const inicio = localizarInicioDados(data);

  if (data.length <= inicio) {
    ui.alert('Aviso', 'Não há notas importadas na planilha para exportar.', ui.ButtonSet.OK);
    return;
  }

  let txtContent = "";
  for (let i = inicio; i < data.length; i++) {
    const cnpj = (data[i][0] || '').toString().trim();
    if (cnpj === '' || cnpj === '00.000.001/0001-00') continue; // pula vazias e a linha de exemplo do Domínio
    const linha = data[i].map((valor, idx) => formatarCelula(valor, TIPOS_COLUNA[idx])).join(";");
    txtContent += linha + "\r\n";
  }

  if (txtContent === '') {
    ui.alert('Aviso', 'Não há notas importadas na planilha para exportar.', ui.ButtonSet.OK);
    return;
  }

  const nomeArquivo = "ServicosTomados_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") + ".txt";
  DriveApp.createFile(nomeArquivo, txtContent, MimeType.PLAIN_TEXT);

  ui.alert('Pronto!', `O arquivo foi gerado com sucesso no seu Google Drive com o nome:\n\n${nomeArquivo}\n\nFaça o download e importe na Escrita Fiscal do Domínio (Conjunto de Dados: 'Notas Fiscais de Serviços Tomados (Leiaute Domínio - Excel)').`, ui.ButtonSet.OK);
}