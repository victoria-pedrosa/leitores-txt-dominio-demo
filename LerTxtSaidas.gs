// Substitua pela sua chave de API do Gemini
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

// Tipos de arquivo aceitos ao varrer uma pasta
const MIME_ACEITOS = ['application/pdf', 'image/jpeg', 'image/png', 'image/jpg'];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Realizar Leitura')
    .addItem('1. Importar Nota (1 link)', 'importarNotaSaida')
    .addItem('2. Importar Vários Links', 'importarVariosLinks')
    .addItem('3. Importar Pasta Inteira (Drive)', 'importarPastaSaida')
    .addSeparator()
    .addItem('4. Gerar Arquivo TXT (Domínio)', 'gerarTxtDominio')
    .addToUi();
}

// ------------------------------------------------------------------
// 1. Um único link/ID
// ------------------------------------------------------------------
function importarNotaSaida() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Importar Nota de Saída',
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
    const msg = resultado.linhas > 1
      ? `Nota de Saída importada com sucesso em ${resultado.linhas} linhas (1 por item da nota).`
      : 'Nota de Saída importada com sucesso!';
    ui.alert('Sucesso', msg, ui.ButtonSet.OK);
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
function importarPastaSaida() {
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

    // Leiaute "Notas Fiscais de Saídas (Leiaute Domínio - Excel) (1.1)" - 33 campos
    // Cada item da nota tem seu próprio CFOP. Quando a nota tem itens com CFOPs diferentes, o Domínio
    // espera uma LINHA POR GRUPO DE CFOP (confirmado batendo os subtotais de uma DANFE real com o
    // cabeçalho "CÁLCULO DO IMPOSTO"), não uma única linha fechada pra nota inteira.
    const promptTexto = `Leia esta nota fiscal de saída (produtos/mercadorias).
    Retorne APENAS um objeto JSON válido, sem formatação markdown.
    Para valores monetários vazios ou não encontrados, use "0,00". Para textos vazios, use "".
    Sempre use vírgula para casas decimais (ex: "1500,50").

    Cada item da nota tem seu próprio CFOP (coluna CFOP na tabela de produtos/serviços). Quando houver mais
    de um CFOP diferente entre os itens, agrupe os itens por CFOP e devolva um subtotal por grupo — nunca
    misture itens de CFOPs diferentes no mesmo grupo.

    O JSON deve conter EXATAMENTE esta estrutura:
    {
      "cnpj_cliente": "", "razao_cliente": "", "uf_cliente": "", "municipio_cliente": "",
      "endereco_cliente": "", "numero_nf": "",
      "serie": "" (número de série impresso na nota/DANFE; use "1" apenas se a nota realmente não
      mostrar nenhuma série),
      "chave_acesso": "" (os 44 números da CHAVE DE ACESSO impressa no topo da DANFE, só dígitos,
      sem espaços — usada para conferir a série; se não conseguir ler os 44 dígitos, deixe em branco),
      "data_emissao": "" (formato DD/MM/AAAA),
      "valor_descontos": "" (desconto total da nota, se houver, senão "0,00"),
      "cst_piscofins": "" (apenas números, ex: 01, 49, 99. Se não achar use "99"),
      "bc_piscofins": "", "aliq_pis": "", "valor_pis": "", "aliq_cofins": "", "valor_cofins": "",
      "itens": [
        {
          "codigo_item": "" (CÓDIGO PRODUTO impresso na coluna "Código Produto" da tabela de itens da nota/DANFE — copie exatamente como está escrito, sem remover zeros à esquerda),
          "cfop": "" (CFOP impresso na linha deste item, apenas números, ex.: 5102),
          "quantidade": "" (quantidade deste item),
          "valor_unitario": "", "valor_total": "" (valor total deste item),
          "bc_icms": "", "aliq_icms": "", "valor_icms": "",
          "bc_ipi": "", "aliq_ipi": "", "valor_ipi": ""
        }
      ]
    }
    Retorne um objeto dentro de "itens" para CADA linha de produto/serviço da nota, na mesma ordem em que
    aparecem na tabela de itens — nunca agrupe ou resuma itens diferentes em um só.`;

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

    // Aceita tanto o formato novo (itens) quanto, por segurança, um retorno antigo
    // de nota sem itens detalhados, caso a IA não consiga separar por algum motivo.
    const itens = (Array.isArray(dados.itens) && dados.itens.length > 0)
      ? dados.itens
      : [{
          codigo_item: '1', cfop: dados.cfop || '', quantidade: '1', valor_unitario: dados.valor_produtos,
          valor_total: dados.valor_produtos,
          bc_icms: dados.bc_icms, aliq_icms: dados.aliq_icms, valor_icms: dados.valor_icms,
          bc_ipi: dados.bc_ipi, aliq_ipi: dados.aliq_ipi, valor_ipi: dados.valor_ipi
        }];

    // MAPEAMENTO DAS 33 POSIÇÕES DO LEIAUTE DOMÍNIO - NOTAS FISCAIS DE SAÍDAS (1.1)
    // Uma linha por grupo de CFOP: os dados do cliente/nota se repetem em todas as linhas,
    // e os campos que são únicos da nota (desconto, PIS/COFINS) só entram na primeira linha,
    // pra não duplicar esse valor quando o Domínio somar todas as linhas da nota.
    itens.forEach((item, idx) => {
      const primeiraLinha = idx === 0;
      const valorTotalItem = paraNumero(item.valor_total);
      const descontoNota = primeiraLinha ? paraNumero(dados.valor_descontos) : 0;

      const linhaNova = [
        dados.cnpj_cliente || '',                               // 1. CPF/CNPJ
        dados.razao_cliente || '',                              // 2. Razão Social
        dados.uf_cliente || '',                                 // 3. UF
        dados.municipio_cliente || '',                          // 4. Município
        dados.endereco_cliente || '',                           // 5. Endereço
        dados.numero_nf || '',                                  // 6. Número Documento
        extrairSerieDaChave(dados.chave_acesso) || dados.serie || '1',  // 7. Série (conferida pela chave de acesso)
        dados.data_emissao || '',                               // 8. Data
        '0',                                                     // 9. Situação (0-Regular / 2-Cancelada)
        calcularAcumulador(item.cfop),                          // 10. Acumulador (101/102/201/202 conforme CFOP impresso)
        item.cfop || '',                                        // 11. CFOP (extraído da nota/DANFE, por item)
        valorTotalItem,                                         // 12. Valor Produtos (valor deste item)
        primeiraLinha ? paraNumero(dados.valor_descontos) : 0,  // 13. Valor Descontos (só na 1ª linha)
        valorTotalItem - descontoNota,                          // 14. Valor Contábil
        paraNumero(item.bc_icms),                               // 15. Base de Cálculo ICMS
        paraNumero(item.aliq_icms),                             // 16. Alíquota ICMS
        paraNumero(item.valor_icms),                            // 17. Valor ICMS
        0,                                                       // 18. Outras ICMS
        0,                                                       // 19. Isentas ICMS
        paraNumero(item.bc_ipi),                                // 20. Base de Cálculo IPI
        paraNumero(item.aliq_ipi),                              // 21. Alíquota IPI
        paraNumero(item.valor_ipi),                             // 22. Valor IPI
        0,                                                       // 23. Outras IPI
        0,                                                       // 24. Isentas IPI
        (item.codigo_item || '').toString().trim() || '1',      // 25. Código do Item (CÓDIGO PRODUTO da DANFE)
        paraNumero(item.quantidade) || 1,                       // 26. Quantidade
        paraNumero(item.valor_unitario),                        // 27. Valor Unitário
        primeiraLinha ? (dados.cst_piscofins || '99') : '',     // 28. CST PIS/COFINS (só na 1ª linha)
        primeiraLinha ? paraNumero(dados.bc_piscofins) : 0,     // 29. Base de Cálculo PIS/COFINS
        primeiraLinha ? paraNumero(dados.aliq_pis) : 0,         // 30. Alíquota PIS
        primeiraLinha ? paraNumero(dados.valor_pis) : 0,        // 31. Valor PIS
        primeiraLinha ? paraNumero(dados.aliq_cofins) : 0,      // 32. Alíquota COFINS
        primeiraLinha ? paraNumero(dados.valor_cofins) : 0      // 33. Valor COFINS
      ];

      sheet.appendRow(linhaNova);
    });

    return { ok: true, linhas: itens.length };

  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

// A chave de acesso da NF-e tem 44 dígitos fixos: UF(2) AAMM(4) CNPJ(14) modelo(2)
// SÉRIE(3) número(9) tpEmis(1) cNF(8) DV(1). A série real da nota são os 3 dígitos
// logo depois do modelo (posições 23 a 25) — mais confiável que ler o campo "Série"
// impresso, que às vezes sai borrado ou mal enquadrado no recorte da DANFE.
function extrairSerieDaChave(chaveAcesso) {
  const digitos = (chaveAcesso || '').toString().replace(/[^\d]/g, '');
  if (digitos.length !== 44) return '';
  const serieBruta = digitos.substring(22, 25);
  const serieNum = parseInt(serieBruta, 10);
  return isNaN(serieNum) ? '' : serieNum.toString();
}

function calcularAcumulador(cfopBruto) {
  const digitos = (cfopBruto || '').toString().replace(/[^\d]/g, '');
  if (digitos.length !== 4) return '';
  const mapaPrimeiroDigito = { '5': '1', '6': '2' };
  const primeiroDigito = mapaPrimeiroDigito[digitos[0]];
  const ultimoDigito = digitos[3];
  if (!primeiroDigito || (ultimoDigito !== '1' && ultimoDigito !== '2')) return '';
  return primeiroDigito + '0' + ultimoDigito;
}

// Converte "1.500,50", "1500,50" ou um número já pronto em number do JS.
function paraNumero(valor) {
  if (typeof valor === 'number') return valor;
  if (!valor) return 0;
  const limpo = valor.toString().trim().replace(/\./g, '').replace(',', '.');
  const numero = parseFloat(limpo);
  return isNaN(numero) ? 0 : numero;
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

// Tipo de cada uma das 33 colunas do leiaute (Entrada/Saída), na ordem exata.
// "INTEIRO": números sem casas decimais (Número Documento, Série, Situação,
// Acumulador, CFOP, Código do Item, Quantidade, CST PIS/COFINS).
// "DECIMAL": valores monetários/alíquotas, sempre com vírgula e 2 casas.
// "TEXTO": tudo o resto (CNPJ, nomes, UF, município, endereço).
const TIPOS_COLUNA = [
  'TEXTO', 'TEXTO', 'TEXTO', 'TEXTO', 'TEXTO',
  'INTEIRO', 'INTEIRO', 'DATA', 'INTEIRO', 'INTEIRO', 'INTEIRO',
  'DECIMAL', 'DECIMAL', 'DECIMAL',
  'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL',
  'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL',
  'TEXTO', 'INTEIRO', 'DECIMAL',
  'INTEIRO', 'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL', 'DECIMAL'
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

  const nomeArquivo = "Saidas_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") + ".txt";
  DriveApp.createFile(nomeArquivo, txtContent, MimeType.PLAIN_TEXT);

  ui.alert('Pronto!', `O arquivo foi gerado com sucesso no seu Google Drive com o nome:\n\n${nomeArquivo}\n\nFaça o download e importe na Escrita Fiscal do Domínio (Conjunto de Dados: 'Notas Fiscais de Saídas (Leiaute Domínio - Excel) (1.1)').`, ui.ButtonSet.OK);
}