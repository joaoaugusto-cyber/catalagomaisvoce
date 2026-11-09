/**
 * Recebe as fichas de avaliação do catálogo Mais Você e grava na planilha.
 * As assinaturas viram imagens PNG numa pasta do Drive; a planilha guarda o link.
 *
 * Como instalar: veja planilha/COMO-INSTALAR.md
 */

var ABA = 'Fichas';
var PASTA_ASSINATURAS = 'Mais Você - Assinaturas das fichas';
var CABECALHO = [
  'Recebida em', 'Nome', 'CPF', 'Nascimento', 'Telefone', 'E-mail', 'CEP', 'Endereço', 'Sexo',
  'Áreas de interesse', 'Serviços', 'Direito de imagem', 'Assinatura (imagem)',
  'Termo de ciência', 'Assinatura (termo)', 'Data na ficha'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.site) return resposta({ ok: true }); // campo-isca preenchido: robô, ignora
    if (!d.nome || !d.cpf) return resposta({ ok: false, erro: 'dados incompletos' });

    var aba = abaFichas();
    var pasta = pastaAssinaturas();
    var base = (d.nome + ' ' + d.cpf).replace(/[^\wÀ-ú .-]/g, '');
    var linkImg = salvarAssinatura(pasta, d.assinaturaImagem, base + ' - direito de imagem');
    var linkTermo = salvarAssinatura(pasta, d.assinaturaTermo, base + ' - termo de ciência');

    aba.appendRow([
      new Date(), d.nome, d.cpf, d.nascimento, d.telefone, d.email, d.cep, d.endereco, d.sexo,
      d.areas, d.servicos, d.direitoImagem, linkImg,
      d.termo ? 'Ciente e de acordo' : '', linkTermo, d.data
    ]);
    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resposta({ ok: true, servico: 'fichas Mais Você' });
}

function abaFichas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
  if (aba.getLastRow() === 0) {
    aba.appendRow(CABECALHO);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold').setBackground('#3d8f74').setFontColor('#ffffff');
    aba.getRange('C:C').setNumberFormat('@'); // CPF como texto
  }
  return aba;
}

function pastaAssinaturas() {
  var it = DriveApp.getFoldersByName(PASTA_ASSINATURAS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_ASSINATURAS);
}

function salvarAssinatura(pasta, dataUrl, nome) {
  if (!dataUrl || dataUrl.indexOf('base64,') < 0) return '';
  var bytes = Utilities.base64Decode(dataUrl.split('base64,')[1]);
  var arq = pasta.createFile(Utilities.newBlob(bytes, 'image/png', nome + '.png'));
  return arq.getUrl();
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Rode uma vez pelo editor para autorizar e criar a aba e a pasta. */
function instalar() {
  abaFichas();
  pastaAssinaturas();
}
