/**
 * Abre a janela para escolher a aba que contém os participantes.
 */
function enviarConvites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ss.getSheets().map(function(sheet) {
    return sheet.getName();
  });

  const html = HtmlService.createTemplateFromFile("SheetPicker");
  html.abas = abas;

  const dialog = html.evaluate()
    .setWidth(440)
    .setHeight(280);

  SpreadsheetApp.getUi().showModalDialog(
    dialog,
    "Enviar convites do curso"
  );
}


/**
 * Mantém compatibilidade com o nome usado na versão de certificados.
 */
function enviarCertificados() {
  enviarConvites();
}


/**
 * Execute uma vez pelo editor do Apps Script para solicitar e validar
 * as permissões de acesso à planilha e de envio de e-mails.
 */
function autorizarServicos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "Este código deve estar vinculado à planilha que contém os participantes."
    );
  }

  const quotaRestante = MailApp.getRemainingDailyQuota();

  SpreadsheetApp.getUi().alert(
    "Permissões verificadas com sucesso. Cota disponível: " + quotaRestante
  );
}


function enviarConvitesComAba(nomeAba) {
  const lock = LockService.getDocumentLock();

  if (!lock || !lock.tryLock(1000)) {
    throw new Error(
      "Já existe um processo de envio em andamento. Aguarde a conclusão."
    );
  }

  try {
    // ---------------- CONFIGURAÇÕES ----------------

    const nomePrograma = "NOME DO PROGRAMA";
    const nomeTrilha = "NOME DA TRILHA OU CURSO";
    const linkGrupoWhatsApp = "LINK DO GRUPO DO WHATSAPP";
    const nomeEquipe = "NOME DA EQUIPE";
    const nomeInstituto = "NOME DA INSTITUIÇÃO";
    const MAXIMO_POR_EXECUCAO = 40;

    // ------------------------------------------------

    validarConfiguracoes_(linkGrupoWhatsApp);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilha = ss.getSheetByName(nomeAba);

    if (!planilha) {
      throw new Error("Aba não encontrada: " + nomeAba);
    }

    const ultimaLinha = planilha.getLastRow();
    let quotaRestante = MailApp.getRemainingDailyQuota();

    if (ultimaLinha < 2) {
      return criarResumo_(nomeAba, 0, 0, 0, 0, false, false, quotaRestante);
    }

    /*
     * Colunas esperadas:
     * A: Nome
     * B: E-mail
     * C: Status
     * D: Data do envio
     * E: Erro ou observação
     */
    const dados = planilha
      .getRange(2, 1, ultimaLinha - 1, 5)
      .getValues();

    const assuntoEmail =
      "Você foi selecionado(a) - " + nomePrograma;

    let enviados = 0;
    let erros = 0;
    let ignorados = 0;
    let limiteLoteAtingido = false;
    let quotaAtingida = false;

    for (let indice = 0; indice < dados.length; indice++) {
      const numeroLinha = indice + 2;
      const nome = String(dados[indice][0] || "").trim();
      const email = String(dados[indice][1] || "").trim();
      const status = normalizarTexto_(dados[indice][2] || "");

      if (status === "sim") {
        ignorados++;
        continue;
      }

      if (!nome || !email) {
        const mensagem = !nome ? "Nome não informado" : "E-mail não informado";
        planilha.getRange(numeroLinha, 5).setValue(mensagem);
        dados[indice][4] = mensagem;
        erros++;
        continue;
      }

      if (enviados >= MAXIMO_POR_EXECUCAO) {
        limiteLoteAtingido = true;
        break;
      }

      if (quotaRestante <= 0) {
        quotaAtingida = true;
        break;
      }

      if (!emailValido_(email)) {
        const mensagem = "Endereço de e-mail inválido";
        planilha.getRange(numeroLinha, 5).setValue(mensagem);
        dados[indice][4] = mensagem;
        erros++;
        continue;
      }

      try {
        const conteudo = criarMensagemConvite_({
          nome: nome,
          nomePrograma: nomePrograma,
          nomeTrilha: nomeTrilha,
          linkGrupoWhatsApp: linkGrupoWhatsApp,
          nomeEquipe: nomeEquipe,
          nomeInstituto: nomeInstituto
        });

        MailApp.sendEmail({
          to: email,
          subject: assuntoEmail,
          body: conteudo.texto,
          htmlBody: conteudo.html
        });

        const dataEnvio = new Date();
        planilha.getRange(numeroLinha, 3, 1, 3).setValues([
          ["SIM", dataEnvio, ""]
        ]);

        dados[indice][2] = "SIM";
        dados[indice][3] = dataEnvio;
        dados[indice][4] = "";
        enviados++;
        quotaRestante--;
      } catch (erro) {
        const mensagemErro = String(
          erro && erro.message ? erro.message : erro
        ).substring(0, 500);

        planilha.getRange(numeroLinha, 5).setValue(mensagemErro);
        dados[indice][4] = mensagemErro;
        erros++;
      }
    }

    const pendentes = dados.reduce(function(total, linha) {
      const nome = String(linha[0] || "").trim();
      const email = String(linha[1] || "").trim();
      const status = normalizarTexto_(linha[2] || "");
      return nome && email && status !== "sim" ? total + 1 : total;
    }, 0);

    SpreadsheetApp.flush();

    return criarResumo_(
      planilha.getName(),
      enviados,
      erros,
      ignorados,
      pendentes,
      limiteLoteAtingido,
      quotaAtingida,
      quotaRestante
    );
  } finally {
    lock.releaseLock();
  }
}


function criarMensagemConvite_(config) {
  const nomeHtml = escaparHtml_(config.nome);
  const programaHtml = escaparHtml_(config.nomePrograma);
  const trilhaHtml = escaparHtml_(config.nomeTrilha);
  const equipeHtml = escaparHtml_(config.nomeEquipe);
  const institutoHtml = escaparHtml_(config.nomeInstituto);
  const linkHtml = escaparHtml_(config.linkGrupoWhatsApp);

  const html =
    "<p>Olá, <strong>" + nomeHtml + "</strong>!</p>" +
    "<p>Temos o prazer de informar que você foi <strong>selecionado(a) " +
    "para participar da próxima turma do programa " + programaHtml +
    "</strong>, na trilha <strong>" + trilhaHtml + "</strong>. 🎉🤖</p>" +
    "<p><strong>Parabéns pela seleção!</strong> Em breve daremos início às " +
    "atividades da nova turma e realizaremos nossa <strong>aula inaugural</strong>, " +
    "na qual serão apresentadas as principais informações sobre a formação, " +
    "cronograma, acesso à plataforma e orientações para os estudos.</p>" +
    "<h3>📲 Próxima etapa: entrar no grupo do WhatsApp</h3>" +
    "<p>Para acompanhar todas as informações e avisos da turma, é necessário " +
    "entrar no nosso grupo do WhatsApp.</p>" +
    "<p><a href=\"" + linkHtml + "\" style=\"display:inline-block;" +
    "background:#25D366;color:#fff;padding:12px 18px;text-decoration:none;" +
    "border-radius:6px;font-weight:bold\">👉 Entrar no grupo do WhatsApp</a></p>" +
    "<p>Se o botão não funcionar, copie este endereço:<br>" +
    "<a href=\"" + linkHtml + "\">" + linkHtml + "</a></p>" +
    "<p>Todas as orientações sobre a formação serão compartilhadas por lá. " +
    "Por isso, fique atento(a) às mensagens do grupo.</p>" +
    "<p>Seja muito bem-vindo(a) ao programa! 🚀</p>" +
    "<p>Desejamos uma excelente jornada de aprendizado!</p>" +
    "<p>Atenciosamente,<br><strong>" + equipeHtml + "</strong><br>" +
    "<strong>" + institutoHtml + "</strong></p>";

  const texto =
    "Olá, " + config.nome + "!\n\n" +
    "Temos o prazer de informar que você foi selecionado(a) para participar " +
    "da próxima turma do programa " + config.nomePrograma + ", na trilha " +
    config.nomeTrilha + ". 🎉🤖\n\n" +
    "Parabéns pela seleção! Em breve daremos início às atividades da nova " +
    "turma e realizaremos nossa aula inaugural, na qual serão apresentadas " +
    "as principais informações sobre a formação, cronograma, acesso à " +
    "plataforma e orientações para os estudos.\n\n" +
    "PRÓXIMA ETAPA: ENTRAR NO GRUPO DO WHATSAPP\n\n" +
    "Para acompanhar todas as informações e avisos da turma, entre no grupo:\n" +
    config.linkGrupoWhatsApp + "\n\n" +
    "Todas as orientações sobre a formação serão compartilhadas por lá. " +
    "Por isso, fique atento(a) às mensagens do grupo.\n\n" +
    "Seja muito bem-vindo(a) ao programa! 🚀\n\n" +
    "Desejamos uma excelente jornada de aprendizado!\n\n" +
    "Atenciosamente,\n" + config.nomeEquipe + "\n" + config.nomeInstituto;

  return { html: html, texto: texto };
}


function validarConfiguracoes_(linkGrupoWhatsApp) {
  const link = String(linkGrupoWhatsApp || "").trim();

  if (
    !/^https:\/\/chat\.whatsapp\.com\/[A-Za-z0-9]+(?:\?.*)?$/i.test(link)
  ) {
    throw new Error(
      "Configure um link válido do grupo do WhatsApp em linkGrupoWhatsApp antes de enviar."
    );
  }
}


function criarResumo_(
  aba,
  enviados,
  erros,
  ignorados,
  pendentes,
  limiteLoteAtingido,
  quotaAtingida,
  quotaRestante
) {
  return {
    aba: aba,
    enviados: enviados,
    erros: erros,
    ignorados: ignorados,
    pendentes: pendentes,
    limiteLoteAtingido: limiteLoteAtingido,
    quotaAtingida: quotaAtingida,
    quotaRestante: quotaRestante
  };
}


function normalizarTexto_(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}


function emailValido_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}


function escaparHtml_(texto) {
  return String(texto || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
