function enviarCertificados() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var ultimaLinha = planilha.getLastRow();

  // -------- CONFIGURAÇÕES --------
  var pastaId = "ID da pasta do google drive com os certificados";
  var nomeCurso = "Nome do curso";
  var assuntoEmail = "Certificado de Participação - " + nomeCurso;
  var LIMITE_LOTE = 40;
  // -------------------------------

  // Confirmação (preview simples)
  var resposta = ui.alert(
    "Confirmar envio",
    "Deseja enviar os certificados agora?\nCurso: " + nomeCurso,
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) {
    ui.alert("Envio cancelado.");
    return;
  }

  // Corpo do e-mail (HTML seguro)
  var corpoEmailHTML =
    "<p>Olá, <strong>{{NOME}}</strong>!</p>" +
    "<p>É com satisfação que enviamos em anexo o seu <strong>Certificado de Participação</strong> no curso <strong>" + nomeCurso + "</strong>.</p>" +
    "<p>Parabenizamos pelo seu empenho e dedicação durante o curso. Esperamos que os conhecimentos adquiridos possam contribuir para sua formação e abrir novas possibilidades no campo da robótica e da tecnologia.</p>" +
    "<p>Caso tenha alguma dúvida ou precise de mais informações, estamos a disposição pelo contato: <strong>(00) 00000-0000</strong>.</p>" +
    "<p>Atenciosamente,<br>Equipe ....</p>";

  var pasta = DriveApp.getFolderById(pastaId);
  var arquivos = pasta.getFiles();

  // Normalização
  function normalizar(texto) {
    return texto.toString()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  // Mapa de certificados
  var mapaArquivos = {};
  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    mapaArquivos[normalizar(arquivo.getName().replace(".pdf", ""))] = arquivo;
  }

  var enviadosNoLote = 0;

  // Loop principal
  for (var i = 2; i <= ultimaLinha; i++) {
    var nome = planilha.getRange(i, 1).getValue();
    var email = planilha.getRange(i, 2).getValue();
    var status = planilha.getRange(i, 3).getValue();

    if (!nome || !email || status === "SIM") continue;

    try {
      var arquivoPdf = mapaArquivos[normalizar(nome)];

      if (!arquivoPdf) {
        planilha.getRange(i, 5).setValue("Certificado nao encontrado");
        continue;
      }

      var corpoPersonalizado = corpoEmailHTML.replace("{{NOME}}", nome);

      GmailApp.sendEmail(email, assuntoEmail, "", {
        htmlBody: corpoPersonalizado,
        attachments: [arquivoPdf.getBlob()]
      });

      planilha.getRange(i, 3).setValue("SIM");
      planilha.getRange(i, 4).setValue(new Date());
      planilha.getRange(i, 5).setValue("");

      enviadosNoLote++;

      // Controle de lote
      if (enviadosNoLote >= LIMITE_LOTE) {
        Utilities.sleep(30000); // 30 segundos
        enviadosNoLote = 0;
      }

    } catch (erro) {
      planilha.getRange(i, 5).setValue(erro.toString());
    }
  }

  ui.alert("Envio finalizado com sucesso!");
}
