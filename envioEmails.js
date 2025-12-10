function enviarCertificados() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ultimaLinha = planilha.getLastRow();

  // -------- CONFIGURAÇÕES --------
  var pastaId = "ID da pasta do Google Drive onde estão os PDFs";
  var assuntoEmail = "Assunto do Email com o Certificado";
  var nomeCurso = "Nome do Curso";
  // -------------------------------

  var corpoEmail = "Olá {{NOME}},\n\nParabéns por concluir o curso " + nomeCurso + "!\n\nSegue em anexo o seu certificado de conclusão.\n\nAtenciosamente,\nEquipe de Cursos";

  var pasta = DriveApp.getFolderById(pastaId);
  var arquivos = pasta.getFiles();

  // Função para normalizar nomes
  function normalizar(texto) {
    return texto.toString()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  // Mapa de arquivos normalizados
  var mapaArquivos = {};
  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    var nomeArquivo = arquivo.getName().replace(".pdf", "");
    var nomeNormalizado = normalizar(nomeArquivo);
    mapaArquivos[nomeNormalizado] = arquivo;
  }

  // Loop pelos alunos
  for (var i = 2; i <= ultimaLinha; i++) {
    var nome = planilha.getRange(i, 1).getValue();
    var email = planilha.getRange(i, 2).getValue();

    if (!nome || !email) continue;

    var nomeNormalizado = normalizar(nome);
    var arquivoPdf = mapaArquivos[nomeNormalizado];

    if (!arquivoPdf) {
      Logger.log("❌ Arquivo não encontrado para: " + nome);
      continue;
    }

    var corpoPersonalizado = corpoEmail.replace("{{NOME}}", nome);

    GmailApp.sendEmail(email, assuntoEmail, corpoPersonalizado, {
      attachments: [arquivoPdf.getBlob()]
    });

    Logger.log("✅ Enviado para: " + email);
  }

  SpreadsheetApp.getUi().alert("Envio de certificados finalizado!");
}
