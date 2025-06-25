function doGet() {
  return HtmlService.createHtmlOutputFromFile("Form")
    .setTitle("Gerador de Documentos")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function iniciarContador(total) {
  const cache = CacheService.getUserCache();
  cache.put("restante", total.toString(), 3600); // expira em 1 hora
}

function atualizarContador(quantos) {
  const cache = CacheService.getUserCache();
  let atual = parseInt(cache.get("restante")) || 0;
  atual -= quantos;
  if (atual < 0) atual = 0;
  cache.put("restante", atual.toString(), 3600);
}

function consultarContador() {
  const cache = CacheService.getUserCache();
  return cache.get("restante") || "0";
}

function gerarMultiplosDocumentos(nomes, dadosBase) {
  try {
    const modelos = {
      certificado_capacitacao: {
        idModelo: "1uj_VB0LHhGI4wrEvUm3uBgJG5lD_VmBsQbdIBi92j_w",
        pastaId: "1_pWVXznAdmslSjfeB4OzxBw-S6atGtUa",
        tipo: "slide"
      },
      certificado_certificacao: {
        idModelo: "1lfZnEony9zOb300_i2ohZMLfFLJeri0z6fnP_A4LdVU",
        pastaId: "1W_amsv3MsEG33j9BBB5QpodEMypJSGMM",
        tipo: "slide"
      },
      declaracao_carga_horaria: {
        idModelo: "1zIKlEziOKvDadwJCzemoQy1IGMFOju-wkhjFOdEMCvw",
        pastaId: "14dlEmW26bHbxbpPrMJhLGaW7uBr2bB-S",
        tipo: "doc"
      },
      certificado_inovacao: {
        idModelo: "12AJSlRC-osy6Mj3SVgqJcSbuP-XTZ-joDQ8_-c_U4HM",
        pastaId: "1VOymqXMbQ6-tSCNvjsDYiiDIGXDXxR4Y",
        tipo: "slide"
      }
    };

    const modelo = modelos[dadosBase.tipo];
    if (!modelo) throw new Error("Tipo de documento inválido.");

    const conteudoFormatado = (dadosBase.conteudo || "")
      .split("\n")
      .map(item => item.trim())
      .filter(item => item)
      .map(item => "• " + item)
      .join("\n");

    nomes.forEach(nome => {
      const substituicoes = {
        "{{NOME}}": nome,
        "{{CAPACITACAO}}": dadosBase.capacitacao || "",
        "{{CERTIFICACAO}}": dadosBase.certificacao || "",
        "{{EVENTO}}": dadosBase.evento || "",
        "{{CARGAHORARIA}}": dadosBase.cargaHoraria || "",
        "{{DATA}}": dadosBase.data || "",
        "{{PERIODO}}": dadosBase.periodo || "",
        "{{CONTEUDO}}": conteudoFormatado,
        "{{TURNO}}": dadosBase.turno || "",
        "{{HORARIO}}": dadosBase.horario || "",
        "{{PROFESSOR}}": dadosBase.professor || ""
      };

      const nomeArquivo = nome.trim();
      const copia = DriveApp.getFileById(modelo.idModelo).makeCopy(nomeArquivo);
      const pastaDestino = DriveApp.getFolderById(modelo.pastaId);

      if (modelo.tipo === "slide") {
        const slideDoc = SlidesApp.openById(copia.getId());
        slideDoc.getSlides().forEach(slide => {
          slide.getShapes().forEach(shape => {
            if (shape.getText) {
              const text = shape.getText();
              for (let chave in substituicoes) {
                text.replaceAllText(chave, substituicoes[chave]);
              }
            }
          });
        });
        slideDoc.saveAndClose();
      } else {
        const doc = DocumentApp.openById(copia.getId());
        const body = doc.getBody();
        for (let chave in substituicoes) {
          body.replaceText(chave, substituicoes[chave]);
        }
        doc.saveAndClose();
      }

      const pdf = DriveApp.getFileById(copia.getId()).getAs(MimeType.PDF).setName(nomeArquivo + ".pdf");
      pastaDestino.createFile(pdf);
      DriveApp.getFileById(copia.getId()).setTrashed(true);
    });

    atualizarContador(nomes.length);
    return `✅ ${nomes.length} documento(s) gerado(s) com sucesso.`;
  } catch (erro) {
    return "❌ Erro: " + erro.message;
  }
}
