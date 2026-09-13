// Gera o DANFSe (PDF) localmente a partir do XML da NFS-e.
//
// Por quê isso existe: a API do governo que gerava o DANFSe foi desativada
// em 01/07/2026 (ver aviso 🛑 em lib/nfse.js) — a partir daí, virou
// responsabilidade de cada sistema montar esse PDF, seguindo o layout da
// Nota Técnica nº 008 (SE/CGNFS-e, 05/05/2026, "Especificações Técnicas do
// DANFSe"). Essa nota técnica descreve o layout permitido, mas deixa
// explícito (item 2.1) que as posições exatas em cm (item 2.4.5) são
// "sugestão", não obrigatórias — o obrigatório é: papel A4, os blocos de
// campo previstos, e os tamanhos mínimos de fonte. Por isso este gerador
// monta os blocos de forma legível e completa, sem tentar replicar
// milimetricamente as coordenadas da nota técnica.
//
// ⚠️ Primeira versão — cobre os blocos principais (identificação,
// prestador, tomador, serviço, tributação municipal, tributação federal,
// valor total, informações complementares, QR Code, marca d'água de
// cancelada). NÃO cobre ainda: bloco de canhoto (opcional, item 2.1.13),
// tributação IBS/CBS (opcional em 2026, nenhuma das notas reais vistas até
// agora tinha esse grupo), destinatário/intermediário da operação (idem —
// sempre impressos como "não identificado", igual a nota técnica permite
// quando o grupo não existe no XML).
//
// ⚠️ Os valores de retenção federal (IRRF, contribuição previdenciária,
// CSLL) usam os mesmos nomes de tag já usados em lib/nfse.js
// (vRetIRRF/vRetCP/vRetCSLL) — ainda não confirmados contra uma nota real
// com esse tipo de retenção (só confirmamos retenção de ISSQN até agora).

import { PDFDocument, StandardFonts, rgb, degrees } from "pdf-lib";
import QRCode from "qrcode";
import { extrairValor, extrairBloco } from "./nfse.js";

const A4_LARGURA = 595.28;
const A4_ALTURA = 841.89;
const MARGEM = 36;

function formatarCNPJ(valor) {
  const d = (valor || "").replace(/\D/g, "");
  if (d.length !== 14) return valor || "-";
  return `${d.slice(0, 2)}.${d.slice(2, 5)}.${d.slice(5, 8)}/${d.slice(8, 12)}-${d.slice(12, 14)}`;
}

function formatarCEP(valor) {
  const d = (valor || "").replace(/\D/g, "");
  if (d.length !== 8) return valor || "-";
  return `${d.slice(0, 5)}-${d.slice(5, 8)}`;
}

function formatarDataHora(iso) {
  if (!iso) return "-";
  const data = new Date(iso);
  if (Number.isNaN(data.getTime())) return "-";
  const dd = String(data.getDate()).padStart(2, "0");
  const mm = String(data.getMonth() + 1).padStart(2, "0");
  const aaaa = data.getFullYear();
  const hh = String(data.getHours()).padStart(2, "0");
  const mi = String(data.getMinutes()).padStart(2, "0");
  return `${dd}/${mm}/${aaaa} ${hh}:${mi}`;
}

function formatarData(dataStr) {
  if (!dataStr) return "-";
  const [ano, mes, dia] = dataStr.split("-");
  return ano && mes && dia ? `${dia}/${mes}/${ano}` : dataStr;
}

function formatarMoeda(valor) {
  const n = Number(valor);
  if (!valor || Number.isNaN(n)) return "-";
  return n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Extrai os dados de uma "parte" (prestador ou tomador) já isolando o
// bloco certo — evita pegar endereço/telefone de quem não é essa parte.
// <prest> não repete endereço/nome (já estão em <emit>, ver aviso em
// lib/nfse.js), então cai pro <emit> como reserva quando faltar.
function extrairParte(xml, blocoTags, { blocoReserva } = {}) {
  const bloco = extrairBloco(xml, blocoTags);
  const reserva = blocoReserva ? extrairBloco(xml, blocoReserva) : "";

  function campo(tags) {
    const doBloco = extrairValor(bloco, tags);
    return doBloco || (reserva ? extrairValor(reserva, tags) : "");
  }

  const enderecoBloco = extrairBloco(bloco, ["end", "End"]) || bloco;
  const enderecoReserva = reserva ? extrairBloco(reserva, ["enderNac", "EnderNac"]) || reserva : "";

  function campoEndereco(tags) {
    const doBloco = extrairValor(enderecoBloco, tags);
    return doBloco || (enderecoReserva ? extrairValor(enderecoReserva, tags) : "");
  }

  return {
    cnpj: campo(["CNPJ", "Cnpj"]),
    im: campo(["IM"]),
    nome: campo(["xNome", "XNome", "razaoSocial", "RazaoSocial"]),
    telefone: campo(["fone"]),
    email: campo(["email"]),
    xLgr: campoEndereco(["xLgr"]),
    nro: campoEndereco(["nro"]),
    xBairro: campoEndereco(["xBairro"]),
    cMun: campoEndereco(["cMun", "xCidade"]),
    uf: campoEndereco(["UF"]),
    cep: campoEndereco(["CEP", "cEndPost"]),
  };
}

function enderecoLinha(parte) {
  const partes = [parte.xLgr, parte.nro, parte.xBairro].filter(Boolean);
  return partes.length ? partes.join(", ") : "-";
}

function municipioUf(parte) {
  if (!parte.cMun && !parte.uf) return "-";
  return [parte.cMun, parte.uf].filter(Boolean).join(" / ");
}

// Quebra um texto em linhas que cabem em `larguraMax` pontos, na fonte e
// tamanho dados — pdf-lib não quebra texto sozinho.
function quebrarLinhas(texto, fonte, tamanho, larguraMax) {
  const palavras = (texto || "-").split(/\s+/);
  const linhas = [];
  let linhaAtual = "";
  for (const palavra of palavras) {
    const tentativa = linhaAtual ? `${linhaAtual} ${palavra}` : palavra;
    if (fonte.widthOfTextAtSize(tentativa, tamanho) > larguraMax && linhaAtual) {
      linhas.push(linhaAtual);
      linhaAtual = palavra;
    } else {
      linhaAtual = tentativa;
    }
  }
  if (linhaAtual) linhas.push(linhaAtual);
  return linhas;
}

export async function gerarDanfsePdf(xml, { cancelada = false } = {}) {
  const pdfDoc = await PDFDocument.create();
  const pagina = pdfDoc.addPage([A4_LARGURA, A4_ALTURA]);
  const fonte = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fonteNegrito = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const larguraUtil = A4_LARGURA - MARGEM * 2;
  let y = A4_ALTURA - MARGEM;

  function texto(conteudo, x, tamanho, { negrito = false, cor = rgb(0, 0, 0) } = {}) {
    pagina.drawText(conteudo, { x, y, size: tamanho, font: negrito ? fonteNegrito : fonte, color: cor });
  }

  function linha(espessura = 0.5) {
    pagina.drawLine({
      start: { x: MARGEM, y },
      end: { x: A4_LARGURA - MARGEM, y },
      thickness: espessura,
      color: rgb(0.6, 0.6, 0.6),
    });
  }

  function tituloBloco(rotulo) {
    y -= 14;
    pagina.drawRectangle({ x: MARGEM, y: y - 2, width: larguraUtil, height: 14, color: rgb(0.92, 0.92, 0.92) });
    texto(rotulo.toUpperCase(), MARGEM + 4, 8, { negrito: true });
    y -= 14;
  }

  function campoValor(rotulo, valor, { x = MARGEM, largura = larguraUtil } = {}) {
    texto(rotulo, x, 6, { cor: rgb(0.4, 0.4, 0.4) });
    y -= 9;
    const linhas = quebrarLinhas(String(valor ?? "-"), fonte, 8, largura);
    for (const l of linhas) {
      texto(l, x, 8);
      y -= 10;
    }
  }

  function duasColunas(a, b) {
    const yInicial = y;
    const largura = larguraUtil / 2 - 8;
    campoValor(a.rotulo, a.valor, { largura });
    const yDepoisA = y;
    y = yInicial;
    campoValor(b.rotulo, b.valor, { x: MARGEM + larguraUtil / 2 + 8, largura });
    y = Math.min(y, yDepoisA);
  }

  // --- Extração dos dados do XML ---
  const idCompleto = (xml.match(/Id="([^"]+)"/) || [])[1] || "";
  const chaveAcesso = idCompleto.replace(/^NFS/, "");
  const numeroNota = extrairValor(xml, ["nNFSe"]);
  const dCompet = extrairValor(xml, ["dCompet"]);
  const dhProc = extrairValor(xml, ["dhProc"]);
  const dhEmi = extrairValor(xml, ["dhEmi"]);
  const nDPS = extrairValor(xml, ["nDPS"]);
  const serie = extrairValor(xml, ["serie"]);
  const cStat = extrairValor(xml, ["cStat"]);
  const tpAmb = extrairValor(xml, ["tpAmb"]);
  const xLocEmi = extrairValor(xml, ["xLocEmi"]);
  const xLocPrestacao = extrairValor(xml, ["xLocPrestacao"]);

  const blocoEmit = extrairBloco(xml, ["emit", "Emit"]);
  const prestador = extrairParte(xml, ["prest", "Prest"], { blocoReserva: ["emit", "Emit"] });
  if (!prestador.nome) prestador.nome = extrairValor(blocoEmit, ["xNome"]);
  const tomador = extrairParte(xml, ["toma", "Toma"]);

  const xTribNac = extrairValor(xml, ["xTribNac"]);
  const cTribNac = extrairValor(xml, ["cTribNac"]);
  const cTribMun = extrairValor(xml, ["cTribMun"]);
  const xNBS = extrairValor(xml, ["xNBS"]);
  const xDescServ = extrairValor(xml, ["xDescServ"]);

  const vServ = extrairValor(xml, ["vServ"]);
  const vDescIncond = extrairValor(xml, ["vDescIncond"]);
  const vDescCond = extrairValor(xml, ["vDescCondIncond"]);
  const vBC = extrairValor(xml, ["vBC"]);
  const pAliqAplic = extrairValor(xml, ["pAliqAplic"]);
  const tpRetISSQN = extrairValor(xml, ["tpRetISSQN"]);
  const vISSQN = extrairValor(xml, ["vISSQN"]);
  const vRetIRRF = extrairValor(xml, ["vRetIRRF"]);
  const vRetCP = extrairValor(xml, ["vRetCP"]);
  const vRetCSLL = extrairValor(xml, ["vRetCSLL"]);
  const vPIS = extrairValor(xml, ["vPIS"]);
  const vCOFINS = extrairValor(xml, ["vCOFINS"]);
  const vTotalRet = extrairValor(xml, ["vTotalRet"]);
  const vLiq = extrairValor(xml, ["vLiq"]);
  const pTotTribSN = extrairValor(xml, ["pTotTribSN"]);

  const chSubstda = extrairValor(xml, ["chSubstda"]);

  const homologacao = tpAmb === "2";
  const situacaoNota = cStat === "100" ? "Normal" : cStat ? `Código ${cStat}` : "-";
  const retencaoIssTexto = tpRetISSQN === "2" ? "Retido pelo tomador" : tpRetISSQN === "1" ? "Não retido" : "-";

  // --- Cabeçalho ---
  texto(xLocEmi || "-", A4_LARGURA - MARGEM - 150, 8, { cor: rgb(0.3, 0.3, 0.3) });
  texto("DANFSe", MARGEM, 16, { negrito: true });
  y -= 14;
  texto("Documento Auxiliar da NFS-e", MARGEM, 9);
  y -= 14;
  if (homologacao) {
    pagina.drawText("NFS-E SEM VALIDADE JURÍDICA", {
      x: MARGEM,
      y: y - 2,
      size: 10,
      font: fonteNegrito,
      color: rgb(0.7, 0, 0),
    });
    y -= 12;
  }
  y -= 8;
  linha();

  // --- Identificação da NFS-e ---
  tituloBloco("Identificação da NFS-e");
  campoValor("Chave de acesso", chaveAcesso);
  duasColunas(
    { rotulo: "Número da NFS-e", valor: numeroNota },
    { rotulo: "Competência", valor: formatarData(dCompet) }
  );
  duasColunas(
    { rotulo: "Data/hora emissão NFS-e", valor: formatarDataHora(dhProc) },
    { rotulo: "Número / série DPS", valor: `${nDPS || "-"} / ${serie || "-"}` }
  );
  duasColunas(
    { rotulo: "Data/hora emissão DPS", valor: formatarDataHora(dhEmi) },
    { rotulo: "Situação da NFS-e", valor: situacaoNota }
  );
  if (chSubstda) {
    campoValor("NFS-e substituída (chave de acesso)", chSubstda);
  }

  // --- Prestador ---
  tituloBloco("Prestador / Fornecedor");
  duasColunas(
    { rotulo: "CNPJ", valor: formatarCNPJ(prestador.cnpj) },
    { rotulo: "Inscrição Municipal", valor: prestador.im || "-" }
  );
  campoValor("Nome / Nome empresarial", prestador.nome);
  duasColunas(
    { rotulo: "Município / UF", valor: xLocPrestacao || municipioUf(prestador) },
    { rotulo: "CEP", valor: formatarCEP(prestador.cep) }
  );
  campoValor("Endereço", enderecoLinha(prestador));

  // --- Tomador ---
  tituloBloco("Tomador / Adquirente");
  if (!tomador.cnpj) {
    campoValor("", "TOMADOR/ADQUIRENTE DA OPERAÇÃO NÃO IDENTIFICADO NA NFS-e");
  } else {
    duasColunas(
      { rotulo: "CNPJ", valor: formatarCNPJ(tomador.cnpj) },
      { rotulo: "Inscrição Municipal", valor: tomador.im || "-" }
    );
    campoValor("Nome / Nome empresarial", tomador.nome);
    duasColunas(
      { rotulo: "Município / UF", valor: municipioUf(tomador) },
      { rotulo: "CEP", valor: formatarCEP(tomador.cep) }
    );
    campoValor("Endereço", enderecoLinha(tomador));
  }

  // --- Destinatário / Intermediário (não presentes nas notas vistas até agora) ---
  tituloBloco("Destinatário da operação");
  campoValor("", "DESTINATÁRIO DA OPERAÇÃO NÃO IDENTIFICADO NA NFS-e");

  // --- Serviço prestado ---
  tituloBloco("Serviço prestado");
  duasColunas(
    { rotulo: "Código de tributação nacional / municipal", valor: `${cTribNac || "-"} / ${cTribMun || "-"}` },
    { rotulo: "Código NBS", valor: xNBS || "-" }
  );
  campoValor("Descrição do código de tributação nacional", xTribNac);
  campoValor("Descrição do serviço", xDescServ);

  // --- Tributação municipal (ISSQN) ---
  tituloBloco("Tributação municipal (ISSQN)");
  duasColunas(
    { rotulo: "BC ISSQN", valor: `R$ ${formatarMoeda(vBC)}` },
    { rotulo: "Alíquota aplicada", valor: pAliqAplic ? `${pAliqAplic}%` : "-" }
  );
  duasColunas(
    { rotulo: "Retenção do ISSQN", valor: retencaoIssTexto },
    { rotulo: "ISSQN apurado", valor: `R$ ${formatarMoeda(vISSQN)}` }
  );

  // --- Tributação federal (exceto CBS) ---
  tituloBloco("Tributação federal (exceto CBS)");
  duasColunas(
    { rotulo: "IRRF", valor: `R$ ${formatarMoeda(vRetIRRF)}` },
    { rotulo: "Contribuição previdenciária retida", valor: `R$ ${formatarMoeda(vRetCP)}` }
  );
  duasColunas(
    { rotulo: "Contribuições sociais retidas (CSLL)", valor: `R$ ${formatarMoeda(vRetCSLL)}` },
    { rotulo: "PIS / COFINS débito apuração própria", valor: `R$ ${formatarMoeda(vPIS)} / R$ ${formatarMoeda(vCOFINS)}` }
  );

  // --- Valor total da NFS-e ---
  tituloBloco("Valor total da NFS-e");
  duasColunas(
    { rotulo: "Valor da operação / serviço", valor: `R$ ${formatarMoeda(vServ)}` },
    { rotulo: "Desconto incondicionado / condicionado", valor: `R$ ${formatarMoeda(vDescIncond)} / R$ ${formatarMoeda(vDescCond)}` }
  );
  duasColunas(
    { rotulo: "Total das retenções (ISSQN / federais)", valor: `R$ ${formatarMoeda(vTotalRet)}` },
    { rotulo: "Valor líquido da NFS-e", valor: `R$ ${formatarMoeda(vLiq)}` }
  );

  // --- Informações complementares ---
  tituloBloco("Informações complementares");
  const totaisAproximados = pTotTribSN
    ? `Totais Aproximados dos Tributos cfe. Lei nº 12.741/2012: ${pTotTribSN}% (Simples Nacional)`
    : "Totais Aproximados dos Tributos cfe. Lei nº 12.741/2012: -";
  campoValor("", totaisAproximados);

  // --- QR Code (consulta pública) ---
  if (chaveAcesso) {
    const urlConsulta = `https://www.nfse.gov.br/ConsultaPublica/?tpc=1&chave=${chaveAcesso}`;
    const qrPngBytes = await QRCode.toBuffer(urlConsulta, { width: 130, margin: 1 });
    const qrImagem = await pdfDoc.embedPng(qrPngBytes);
    const qrTamanho = 90;
    pagina.drawImage(qrImagem, {
      x: A4_LARGURA - MARGEM - qrTamanho,
      y: A4_ALTURA - MARGEM - qrTamanho - 4,
      width: qrTamanho,
      height: qrTamanho,
    });
    pagina.drawText("Consulte a autenticidade desta NFS-e", {
      x: A4_LARGURA - MARGEM - qrTamanho,
      y: A4_ALTURA - MARGEM - qrTamanho - 14,
      size: 6,
      font: fonte,
      color: rgb(0.4, 0.4, 0.4),
    });
  }

  // --- Marca d'água de cancelada ---
  if (cancelada) {
    pagina.drawText("CANCELADA", {
      x: A4_LARGURA / 2 - 150,
      y: A4_ALTURA / 2,
      size: 50,
      font: fonteNegrito,
      color: rgb(0.6, 0.6, 0.6),
      opacity: 0.5,
      rotate: degrees(35),
    });
  }

  const bytes = await pdfDoc.save();
  return Buffer.from(bytes);
}
