// Integração com a API Nacional da NFS-e (ADN — Ambiente de Dados Nacional).
//
// ✅ Caminhos confirmados em 2026-09-10 a partir do Swagger oficial
// (baixado direto de https://adn.producaorestrita.nfse.gov.br/contribuintes/docs/index.html
// com um certificado real). A API só documenta dois caminhos, os dois sob
// o prefixo /contribuintes:
//   GET /DFe/{NSU}?cnpjConsulta=...   -> lote de documentos a partir do NSU
//   GET /NFSe/{ChaveAcesso}/Eventos    -> eventos (cancelamento etc) de uma
//                                         nota específica — não é o XML da
//                                         nota em si, por isso não usamos
//                                         esse caminho pra buscar o documento.
// A resposta de /DFe/{NSU} é um objeto (LoteDistribuicaoNSUResponse), não uma
// lista solta: { StatusProcessamento, LoteDFe: [...], Erros: [...], ... }.
// Cada item de LoteDFe já vem com o XML embutido em ArquivoXml (gzip+base64)
// quando TipoDocumento é "NFSE" — não existe um segundo endpoint pra buscar
// o XML separadamente, então se ArquivoXml não vier, a nota é ignorada.
// Os nomes dos campos vêm em PascalCase (NSU, ChaveAcesso, ArquivoXml,
// TipoDocumento, StatusProcessamento) — reparem nisso se for mexer aqui.
//
// ✅ DANFSe (PDF) — a API do governo que gerava esse PDF
// (https://adn.nfse.gov.br/danfse/docs/index.html) foi desativada em
// 1º de julho de 2026 (confirmado na Nota Técnica nº 008, SE/CGNFS-e,
// 05/05/2026): a partir daí virou responsabilidade de cada software gerar
// o DANFSe localmente. É o que lib/danfse.js faz — ver os avisos lá pra
// saber o que já está coberto e o que ainda falta (ex: bloco de canhoto,
// tributação IBS/CBS).
//
// ⚠️ Busca por período é feita varrendo o NSU (não existe filtro de data
// pronto na API) e filtrando cada nota pela COMPETÊNCIA (dCompet) dela —
// não pela data de emissão, porque uma nota de substituição pode ser
// emitida num mês e valer pra competência de outro (decidido com a
// contabilidade em 2026-09-12). Pode ficar lento pra CNPJs com muito
// histórico.
//
// ⚠️ TERCEIRO PONTO IMPORTANTE — nomes de campo no XML:
// Ver o comentário em CAMPOS_XML (mais abaixo) pro que já foi confirmado
// contra uma NFS-e real em 2026-09-10 e o que ainda falta confirmar
// (principalmente os valores de retenção federal — a nota real recebida
// não tinha nenhuma retenção pra comparar).

// ⚠️ QUARTO PONTO IMPORTANTE — leitura do .pfx:
// Versões recentes do Node.js (OpenSSL 3) recusam abrir diretamente muitos
// arquivos .pfx de certificado brasileiro, porque o "envelope" de
// criptografia do arquivo usa um algoritmo que o OpenSSL 3 desativou por
// padrão (isso não tem nada a ver com o certificado estar errado ou
// vencido). Por isso lemos o .pfx com a biblioteca `node-forge`, que faz
// essa leitura sozinha, sem depender dessa parte restrita do Node — e
// funciona igual tanto local quanto hospedado.

import https from "https";
import zlib from "zlib";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import forge from "node-forge";
import { gerarDanfsePdf } from "./danfse.js";

const AMBIENTES = {
  homologacao: "https://adn.producaorestrita.nfse.gov.br",
  producao: "https://adn.nfse.gov.br",
};

const PATHS = {
  // GET -> lote de documentos fiscais a partir de um NSU, filtrado pelo
  // CNPJ consultado. Confirmado no Swagger oficial (ver aviso no topo).
  dfe: (nsu, cnpjConsulta) => {
    const params = new URLSearchParams({ cnpjConsulta });
    return `/contribuintes/DFe/${nsu}?${params.toString()}`;
  },
};

// Nomes de tag confirmados em 2026-09-10 com uma NFS-e real (layout NFS-e
// Nacional v1.01, xmlns sped.fazenda.gov.br/nfse) e cruzados com a tabela
// de campos da Nota Técnica nº 008 (DANFSe), que lista o "Caminho no XML"
// de cada campo. dhEmi, vServ e nNFSe já bateram certinho na nota real
// recebida. Os campos de retenção federal ficam dentro de
// NFSe/infNFSe/DPS/infDPS/valores/trib/ — tribMun (vISSQN, tpRetISSQN) e
// tribFed (vRetIRRF, vRetCSLL, vPIS, vCOFINS). A nota real que recebemos
// era do Simples Nacional sem nenhuma retenção, então esse grupo nem
// aparecia no XML (o que é esperado quando não há retenção) — ainda falta
// confirmar esses nomes contra uma nota com retenção de verdade.
const CAMPOS_XML = {
  numeroNota: ["nNFSe", "NumeroNfse", "numeroNfse"],
  dataEmissao: ["dhEmi", "DataEmissao", "dataEmissao"],
  // Mês/ano que o serviço se refere (pode ser diferente do mês em que a
  // nota foi de fato emitida — ex: nota de substituição emitida em junho
  // pra competência de maio). Decidido com a contabilidade em 2026-09-12:
  // o período da busca filtra por competência, não por emissão, porque é
  // isso que importa pra apuração — senão nota de substituição tardia
  // pode ficar de fora do mês certo.
  dataCompetencia: ["dCompet", "DataCompetencia", "dataCompetencia"],
  valorServico: ["vServ", "ValorServicos", "valorServico"],
  // ✅ Confirmado em 2026-09-12 comparando duas notas reais (uma com
  // retenção de ISS, outra sem): vISSQN só existe no XML quando HÁ
  // retenção — nesse caso ele vem sempre igual a vTotalRet (a retenção é
  // de 100% do ISSQN apurado, não parcial). Quando não há retenção, a tag
  // nem aparece. Ou seja "ISSQN" e "ISS retido" saírem iguais na planilha
  // não é bug — é o esperado sempre que houver retenção.
  issqn: ["vISSQN", "ValorIssqn"],
  // ✅ Confirmado com o mesmo par de notas: tpRetISSQN="1" na nota sem
  // retenção, "2" na nota com retenção (e com vISSQN presente). O leiaute
  // prevê uma 3ª opção que ainda não vimos na prática — tratamos qualquer
  // valor diferente de "2" como "não retido", o que é o lado seguro.
  tpRetISSQN: ["tpRetISSQN"],
  pis: ["vPIS", "ValorPis"],
  cofins: ["vCOFINS", "ValorCofins"],
  csll: ["vRetCSLL", "vCSLL", "ValorCsll"],
  ir: ["vRetIRRF", "vIR", "ValorIr"],
  // 🛑 cStat: só temos "100" (nota normal) confirmado com segurança. Uma
  // nota real de SUBSTITUIÇÃO (válida, não cancelada — tinha bloco <subst>
  // com o motivo da troca) veio com cStat="101", o que derrubou a hipótese
  // de "qualquer valor ≠ 100 é cancelada". Por isso cStat não decide mais
  // cancelamento sozinho (ver EVENTOS_DE_CANCELAMENTO em buscarNotas, que
  // usa o TipoEvento confirmado no Swagger em vez disso) — fica guardado
  // aqui só pra referência/depuração futura.
  cStat: ["cStat"],
};

// TipoEvento (enum confirmado no Swagger oficial) que realmente significam
// "a nota está cancelada" — de propósito NÃO inclui
// SOLICITACAO_CANCELAMENTO_ANALISE_FISCAL (só um pedido, ainda não
// decidido) nem CANCELAMENTO_INDEFERIDO_ANALISE_FISCAL (pedido negado —
// a nota continua válida).
const EVENTOS_DE_CANCELAMENTO = new Set([
  "CANCELAMENTO",
  "CANCELAMENTO_POR_SUBSTITUICAO",
  "CANCELAMENTO_DEFERIDO_ANALISE_FISCAL",
  "CANCELAMENTO_POR_OFICIO",
]);

function getBaseUrl() {
  const ambiente = process.env.NFSE_AMBIENTE === "producao" ? "producao" : "homologacao";
  return AMBIENTES[ambiente];
}

// Lê o .pfx (com node-forge, ver aviso 4 no topo) e devolve a chave
// privada e o(s) certificado(s) em formato PEM — é isso que o Node usa
// pra fazer a conexão mTLS de verdade (options.key / options.cert),
// substituindo o pfx+passphrase que o Node não consegue abrir sozinho.
function lerCertificado(pfxBuffer, senha) {
  let p12;
  try {
    const p12Asn1 = forge.asn1.fromDer(forge.util.createBuffer(pfxBuffer.toString("binary")));
    p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, senha);
  } catch (err) {
    throw new Error(
      "Não foi possível abrir o certificado — confira se a senha está certa e se o arquivo é um .pfx/.p12 válido."
    );
  }

  // A chave privada pode vir em pkcs8ShroudedKeyBag (mais comum, protegida
  // por senha) ou em keyBag (sem proteção extra) — tentamos os dois.
  let chavePrivada = null;
  const shroudedBags = p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag });
  const shroudedList = shroudedBags[forge.pki.oids.pkcs8ShroudedKeyBag];
  if (shroudedList && shroudedList[0]) chavePrivada = shroudedList[0].key;

  if (!chavePrivada) {
    const keyBags = p12.getBags({ bagType: forge.pki.oids.keyBag });
    const keyList = keyBags[forge.pki.oids.keyBag];
    if (keyList && keyList[0]) chavePrivada = keyList[0].key;
  }

  const certBags = p12.getBags({ bagType: forge.pki.oids.certBag });
  const certList = certBags[forge.pki.oids.certBag] || [];

  if (!chavePrivada || certList.length === 0) {
    throw new Error("O arquivo .pfx não parece conter uma chave privada e um certificado válidos.");
  }

  const chavePem = forge.pki.privateKeyToPem(chavePrivada);
  // Junta o certificado (e qualquer intermediário que esteja dentro do
  // mesmo .pfx) numa cadeia só, na ordem em que vieram no arquivo.
  const certPem = certList.map((bag) => forge.pki.certificateToPem(bag.cert)).join("\n");

  return { chavePem, certPem };
}

// Mensagens amigáveis pros códigos de erro que o governo já devolveu de
// verdade pra gente — sem isso, o erro aparecia como o JSON cru inteiro
// na tela. Complementa (não substitui) a descrição que vier no corpo.
const MENSAGENS_ERRO_GOVERNO = {
  // Testado em 2026-09-16: acontece quando o CNPJ digitado no formulário
  // não é o mesmo CNPJ (base) do certificado anexado — o governo só deixa
  // consultar o próprio CNPJ do certificado.
  E2243: "O CNPJ digitado não é o mesmo CNPJ do certificado anexado. Confirme se o certificado é da mesma empresa do CNPJ informado.",
};

// Faz uma requisição GET autenticada com o certificado do cliente (mTLS).
function requisitarComCertificado({ chavePem, certPem, path, accept = "application/json" }) {
  return new Promise((resolve, reject) => {
    const baseUrl = new URL(getBaseUrl());
    const options = {
      hostname: baseUrl.hostname,
      path,
      method: "GET",
      key: chavePem,
      cert: certPem,
      headers: { Accept: accept },
      // nunca desligue isso: garante que estamos falando com o servidor
      // certo (e não um golpe/man-in-the-middle).
      rejectUnauthorized: true,
      // agent: false força uma conexão TCP/TLS totalmente nova a cada
      // chamada, sem reaproveitar nada de uma requisição anterior — evita
      // um tipo de erro de "decryption failed / bad record mac" que pode
      // acontecer quando alguma coisa no meio do caminho (ex: antivírus
      // com inspeção HTTPS) se confunde tentando reaproveitar a conexão.
      agent: false,
      // exige pelo menos TLS 1.2 (nunca versões antigas e inseguras); não
      // trava numa versão exata, só estabelece um piso.
      minVersion: "TLSv1.2",
    };

    const req = https.request(options, (res) => {
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => {
        const body = Buffer.concat(chunks);
        // A API do governo responde 404 (em vez de 200 com lista vazia)
        // quando não existe mais nenhum documento a partir do NSU pedido
        // (StatusProcessamento: "NENHUM_DOCUMENTO_LOCALIZADO", código
        // E2220) — confirmado testando de verdade em 2026-09-10. Isso não
        // é erro, é "acabaram as notas", então tratamos como sucesso e
        // deixamos o JSON do corpo (que já tem StatusProcessamento/Erros)
        // ser interpretado por quem chamou.
        const semDocumentoNoNsu = res.statusCode === 404 && accept === "application/json";
        if ((res.statusCode >= 200 && res.statusCode < 300) || semDocumentoNoNsu) {
          resolve({ statusCode: res.statusCode, body });
        } else {
          const textoBody = body.toString("utf-8");
          let mensagemAmigavel = null;
          try {
            const erros = JSON.parse(textoBody).Erros || [];
            if (erros.length > 0) {
              mensagemAmigavel = erros
                .map((e) => MENSAGENS_ERRO_GOVERNO[e.Codigo] || e.Descricao || e.Codigo)
                .join(" ");
            }
          } catch {
            // corpo não veio em JSON — usa o texto cru mesmo, tratado abaixo.
          }
          reject(
            new Error(
              mensagemAmigavel || `Governo respondeu ${res.statusCode} em ${path}: ${textoBody.slice(0, 500)}`
            )
          );
        }
      });
    });

    req.on("error", (err) => {
      // Erro de certificado (mTLS) costuma aparecer aqui, ex:
      // "unable to verify the first certificate" ou certificado vencido.
      reject(err);
    });

    req.end();
  });
}

function descompactarXml(base64Gzip) {
  const comprimido = Buffer.from(base64Gzip, "base64");
  return zlib.gunzipSync(comprimido).toString("utf-8");
}

// Extrai o conteúdo de uma tag simples (<tag>valor</tag>), tentando cada
// nome candidato em ordem até achar um que exista. Exportada pra
// lib/danfse.js reaproveitar (mesma lógica, sem duplicar regex).
export function extrairValor(xmlOuBloco, tags) {
  for (const tag of tags) {
    const m = xmlOuBloco.match(new RegExp(`<${tag}[^>]*>([^<]*)</${tag}>`, "i"));
    if (m) return m[1].trim();
  }
  return "";
}

// Extrai o bloco inteiro de uma tag composta (<tag>...conteúdo...</tag>),
// usado pra isolar <prest>...</prest> de <toma>...</toma> antes de procurar
// o CNPJ dentro — senão um regex simples pegaria o CNPJ errado (o primeiro
// que aparecer no XML inteiro).
export function extrairBloco(xml, tags) {
  for (const tag of tags) {
    const m = xml.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`, "i"));
    if (m) return m[1];
  }
  return "";
}

// A partir do XML de uma nota, descobre: se o CNPJ do cliente é o
// prestador (nota emitida) ou o tomador (nota tomada), a data de emissão,
// e os valores de retenção. Devolve null se não conseguir classificar a
// nota com segurança (melhor ignorar do que colocar na planilha errada).
function extrairDadosNota(xml, cnpjCliente) {
  // Confirmado em 2026-09-10 com uma NFS-e real: os blocos são <prest> e
  // <toma> (dentro de NFSe/infNFSe/DPS/infDPS/), não <prestador>/<tomador>.
  const blocoPrestador = extrairBloco(xml, ["prest", "Prest"]);
  const blocoTomador = extrairBloco(xml, ["toma", "Toma"]);
  const blocoEmitente = extrairBloco(xml, ["emit", "Emit"]);

  const cnpjPrestador = extrairValor(blocoPrestador, ["CNPJ", "Cnpj"]).replace(/\D/g, "");
  const cnpjTomador = extrairValor(blocoTomador, ["CNPJ", "Cnpj"]).replace(/\D/g, "");
  const cnpjLimpo = (cnpjCliente || "").replace(/\D/g, "");

  let papel = null;
  if (cnpjPrestador && cnpjPrestador === cnpjLimpo) papel = "emitidas";
  else if (cnpjTomador && cnpjTomador === cnpjLimpo) papel = "tomadas";
  if (!papel) return null;

  const dataStr = extrairValor(xml, CAMPOS_XML.dataEmissao);
  const dataEmissao = dataStr ? new Date(dataStr) : null;

  // dCompet vem só como "AAAA-MM-DD" (sem hora) — string assim o
  // JavaScript já interpreta como meia-noite UTC, então não tem ambiguidade
  // de fuso horário aqui (ao contrário de dhEmi, que vem com hora e fuso).
  const dCompetStr = extrairValor(xml, CAMPOS_XML.dataCompetencia);
  const dataCompetencia = dCompetStr ? new Date(dCompetStr) : null;

  const blocoOutraParte = papel === "emitidas" ? blocoTomador : blocoPrestador;
  const nomesCandidatos = ["xNome", "XNome", "razaoSocial", "RazaoSocial"];
  // <prest> às vezes não repete o xNome (já está em <emit>, no topo do
  // XML) — se não achar em <prest>, cai pro <emit> como reserva.
  const nomeOutraParte =
    extrairValor(blocoOutraParte, nomesCandidatos) ||
    (papel === "tomadas" ? extrairValor(blocoEmitente, nomesCandidatos) : "");

  // ISSQN é um valor só — a diferença entre "ISSQN" (próprio) e "ISS
  // retido" é só QUEM paga, não dois valores diferentes (explicado pelo
  // fiscal em 2026-09-12, olhando o campo "Retenção de ISSQN" do DANFSe:
  // "Não retido" = ISS próprio; "Retido pelo tomador" = ISS retido). Por
  // isso as duas colunas são mutuamente exclusivas: o valor vai pra uma ou
  // pra outra, nunca pras duas ao mesmo tempo.
  const issqnBruto = extrairValor(xml, CAMPOS_XML.issqn);
  const tpRetISSQN = extrairValor(xml, CAMPOS_XML.tpRetISSQN);
  const retido = tpRetISSQN === "2";

  return {
    papel, // "emitidas" | "tomadas"
    numeroNota: extrairValor(xml, CAMPOS_XML.numeroNota),
    dataEmissao,
    dataCompetencia,
    nomeOutraParte,
    cnpjOutraParte: papel === "emitidas" ? cnpjTomador : cnpjPrestador,
    valorServico: extrairValor(xml, CAMPOS_XML.valorServico),
    // ISSQN próprio (não retido) — em branco quando a nota tem retenção.
    issqn: retido ? "" : issqnBruto,
    // ISS retido pelo tomador — em branco quando a nota não tem retenção.
    issRetido: retido ? issqnBruto : "",
    pis: extrairValor(xml, CAMPOS_XML.pis),
    cofins: extrairValor(xml, CAMPOS_XML.cofins),
    csll: extrairValor(xml, CAMPOS_XML.csll),
    ir: extrairValor(xml, CAMPOS_XML.ir),
    // cStat não decide mais cancelamento sozinho (ver aviso 🛑 em
    // CAMPOS_XML.cStat) — quem marca cancelada é só o evento de
    // cancelamento confirmado, em buscarNotas.
    cancelada: false,
  };
}

// Layout de colunas definido junto com o setor fiscal (planilha de
// referência deles): N° Nota, Chave de acesso, Data da emissão,
// Tomador/Prestador, CNPJ, Valor do serviço, ISSQN, ISS retido, PIS,
// COFINS, CSLL, IRRF (o IRRF voltou a pedido em 2026-09-16 — tinha sido
// tirado antes só pra bater com a planilha de referência, que não tinha
// essa coluna, mas o valor sempre foi extraído em extrairDadosNota).
function linhaPlanilha(doc, colunaOutraParte) {
  return {
    "N° Nota": doc.dados.numeroNota || "",
    "Chave de acesso": doc.chaveAcesso,
    "Data de emissão": doc.dados.dataEmissao ? doc.dados.dataEmissao.toISOString().slice(0, 10) : "",
    [colunaOutraParte]: doc.dados.nomeOutraParte || "",
    "CNPJ": doc.dados.cnpjOutraParte || "",
    "Valor do serviço": Number(doc.dados.valorServico) || 0,
    "ISSQN": Number(doc.dados.issqn) || 0,
    "ISS retido": Number(doc.dados.issRetido) || 0,
    "PIS": Number(doc.dados.pis) || 0,
    "COFINS": Number(doc.dados.cofins) || 0,
    "CSLL": Number(doc.dados.csll) || 0,
    "IRRF": Number(doc.dados.ir) || 0,
  };
}

function gerarPlanilhaRetencoes(documentos, tipo) {
  const colunaOutraParte = tipo === "emitidas" ? "Tomador" : "Prestador";
  const normais = documentos.filter((doc) => !doc.dados.cancelada);
  const canceladas = documentos.filter((doc) => doc.dados.cancelada);

  const livro = XLSX.utils.book_new();
  const nomeAba = tipo === "emitidas" ? "Notas emitidas" : "Notas tomadas";
  XLSX.utils.book_append_sheet(
    livro,
    XLSX.utils.json_to_sheet(normais.map((doc) => linhaPlanilha(doc, colunaOutraParte))),
    nomeAba
  );
  if (canceladas.length > 0) {
    XLSX.utils.book_append_sheet(
      livro,
      XLSX.utils.json_to_sheet(canceladas.map((doc) => linhaPlanilha(doc, colunaOutraParte))),
      "Canceladas"
    );
  }
  return XLSX.write(livro, { type: "buffer", bookType: "xlsx" });
}

// Função principal chamada pela rota da API. Varre o NSU do cliente,
// filtra pelo período e pelo tipo (emitidas/tomadas) pedidos, monta o zip
// no formato pedido (xml ou pdf) e sempre monta a planilha de retenções
// junto — assim a pessoa não precisa anexar o certificado de novo só pra
// pegar a planilha depois.
export async function buscarNotas({
  pfxBuffer,
  senha,
  cnpjCliente,
  dataInicial,
  dataFinal,
  tipo,
  formato,
  // Opcional: chamado a cada página do NSU e a cada PDF buscado, pra quem
  // chamou poder mostrar uma barra de progresso. Nunca lança erro por causa
  // disso — se o callback falhar, ignoramos e seguimos a busca.
  onProgresso,
}) {
  const { chavePem, certPem } = lerCertificado(pfxBuffer, senha);

  const cnpjLimpo = (cnpjCliente || "").replace(/\D/g, "");
  // Decidido com a contabilidade em 2026-09-12: o período filtra pela
  // COMPETÊNCIA (dCompet), não pela data de emissão — uma nota de
  // substituição pode ser emitida num mês e valer pra competência de outro
  // (ex: nota do dia 1 já pode sair com competência do mês anterior, antes
  // de fechar a apuração), e é a competência que importa pra contabilidade.
  // dCompet vem sempre como "AAAA-MM-DD" (sem hora), que o JS já entende
  // como meia-noite UTC — por isso as duas pontas aqui também são "Z", sem
  // depender do fuso do servidor.
  const inicio = new Date(`${dataInicial}T00:00:00Z`);
  const fim = new Date(`${dataFinal}T00:00:00Z`);
  const incluirPdf = formato === "pdf";

  const documentos = [];
  // Chaves de acesso que apareceram com um evento de cancelamento
  // confirmado (ver EVENTOS_DE_CANCELAMENTO) — cobre o caso de a nota em
  // si ainda estar com cStat=100 mas ter sido cancelada por um evento
  // separado depois.
  const chavesCanceladasPorEvento = new Set();
  // Evita processar a mesma nota duas vezes — proteção pro caso de o
  // cursor de NSU pedir de novo um trecho já visto (ver comentário mais
  // abaixo sobre inclusivo x exclusivo).
  const chavesVistas = new Set();
  let nsuAtual = "0";
  // Antes era 5000 e, se estourado, a busca simplesmente parava e devolvia
  // o que tinha achado até ali como se fosse o resultado completo — sem
  // avisar nada. Pra CNPJ com histórico grande isso dava número errado em
  // silêncio (bug real, achado em 2026-09-16). Agora, além de bem mais
  // folgado, se ainda assim for atingido a busca falha com erro explícito
  // em vez de entregar um resultado que pode estar incompleto (ver
  // `varreduraCompleta` abaixo).
  const LIMITE_DE_SEGURANCA = 20000;

  function notificar(evento) {
    try {
      onProgresso?.(evento);
    } catch {
      // Falha ao notificar progresso não pode derrubar a busca em si.
    }
  }

  let varreduraCompleta = false;
  for (let i = 0; i < LIMITE_DE_SEGURANCA; i++) {
    notificar({ etapa: "buscando", pagina: i + 1, notasEncontradas: documentos.length });

    const { body } = await requisitarComCertificado({
      chavePem,
      certPem,
      path: PATHS.dfe(nsuAtual, cnpjLimpo),
    });

    let resposta;
    try {
      resposta = JSON.parse(body.toString("utf-8"));
    } catch {
      break;
    }

    if (resposta.StatusProcessamento === "REJEICAO") {
      const mensagens = (resposta.Erros || [])
        .map((e) => MENSAGENS_ERRO_GOVERNO[e.Codigo] || e.Descricao || e.Codigo)
        .join("; ");
      throw new Error(mensagens || "Governo rejeitou a consulta, sem detalhes.");
    }

    const lote = resposta.LoteDFe || [];
    if (lote.length === 0) {
      varreduraCompleta = true;
      break;
    }

    let maiorNsu = null;
    for (const item of lote) {
      if (item.NSU && (maiorNsu === null || Number(item.NSU) > Number(maiorNsu))) {
        maiorNsu = item.NSU;
      }

      if (item.TipoDocumento === "EVENTO" || item.TipoDocumento === "CNC") {
        if (item.ChaveAcesso && EVENTOS_DE_CANCELAMENTO.has(item.TipoEvento)) {
          chavesCanceladasPorEvento.add(item.ChaveAcesso);
        }
        continue;
      }
      if (item.TipoDocumento !== "NFSE") continue; // ignora DPS, etc.
      const chaveAcesso = item.ChaveAcesso;
      if (!chaveAcesso || !item.ArquivoXml) continue;
      if (chavesVistas.has(chaveAcesso)) continue; // já processada numa página anterior
      chavesVistas.add(chaveAcesso);

      let xml;
      try {
        xml = descompactarXml(item.ArquivoXml);
      } catch {
        xml = Buffer.from(item.ArquivoXml, "base64").toString("utf-8");
      }

      const dados = extrairDadosNota(xml, cnpjCliente);
      if (!dados) continue; // não deu pra classificar emitida/tomada com segurança
      if (dados.papel !== tipo) continue;
      // Filtra por competência, não por emissão (ver comentário acima, no
      // início de buscarNotas) — nota sem dCompet no XML (não deveria
      // acontecer, mas por segurança) não é descartada por falta de dado.
      if (dados.dataCompetencia && (dados.dataCompetencia < inicio || dados.dataCompetencia > fim)) continue;

      documentos.push({ chaveAcesso, xml, dados });
    }

    notificar({ etapa: "buscando", pagina: i + 1, notasEncontradas: documentos.length });

    // sem NSU novo no lote: não tem como avançar o cursor, para pra não girar em loop.
    if (maiorNsu === null) {
      varreduraCompleta = true;
      break;
    }

    // Confirmado com casos reais em 2026-09-16 (nota faltando sempre na
    // fronteira entre duas páginas, em CNPJs de volumes bem diferentes):
    // pedir "maiorNsu + 1" pulava exatamente 1 documento. Sinal de que a
    // API trata o NSU pedido como "o último que você já tem" e devolve só
    // o que vem DEPOIS dele — mesma convenção da Distribuição DFe de
    // NFe/CTe — em vez de "a partir desse NSU" como o código assumia.
    // Por isso pedimos o próprio maiorNsu nas próximas chamadas; a
    // deduplicação por chave de acesso (chavesVistas, acima) cobre o caso
    // de a API devolver de novo algo que já vimos.
    const proximoNsu = String(maiorNsu);
    if (proximoNsu === nsuAtual) {
      // não avançou nada — evita girar pra sempre se a suposição acima
      // estiver errada pra essa resposta específica.
      varreduraCompleta = true;
      break;
    }
    nsuAtual = proximoNsu;

    if (resposta.StatusProcessamento !== "DOCUMENTOS_LOCALIZADOS") {
      varreduraCompleta = true;
      break;
    }
  }

  if (!varreduraCompleta) {
    throw new Error(
      `A busca foi interrompida por segurança depois de ${LIMITE_DE_SEGURANCA} páginas de NSU, ` +
        `sem terminar de varrer o histórico desse CNPJ (provavelmente ele tem volume muito grande). ` +
        `${documentos.length} nota(s) foram encontradas até parar, mas o resultado NÃO está completo ` +
        `e por isso não foi entregue. Tente buscar um período menor (ex: mês a mês) pra esse CNPJ.`
    );
  }

  // Reconcilia com os eventos de cancelamento vistos em qualquer ponto da
  // varredura (podem ter aparecido antes ou depois da própria nota, já que
  // NSU é ordem de disponibilização, não ordem de emissão/cancelamento).
  for (const doc of documentos) {
    if (chavesCanceladasPorEvento.has(doc.chaveAcesso)) doc.dados.cancelada = true;
  }

  notificar({ etapa: "montando_arquivos", notasEncontradas: documentos.length });

  const zip = new JSZip();
  for (const [indice, doc] of documentos.entries()) {
    // Notas canceladas ficam isoladas numa pasta própria dentro do zip —
    // as normais continuam soltas na raiz, pra não misturar sem querer.
    const pasta = doc.dados.cancelada ? "canceladas/" : "";
    if (!incluirPdf) {
      zip.file(`${pasta}${doc.chaveAcesso}.xml`, doc.xml);
      continue;
    }
    notificar({
      etapa: "baixando_pdf",
      notasEncontradas: documentos.length,
      notaAtual: indice + 1,
    });
    try {
      const pdfBuffer = await gerarDanfsePdf(doc.xml, { cancelada: doc.dados.cancelada });
      zip.file(`${pasta}${doc.chaveAcesso}.pdf`, pdfBuffer);
    } catch (err) {
      zip.file(`${pasta}${doc.chaveAcesso}-ERRO.txt`, `Falha ao gerar o PDF: ${err.message}`);
    }
  }

  const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });
  const planilhaBuffer =
    documentos.length > 0 ? gerarPlanilhaRetencoes(documentos, tipo) : null;

  return { zipBuffer, planilhaBuffer, totalNotas: documentos.length };
}
