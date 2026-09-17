"use client";

import { useState, useRef } from "react";

// Converte um base64 (vindo da API) num download de verdade no navegador,
// sem precisar de outra requisição — os dados já estão na memória do
// navegador, então "Baixar planilha" depois de "Baixar XML/PDF" é
// instantâneo e não pede o certificado de novo.
function baixarBase64(base64, nomeArquivo, mime) {
  const binario = atob(base64);
  const bytes = new Uint8Array(binario.length);
  for (let i = 0; i < binario.length; i++) bytes[i] = binario.charCodeAt(i);
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nomeArquivo;
  a.click();
  URL.revokeObjectURL(url);
}

const MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

// Formata enquanto digita: mantém só os dígitos (até 14) e vai encaixando
// pontuação no formato xx.xxx.xxx/xxxx-xx conforme a pessoa digita.
function formatarCNPJ(valor) {
  const d = valor.replace(/\D/g, "").slice(0, 14);
  let out = d.slice(0, 2);
  if (d.length > 2) out += "." + d.slice(2, 5);
  if (d.length > 5) out += "." + d.slice(5, 8);
  if (d.length > 8) out += "/" + d.slice(8, 12);
  if (d.length > 12) out += "-" + d.slice(12, 14);
  return out;
}

// Mesma ideia do CNPJ, pra data: só dígitos (até 8), encaixando as barras
// no formato dd/mm/aaaa conforme digita.
function formatarData(valor) {
  const d = valor.replace(/\D/g, "").slice(0, 8);
  let out = d.slice(0, 2);
  if (d.length > 2) out += "/" + d.slice(2, 4);
  if (d.length > 4) out += "/" + d.slice(4, 8);
  return out;
}

// Converte "dd/mm/aaaa" (o que a pessoa digita) pro formato que a API
// espera ("aaaa-mm-dd"). Devolve null se a data estiver incompleta ou não
// existir de verdade no calendário (ex: 31/02) — quem chama decide o que
// fazer com null (mostrar erro em vez de mandar uma data inválida).
function dataParaISO(valorBR) {
  const d = valorBR.replace(/\D/g, "");
  if (d.length !== 8) return null;
  const dia = Number(d.slice(0, 2));
  const mes = Number(d.slice(2, 4));
  const ano = Number(d.slice(4, 8));
  const data = new Date(ano, mes - 1, dia);
  const valida = data.getFullYear() === ano && data.getMonth() === mes - 1 && data.getDate() === dia;
  return valida ? `${d.slice(4, 8)}-${d.slice(2, 4)}-${d.slice(0, 2)}` : null;
}

// Descreve o evento de progresso (ver onProgresso em lib/nfse.js) numa
// frase curta pra mostrar embaixo da barra.
function textoProgresso(p) {
  if (!p) return "Buscando…";
  if (p.etapa === "baixando_pdf") {
    return `Baixando PDF ${p.notaAtual} de ${p.notasEncontradas}…`;
  }
  if (p.etapa === "montando_arquivos") {
    return `Montando arquivos (${p.notasEncontradas} nota(s) encontrada(s))…`;
  }
  return p.notasEncontradas > 0
    ? `Buscando… ${p.notasEncontradas} nota(s) encontrada(s) até agora (página ${p.pagina})`
    : `Buscando… nenhuma nota encontrada ainda (página ${p.pagina})`;
}

export default function BuscaNotas() {
  const [cnpj, setCnpj] = useState("");
  const [dataInicial, setDataInicial] = useState("");
  const [dataFinal, setDataFinal] = useState("");
  const [tipo, setTipo] = useState("emitidas");
  const [carregando, setCarregando] = useState(false);
  const [progresso, setProgresso] = useState(null); // { etapa, pagina, notasEncontradas, notaAtual }
  const [mensagem, setMensagem] = useState(null); // { texto, tipo }
  const [resultado, setResultado] = useState(null); // { planilhaBase64, nomeArquivoPlanilha, totalNotas }
  const formRef = useRef(null);

  // "Baixar PDF" gera o DANFSe localmente (lib/danfse.js) desde que a API
  // do governo pra isso foi desativada em 01/07/2026 — ver aviso em
  // lib/nfse.js. É uma primeira versão: cobre os blocos principais do
  // layout, mas ainda não tem canhoto nem tributação IBS/CBS.
  //
  // A resposta da API vem em NDJSON (uma linha JSON por evento), não um
  // JSON único — assim dá pra mostrar o progresso da busca (página do NSU,
  // quantas notas já achou) em vez de "Buscando…" parado. A última linha
  // é sempre o evento final: "resultado", "aviso" ou "erro".
  async function buscar(formato) {
    const formEl = formRef.current;
    if (!formEl) return;

    const dataInicialISO = dataParaISO(dataInicial);
    const dataFinalISO = dataParaISO(dataFinal);
    if (!dataInicialISO || !dataFinalISO) {
      setMensagem({ texto: "Preencha as duas datas do período corretamente (dd/mm/aaaa).", tipo: "erro" });
      return;
    }

    const formData = new FormData(formEl);
    formData.set("tipo", tipo);
    formData.set("formato", formato);
    formData.set("dataInicial", dataInicialISO);
    formData.set("dataFinal", dataFinalISO);

    setCarregando(true);
    setProgresso(null);
    setMensagem(null);
    setResultado(null);

    let final = null;
    try {
      const res = await fetch("/api/buscar", { method: "POST", body: formData });
      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let restante = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        restante += decoder.decode(value, { stream: true });

        let fimDaLinha;
        while ((fimDaLinha = restante.indexOf("\n")) >= 0) {
          const linha = restante.slice(0, fimDaLinha).trim();
          restante = restante.slice(fimDaLinha + 1);
          if (!linha) continue;

          const evento = JSON.parse(linha);
          if (evento.tipo === "progresso") {
            setProgresso(evento);
          } else {
            final = evento; // aviso | erro | resultado
          }
        }
      }
    } catch {
      setCarregando(false);
      setMensagem({ texto: "Falha de conexão. Tente de novo.", tipo: "erro" });
      return;
    }

    setCarregando(false);
    setProgresso(null);

    if (!final || final.tipo === "erro") {
      setMensagem({ texto: final?.erro || "Algo deu errado.", tipo: "erro" });
      return;
    }
    if (final.tipo === "aviso") {
      setMensagem({ texto: final.aviso, tipo: "aviso" });
      return;
    }

    // Baixa automaticamente o arquivo pedido (XML ou PDF).
    baixarBase64(final.zipBase64, final.nomeArquivoZip, "application/zip");

    setMensagem({
      texto: `${final.totalNotas} nota(s) baixada(s) em ${formato.toUpperCase()}.`,
      tipo: "ok",
    });
    setResultado({
      planilhaBase64: final.planilhaBase64,
      nomeArquivoPlanilha: final.nomeArquivoPlanilha,
      totalNotas: final.totalNotas,
    });
  }

  function baixarPlanilha() {
    if (!resultado) return;
    baixarBase64(resultado.planilhaBase64, resultado.nomeArquivoPlanilha, MIME_XLSX);
  }

  return (
    <div>
      <form
        ref={formRef}
        onSubmit={(e) => e.preventDefault()}
        style={{ display: "flex", flexDirection: "column", maxWidth: 460 }}
      >
        <div className="campo">
          <label htmlFor="cnpj">CNPJ do cliente</label>
          <input
            id="cnpj"
            name="cnpj"
            type="text"
            inputMode="numeric"
            placeholder="00.000.000/0000-00"
            value={cnpj}
            onChange={(e) => setCnpj(formatarCNPJ(e.target.value))}
            required
            disabled={carregando}
          />
        </div>

        <div className="grupo">
          <div className="grupo-titulo">Período</div>
          <div className="grupo-linha">
            <div className="campo">
              <label htmlFor="dataInicial">De</label>
              <input
                id="dataInicial"
                type="text"
                inputMode="numeric"
                placeholder="dd/mm/aaaa"
                value={dataInicial}
                onChange={(e) => setDataInicial(formatarData(e.target.value))}
                required
                disabled={carregando}
              />
            </div>
            <div className="campo">
              <label htmlFor="dataFinal">Até</label>
              <input
                id="dataFinal"
                type="text"
                inputMode="numeric"
                placeholder="dd/mm/aaaa"
                value={dataFinal}
                onChange={(e) => setDataFinal(formatarData(e.target.value))}
                required
                disabled={carregando}
              />
            </div>
          </div>
        </div>

        <div className="campo">
          <div className="segmentado" role="group" aria-label="Tipo de nota">
            <button
              type="button"
              className={tipo === "emitidas" ? "ativo" : ""}
              aria-pressed={tipo === "emitidas"}
              disabled={carregando}
              onClick={() => setTipo("emitidas")}
            >
              Notas emitidas
            </button>
            <button
              type="button"
              className={tipo === "tomadas" ? "ativo" : ""}
              aria-pressed={tipo === "tomadas"}
              disabled={carregando}
              onClick={() => setTipo("tomadas")}
            >
              Notas tomadas
            </button>
          </div>
        </div>

        <div className="grupo">
          <div className="grupo-titulo">Certificado</div>
          <div className="grupo-linha">
            <div className="campo">
              <label htmlFor="certificado">Arquivo (.pfx)</label>
              <input id="certificado" name="certificado" type="file" accept=".pfx,.p12" required disabled={carregando} />
            </div>
            <div className="campo">
              <label htmlFor="senhaCertificado">Senha</label>
              <input id="senhaCertificado" name="senhaCertificado" type="password" required disabled={carregando} />
            </div>
          </div>
        </div>

        <div className="segmentado segmentado-largo segmentado-acoes">
          <button type="button" disabled={carregando} onClick={() => buscar("xml")}>
            {carregando ? "Buscando…" : "Baixar XML"}
          </button>
          <button type="button" disabled={carregando} onClick={() => buscar("pdf")}>
            {carregando ? "Buscando…" : "Baixar PDF"}
          </button>
        </div>
        <div style={{ marginTop: 6, fontSize: 12, color: "var(--tinta-suave)" }}>
          O PDF é gerado pelo sistema a partir do XML oficial (a API do governo pra isso foi
          desativada) — a autenticidade da nota se confere pela chave de acesso ou QR Code no
          Portal Nacional da NFS-e.
        </div>
      </form>

      {carregando && (
        <div style={{ marginTop: 16 }}>
          <div className="barra-progresso">
            <div className="barra-progresso-preenchimento" />
          </div>
          <div style={{ marginTop: 8, fontSize: 13, color: "var(--tinta-suave)" }}>
            {textoProgresso(progresso)}
          </div>
        </div>
      )}

      {mensagem && <div className={`mensagem-${mensagem.tipo}`} style={{ marginTop: 16 }}>{mensagem.texto}</div>}

      {resultado && (
        <button className="botao botao-secundario" style={{ marginTop: 8 }} onClick={baixarPlanilha}>
          Baixar planilha de retenções ({resultado.totalNotas} notas)
        </button>
      )}
    </div>
  );
}
