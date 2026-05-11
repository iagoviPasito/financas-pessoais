/* eslint-disable */
import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import Papa from "papaparse";

// ─── Constantes ───────────────────────────────────────────────────────────────
const CATEGORIAS_GASTO = ["Moradia","Alimentação","Transporte","Saúde","Lazer","Educação","Vestuário","Investimento","Moto/Veículo","Outros"];
const CATEGORIAS_RENDA = ["Salário CLT","Freela/PJ","Investimentos/Dividendos","Outros"];
const CORES       = ["#c8a96e","#e8c88a","#a07840","#d4b070","#7a5c30","#f0d898","#604818","#dca850","#8a6428","#b89050"];
const CORES_RENDA = ["#4a9e6a","#2d7a4e","#6ab882","#1a5c38"];
const MESES       = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
const gerarId     = () => Math.random().toString(36).substr(2, 9);
const mesAtual    = new Date().getMonth();
const anoAtual    = new Date().getFullYear();
const fmt         = (v) => Number(v).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

// ─── Categorização automática (regras locais) ─────────────────────────────────
const REGRAS = [
  { palavras: ["aluguel","condominio","iptu","agua","luz","energia","gas","internet","telefone","celular"], cat: "Moradia", tipo: "gasto" },
  { palavras: ["mercado","supermercado","ifood","rappi","delivery","restaurante","lanchonete","mc donald","burger","subway","pizza","sushi","padaria","acougue","alimentacao","refeicao","carrefour","extra","atacadao"], cat: "Alimentação", tipo: "gasto" },
  { palavras: ["uber","99pop","taxi","onibus","metro","combustivel","gasolina","etanol","posto","estacionamento","pedagio","autopecas","veiculo"], cat: "Transporte", tipo: "gasto" },
  { palavras: ["farmacia","drogaria","medico","hospital","clinica","dentista","plano de saude","exame","laboratorio","remedios","consulta"], cat: "Saúde", tipo: "gasto" },
  { palavras: ["netflix","spotify","amazon","disney","cinema","teatro","show","bar","academia","jogo","steam","xbox","playstation","lazer","entretenimento"], cat: "Lazer", tipo: "gasto" },
  { palavras: ["escola","faculdade","curso","livro","udemy","alura","mensalidade","educacao"], cat: "Educação", tipo: "gasto" },
  { palavras: ["renner","cea","zara","hm","riachuelo","americanas","shein","shopee","roupa","calcado","tenis","sapato"], cat: "Vestuário", tipo: "gasto" },
  { palavras: ["nuinvest","xp","btg","rico","clear","acao","tesouro","fundo","cdb","lci","lca","investimento","aplicacao"], cat: "Investimento", tipo: "gasto" },
  { palavras: ["royal enfield","moto","revisao moto","seguro moto","emplacamento","capacete"], cat: "Moto/Veículo", tipo: "gasto" },
  { palavras: ["salario","pagamento empregador","holerite","folha"], cat: "Salário CLT", tipo: "renda" },
  { palavras: ["freela","freelance","nota fiscal","honorario","consultoria","servico prestado"], cat: "Freela/PJ", tipo: "renda" },
  { palavras: ["dividendo","rendimento","juros","resgate","renda fixa","renda variavel"], cat: "Investimentos/Dividendos", tipo: "renda" },
];
function normalizar(str) {
  return String(str).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9 ]/g," ");
}
function categorizarAuto(descricao, valor) {
  const d = normalizar(descricao);
  for (const r of REGRAS) {
    if (r.palavras.some(p => d.includes(normalizar(p)))) return { categoria: r.cat, tipo: r.tipo };
  }
  return valor >= 0 ? { categoria: "Outros", tipo: "renda" } : { categoria: "Outros", tipo: "gasto" };
}

// ─── Parser de voz via Claude AI ──────────────────────────────────────────────
async function interpretarFala(transcricao) {
  const hoje = new Date();
  const prompt = `Você é um assistente financeiro. O usuário disse em voz alta: "${transcricao}"

Extraia a transação financeira e retorne APENAS um JSON válido, sem nenhum texto adicional, no formato:
{
  "descricao": "string curta descrevendo o gasto/renda",
  "valor": number (positivo, sem símbolo de moeda),
  "tipo": "gasto" | "renda",
  "categoria": uma das seguintes opções exatas: "Moradia","Alimentação","Transporte","Saúde","Lazer","Educação","Vestuário","Investimento","Moto/Veículo","Salário CLT","Freela/PJ","Investimentos/Dividendos","Outros",
  "mes": number entre 0 e 11 (0=Janeiro),
  "ano": number (ano com 4 dígitos)
}

Regras:
- Se o usuário não mencionar data, use mês ${hoje.getMonth()} e ano ${hoje.getFullYear()}.
- Se disser "ontem", "semana passada", "mês passado", calcule a partir de hoje (${hoje.toLocaleDateString("pt-BR")}).
- Palavras como "gastei", "paguei", "comprei" → tipo "gasto".
- Palavras como "recebi", "ganhei", "entrou" → tipo "renda".
- Valores por extenso como "cem reais", "duzentos e cinquenta" → converta para número.
- Gasolina, Uber, combustível → Transporte. Mercado, restaurante → Alimentação. Farmácia → Saúde. Etc.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 300,
      messages: [{ role: "user", content: prompt }]
    })
  });

  const data = await response.json();
  const texto = data.content?.find(b => b.type === "text")?.text || "";
  // Extrai JSON mesmo se vier com texto ao redor
  const match = texto.match(/\{[\s\S]*\}/);
  if (!match) throw new Error("Resposta inválida da IA");
  return JSON.parse(match[0]);
}

// ─── Parsers de extrato ───────────────────────────────────────────────────────
function parsearValorBR(str) {
  return parseFloat(String(str||"").trim().replace(/\./g,"").replace(",",".")) || 0;
}
function parsearDataBR(str) {
  const p = String(str||"").trim().split("/");
  if (p.length === 3) return { mes: parseInt(p[1])-1, ano: parseInt(p[2]) };
  return { mes: mesAtual, ano: anoAtual };
}
function parsearOFX(texto) {
  const transacoes = [];
  const blocos = texto.match(/<STMTTRN>([\s\S]*?)<\/STMTTRN>/gi) || [];
  for (const bloco of blocos) {
    const get = (tag) => { const m = bloco.match(new RegExp(`<${tag}>([^<\n\r]+)`,"i")); return m ? m[1].trim() : ""; };
    const valor = parseFloat(get("TRNAMT").replace(",","."));
    const memo  = get("MEMO") || get("NAME") || "Sem descrição";
    const dtStr = get("DTPOSTED");
    const ano   = dtStr.length >= 4 ? parseInt(dtStr.substr(0,4)) : anoAtual;
    const mes   = dtStr.length >= 6 ? parseInt(dtStr.substr(4,2))-1 : mesAtual;
    if (!isNaN(valor)) transacoes.push({ memo, valor, mes, ano });
  }
  return transacoes;
}
function parsearCSV(conteudo) {
  const sep    = conteudo.includes(";") ? ";" : ",";
  const result = Papa.parse(conteudo, { delimiter: sep });
  const rows   = result.data;
  if (rows.length < 2) return null;

  const isBradescoHeader = (row) => {
    const n = row.map(c => normalizar(c));
    return n.some(c => c.includes("historico")) && n.some(c => c.includes("credito")) && n.some(c => c.includes("debito"));
  };

  const headerLines = rows.map((r,i) => ({i,r})).filter(({r}) => isBradescoHeader(r));
  if (headerLines.length > 0) {
    const transacoes = [];
    for (const {i: hi, r: hrow} of headerLines) {
      const n    = hrow.map(c => normalizar(c));
      const iDt  = n.findIndex(c => c.includes("data"));
      const iHi  = n.findIndex(c => c.includes("historico"));
      const iCr  = n.findIndex(c => c.includes("credito"));
      const iDb  = n.findIndex(c => c.includes("debito"));
      for (let j = hi+1; j < rows.length; j++) {
        const r  = rows[j];
        const dc = String(r[iDt]||"").trim();
        if (!dc || isBradescoHeader(r)) break;
        if (!/^\d{2}\/\d{2}\/\d{4}$/.test(dc)) continue;
        const memo = String(r[iHi]||"Sem descrição").trim();
        const cred = parsearValorBR(r[iCr]);
        const deb  = parsearValorBR(r[iDb]);
        const {mes, ano} = parsearDataBR(dc);
        if (cred > 0) transacoes.push({ memo, valor: cred,  mes, ano });
        if (deb  > 0) transacoes.push({ memo, valor: -deb,  mes, ano });
      }
    }
    if (transacoes.length > 0) return transacoes;
  }

  const firstHeader = rows.find(r => r.some(c => String(c).trim().length > 0));
  if (!firstHeader) return null;
  const headers  = firstHeader.map(h => normalizar(h));
  const dataRows = rows.slice(rows.indexOf(firstHeader)+1);

  if (headers.some(h => h.includes("categoria")) && headers.some(h => h.includes("titulo"))) {
    const iDt = headers.findIndex(h => h.includes("data"));
    const iTi = headers.findIndex(h => h.includes("titulo"));
    const iVl = headers.findIndex(h => h.includes("valor"));
    return dataRows.map(r => {
      const valor = parsearValorBR(r[iVl]) * -1;
      const {mes,ano} = parsearDataBR(r[iDt]);
      return { memo: String(r[iTi]||"Sem descrição").trim(), valor, mes, ano };
    }).filter(r => r.memo && r.valor !== 0);
  }

  const iDesc = headers.findIndex(h => h.includes("descri")||h.includes("memo")||h.includes("hist")||h.includes("lancamento")||h.includes("titulo")||h.includes("estabelec"));
  const iVal  = headers.findIndex(h => h.includes("valor")||h.includes("amount"));
  const iData = headers.findIndex(h => h.includes("data")||h.includes("date"));
  if (iDesc < 0 || iVal < 0) return null;

  return dataRows.map(r => {
    const memo  = String(r[iDesc]||"Sem descrição").trim();
    const valor = parsearValorBR(r[iVal]);
    const {mes,ano} = iData >= 0 ? parsearDataBR(r[iData]) : {mes:mesAtual,ano:anoAtual};
    return { memo, valor, mes, ano };
  }).filter(r => r.memo && r.valor !== 0);
}


// ─── Parsers PDF multi-banco ─────────────────────────────────────────────────
async function carregarPdfJS() {
  if (window.pdfjsLib) return window.pdfjsLib;
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      resolve(window.pdfjsLib);
    };
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

async function parsearPDF(arrayBuffer) {
  const pdfjsLib = await carregarPdfJS();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let paginas = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const itens = content.items
      .map(it => ({ texto: it.str.trim(), x: Math.round(it.transform[4]), y: Math.round(it.transform[5]) }))
      .filter(it => it.texto.length > 0);
    const linhaMap = {};
    itens.forEach(it => { if (!linhaMap[it.y]) linhaMap[it.y] = []; linhaMap[it.y].push(it); });
    const linhas = Object.values(linhaMap)
      .sort((a, b) => b[0].y - a[0].y)
      .map(its => its.sort((a, b) => a.x - b.x).map(it => it.texto).join(" "));
    paginas.push(linhas.join("\n"));
  }
  const textoTotal = paginas.join("\n");

  // Detectar banco/tipo pelo conteúdo
  const isItauCard    = /Lançamentos: compras e saques/i.test(textoTotal);
  const isBradCard    = /Lançamentos\s*Data\s*Histórico de Lançamentos|Fatura Mensal.*bradesco|ELO NANQUIM/i.test(textoTotal);
  const isBradCC      = /Extrato de: Agência.*Conta.*Movimentação|Bradesco Celular/i.test(textoTotal);
  const isItauCC      = /extrato conta.*lançamentos|REND PAGO APLIC AUT MAIS/i.test(textoTotal);

  if (isItauCard)  return extrairItauCartao(textoTotal);
  if (isBradCard)  return extrairBradescoCartao(textoTotal);
  if (isBradCC)    return extrairBradescoCC(textoTotal);
  if (isItauCC)    return extrairItauCC(textoTotal);
  // Fallback: tentar Itaú cartão
  return extrairItauCartao(textoTotal);
}

function pVal(s) { return parseFloat(String(s||"").replace(/\./g,"").replace(",",".")) || 0; }
function parseDDMM(dd_mm, anoRef) { const [dd,mm] = dd_mm.split("/"); return { mes: parseInt(mm)-1, ano: anoRef||anoAtual }; }
function parseFullDate(s) { const p = s.split("/"); return p.length===3 ? { mes: parseInt(p[1])-1, ano: parseInt(p[2]) } : { mes: mesAtual, ano: anoAtual }; }
function limparDesc(s) { return s.replace(/\s+\d{2}\/\d{2}$/, "").replace(/\s+(\w{2,}\.?\s*){1,3}$/, s => /[A-Z]{2,}/.test(s) ? "" : s).trim(); }

const RE_IGNORAR_GERAL = /SALDO DO DIA|REND PAGO APLIC|Limite da Conta|saldo em conta|Total para|Total da fatura|Compras parceladas|próximas faturas|Limites de crédito|Encargos cobrados|Simulação|Novo teto|Taxas mensais|PROGRAMA DE FIDELIDADE|Central de atendimento|Mensagem Importante|Folha:|emitido em|período de visualização/i;

// ── Itaú Cartão ──
function extrairItauCartao(texto) {
  const transacoes = [];
  const RE = /^(\d{2}\/\d{2})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s*$/;
  const RE_NEG = /^(\d{2}\/\d{2})\s+(.+?)\s+-\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s*$/;
  const RE_IGNORE = /MULTA|ENCARGO|IOF|ANUIDADE|JUROS|CET|Total|Limite|Valor total|Pagamento|Vencimento|Fechamento|Próxima|Demais|Parcelas|Simulação|Saldo ant/i;
  let anoFatura = anoAtual;
  const mAno = texto.match(/Vencimento[:\s]+(\d{2}\/\d{2}\/(\d{4}))/i);
  if (mAno) anoFatura = parseInt(mAno[2]);
  for (const linha of texto.split("\n")) {
    const l = linha.trim();
    if (!l || RE_IGNORAR_GERAL.test(l) || RE_IGNORE.test(l)) continue;
    if (!/^\d{2}\/\d{2}\s/.test(l)) continue;
    let m = RE_NEG.exec(l);
    if (m) { const [,data,desc,val] = m; const {mes,ano} = parseDDMM(data, anoFatura); transacoes.push({ memo: limparDesc(desc), valor: -pVal(val), mes, ano }); continue; }
    m = RE.exec(l);
    if (m) { const [,data,desc,val] = m; const {mes,ano} = parseDDMM(data, anoFatura); transacoes.push({ memo: limparDesc(desc), valor: pVal(val), mes, ano }); }
  }
  return transacoes.length > 0 ? transacoes : null;
}

// ── Bradesco Cartão ──
function extrairBradescoCartao(texto) {
  const transacoes = [];
  const RE = /^(\d{2}\/\d{2})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s*(-?)\s*$/;
  const RE_IGNORE = /MULTA|ENCARGO|IOF|ANUIDADE|JUROS|PAGTO ANTECIPADO|Total|Limite|Taxas|Programa|CET|Rotativo|Saldo ant/i;
  let anoFatura = anoAtual;
  const mAno = texto.match(/Vencimento[\s\n]+(\d{2}\/\d{2}\/(\d{4}))/i);
  if (mAno) anoFatura = parseInt(mAno[2]);
  for (const linha of texto.split("\n")) {
    const l = linha.trim();
    if (!l || RE_IGNORAR_GERAL.test(l) || RE_IGNORE.test(l)) continue;
    if (!/^\d{2}\/\d{2}\s/.test(l)) continue;
    const m = RE.exec(l);
    if (m) {
      const [,data,desc,val,neg] = m;
      const v = pVal(val);
      const {mes,ano} = parseDDMM(data, anoFatura);
      // neg = crédito/pagamento = renda; normal = gasto
      transacoes.push({ memo: limparDesc(desc), valor: neg ? -v : v, mes, ano });
    }
  }
  return transacoes.length > 0 ? transacoes : null;
}

// ── Bradesco Conta Corrente ──
function extrairBradescoCC(texto) {
  const transacoes = [];
  // Formato: DD/MM/AAAA HISTÓRICO [DESCRIÇÃO] DOCTO CREDITO|DEBITO SALDO
  const RE = /^(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s+[\d.,]+\s*$/;
  const RE_IGNORE = /COD\.\s*LANC|Total\b|^Folha|^Extrato|^Data\s+Histórico|^Nome:|^Bradesco/i;
  for (const linha of texto.split("\n")) {
    const l = linha.trim();
    if (!l || RE_IGNORAR_GERAL.test(l) || RE_IGNORE.test(l)) continue;
    const m = RE.exec(l);
    if (m) {
      const [,data,desc,val] = m;
      const v = pVal(val);
      const {mes,ano} = parseFullDate(data);
      // REM: = recebeu (crédito=renda), DES: = enviou (débito=gasto)
      const tipo = /^REM:|receb/i.test(desc) ? "renda" : "gasto";
      transacoes.push({ memo: desc.replace(/^(DES:|REM:)\s*/i,"").replace(/\s+\d{2}\/\d{2}\s+\d+$/, "").trim(), valor: tipo === "gasto" ? v : -v, mes, ano });
    }
  }
  return transacoes.length > 0 ? transacoes : null;
}

// ── Itaú Conta Corrente ──
function extrairItauCC(texto) {
  const transacoes = [];
  const RE = /^(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})\s*$/;
  const RE_IGNORE = /SALDO DO DIA|REND PAGO APLIC|Limite da Conta|saldo em conta|emitido em|período de|IAGO VIACELLI|agência:|^itaú|^Aviso|^Os saldos|^Consultas/i;
  for (const linha of texto.split("\n")) {
    const l = linha.trim();
    if (!l || RE_IGNORAR_GERAL.test(l) || RE_IGNORE.test(l)) continue;
    const m = RE.exec(l);
    if (m) {
      const [,data,desc,val] = m;
      const neg = val.startsWith("-");
      const v = parseFloat(val.replace("-","").replace(/\./g,"").replace(",","."));
      const {mes,ano} = parseFullDate(data);
      transacoes.push({ memo: desc.trim(), valor: neg ? v : -v, mes, ano });
    }
  }
  return transacoes.length > 0 ? transacoes : null;
}

// ── Extrator de parcelas futuras do Itaú Cartão ──────────────────────────────
function extrairParcelasFuturas(textoTotal) {
  const parcelas = [];
  const secao = textoTotal.match(/Compras parceladas.*?próximas faturas([\s\S]*?)(?:Limites de crédito|$)/i);
  if (!secao) return [];
  const RE = /^(\d{2}\/\d{2})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s*$/;
  let anoRef = anoAtual;
  for (const linha of secao[1].split("\n")) {
    const l = linha.trim();
    if (!l) continue;
    const m = RE.exec(l);
    if (m) {
      const [,data,desc,val] = m;
      const {mes,ano} = parseDDMM(data, anoRef);
      parcelas.push({ memo: limparDesc(desc), valor: pVal(val), mes, ano });
    }
  }
  return parcelas;
}

// ─── Dados iniciais ───────────────────────────────────────────────────────────
const dadosIniciais = {
  transacoes: [
    { id: gerarId(), tipo:"gasto",  descricao:"Aluguel", valor:1800, categoria:"Moradia",     mes:mesAtual, ano:anoAtual },
    { id: gerarId(), tipo:"gasto",  descricao:"Mercado", valor:650,  categoria:"Alimentação", mes:mesAtual, ano:anoAtual },
    { id: gerarId(), tipo:"gasto",  descricao:"Gasolina",valor:300,  categoria:"Transporte",  mes:mesAtual, ano:anoAtual },
    { id: gerarId(), tipo:"renda",  descricao:"Salário", valor:5500, categoria:"Salário CLT", mes:mesAtual, ano:anoAtual },
    { id: gerarId(), tipo:"renda",  descricao:"Freela",  valor:1200, categoria:"Freela/PJ",   mes:mesAtual, ano:anoAtual },
  ],
  metas: [
    { id: gerarId(), categoria:"Moradia",     limite:2000 },
    { id: gerarId(), categoria:"Alimentação", limite:800  },
    { id: gerarId(), categoria:"Lazer",       limite:500  },
  ]
};

// ─── Componente de Voz ────────────────────────────────────────────────────────
function VozInput({ onTransacaoDetectada }) {
  const [estado, setEstado] = useState("idle"); // idle | gravando | processando | confirmando | erro
  const [transcricao, setTranscricao] = useState("");
  const [txDetectada, setTxDetectada] = useState(null);
  const [erro, setErro] = useState("");
  const [pulso, setPulso] = useState(0);
  const recognitionRef = useRef(null);
  const pulsoRef = useRef(null);

  const suportado = typeof window !== "undefined" && ("SpeechRecognition" in window || "webkitSpeechRecognition" in window);

  useEffect(() => {
    if (estado === "gravando") {
      pulsoRef.current = setInterval(() => setPulso(p => (p+1)%3), 500);
    } else {
      clearInterval(pulsoRef.current);
      setPulso(0);
    }
    return () => clearInterval(pulsoRef.current);
  }, [estado]);

  const iniciarGravacao = useCallback(() => {
    if (!suportado) { setErro("Seu navegador não suporta reconhecimento de voz. Use Chrome ou Edge."); setEstado("erro"); return; }
    setErro(""); setTranscricao(""); setTxDetectada(null);
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    const rec = new SR();
    rec.lang = "pt-BR";
    rec.continuous = false;
    rec.interimResults = true;
    rec.maxAlternatives = 1;
    rec.onresult = (e) => {
      const t = Array.from(e.results).map(r => r[0].transcript).join("");
      setTranscricao(t);
    };
    rec.onend = async () => {
      const texto = recognitionRef.current?._finalText;
      if (!texto || texto.trim().length < 3) { setEstado("idle"); return; }
      setEstado("processando");
      try {
        const tx = await interpretarFala(texto);
        if (!tx.valor || tx.valor <= 0) throw new Error("Valor não identificado");
        setTxDetectada({ ...tx, id: gerarId() });
        setEstado("confirmando");
      } catch(e) {
        setErro("Não consegui entender. Tente novamente com mais clareza.");
        setEstado("erro");
      }
    };
    rec.onerror = (e) => {
      if (e.error === "no-speech") { setEstado("idle"); return; }
      setErro("Erro no microfone: " + e.error);
      setEstado("erro");
    };
    // Guarda transcrição final no ref para usar no onend
    rec.onresult = (e) => {
      const isFinal = e.results[e.results.length-1].isFinal;
      const t = Array.from(e.results).map(r => r[0].transcript).join("");
      setTranscricao(t);
      if (isFinal) recognitionRef.current._finalText = t;
    };
    recognitionRef.current = rec;
    recognitionRef.current._finalText = "";
    rec.start();
    setEstado("gravando");
  }, [suportado]);

  const pararGravacao = useCallback(() => {
    recognitionRef.current?.stop();
  }, []);

  const confirmar = () => {
    onTransacaoDetectada(txDetectada);
    setEstado("idle");
    setTranscricao("");
    setTxDetectada(null);
  };

  const cancelar = () => { setEstado("idle"); setTranscricao(""); setTxDetectada(null); };

  const corEstado = { idle:"#c8a96e", gravando:"#e05040", processando:"#7090e0", confirmando:"#4a9e6a", erro:"#e05040" };

  return (
    <div style={{ borderTop:"1px solid #2a2418", marginTop:22, paddingTop:18 }}>
      <div style={{ fontSize:11, color:"#c8a96e", letterSpacing:2, marginBottom:12 }}>🎙 LANÇAMENTO POR VOZ</div>

      {/* Botão principal */}
      {(estado === "idle" || estado === "erro") && (
        <div>
          <button
            onClick={iniciarGravacao}
            style={{ ...btn, width:"100%", background:"#1a0e20", color:"#c090e0", border:"1px solid #6040a0",
              display:"flex", alignItems:"center", justifyContent:"center", gap:8, fontSize:13 }}
          >
            <span style={{ fontSize:18 }}>🎙</span> Falar transação
          </button>
          <div style={{ fontSize:11, color:"#504030", marginTop:8, lineHeight:1.6 }}>
            Ex: <i style={{color:"#706050"}}>"Gastei 100 reais de gasolina"</i><br/>
            ou: <i style={{color:"#706050"}}>"Recebi 1500 de freela ontem"</i>
          </div>
          {estado === "erro" && erro && <div style={{ fontSize:12, color:"#e08070", marginTop:8 }}>⚠ {erro}</div>}
        </div>
      )}

      {/* Gravando */}
      {estado === "gravando" && (
        <div style={{ textAlign:"center" }}>
          <button onClick={pararGravacao} style={{ width:72, height:72, borderRadius:"50%", border:"3px solid #e05040", background:"#2a0a0a", cursor:"pointer", fontSize:28, display:"inline-flex", alignItems:"center", justifyContent:"center", boxShadow:`0 0 ${12 + pulso*8}px #e0504060` }}>
            ⏹
          </button>
          <div style={{ marginTop:12, fontSize:12, color:"#e05040", letterSpacing:2 }}>
            {"● GRAVANDO" + ".".repeat(pulso+1)}
          </div>
          {transcricao && (
            <div style={{ marginTop:10, fontSize:13, color:"#a09070", background:"#1a1510", borderRadius:8, padding:"8px 12px", fontStyle:"italic" }}>
              "{transcricao}"
            </div>
          )}
          <div style={{ fontSize:11, color:"#504030", marginTop:8 }}>Clique em ⏹ ou pare de falar para processar</div>
        </div>
      )}

      {/* Processando */}
      {estado === "processando" && (
        <div style={{ textAlign:"center", padding:"16px 0" }}>
          <div style={{ fontSize:24, marginBottom:10 }}>⚙️</div>
          <div style={{ fontSize:13, color:"#7090e0", letterSpacing:1 }}>Interpretando com IA...</div>
          {transcricao && (
            <div style={{ marginTop:10, fontSize:12, color:"#706050", fontStyle:"italic" }}>"{transcricao}"</div>
          )}
        </div>
      )}

      {/* Confirmação */}
      {estado === "confirmando" && txDetectada && (
        <div style={{ background:"#0e1a10", border:"1px solid #2a4a2a", borderRadius:10, padding:"14px 16px" }}>
          <div style={{ fontSize:11, color:"#4a9e6a", letterSpacing:2, marginBottom:12 }}>✓ TRANSAÇÃO DETECTADA</div>

          {transcricao && (
            <div style={{ fontSize:11, color:"#605040", fontStyle:"italic", marginBottom:10, padding:"6px 8px", background:"#1a1510", borderRadius:6 }}>
              🎙 "{transcricao}"
            </div>
          )}

          <div style={{ display:"grid", gap:6 }}>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:13 }}>
              <span style={{ color:"#706050" }}>Descrição</span>
              <span style={{ color:"#e8dcc8" }}>{txDetectada.descricao}</span>
            </div>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:13 }}>
              <span style={{ color:"#706050" }}>Valor</span>
              <span style={{ fontWeight:"bold", color: txDetectada.tipo==="renda" ? "#4a9e6a" : "#e05040" }}>
                {txDetectada.tipo==="renda" ? "+" : "-"}{fmt(txDetectada.valor)}
              </span>
            </div>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:13 }}>
              <span style={{ color:"#706050" }}>Tipo</span>
              <span style={{ color:"#e8dcc8" }}>{txDetectada.tipo === "gasto" ? "💸 Gasto" : "💰 Renda"}</span>
            </div>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:13 }}>
              <span style={{ color:"#706050" }}>Categoria</span>
              <span style={{ color:"#e8dcc8" }}>{txDetectada.categoria}</span>
            </div>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:13 }}>
              <span style={{ color:"#706050" }}>Data</span>
              <span style={{ color:"#e8dcc8" }}>{MESES[txDetectada.mes]}/{txDetectada.ano}</span>
            </div>
          </div>

          {/* Edição rápida */}
          <div style={{ borderTop:"1px solid #2a3a2a", marginTop:12, paddingTop:12, display:"grid", gap:6 }}>
            <div style={{ fontSize:10, color:"#4a7a4a", letterSpacing:1, marginBottom:4 }}>AJUSTAR SE NECESSÁRIO</div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 }}>
              <div>
                <div style={{ fontSize:10, color:"#506050", marginBottom:3 }}>TIPO</div>
                <select value={txDetectada.tipo} onChange={e => setTxDetectada(t => ({...t, tipo:e.target.value}))} style={{...smInput, color: txDetectada.tipo==="renda"?"#4a9e6a":"#e05040"}}>
                  <option value="gasto">💸 Gasto</option>
                  <option value="renda">💰 Renda</option>
                </select>
              </div>
              <div>
                <div style={{ fontSize:10, color:"#506050", marginBottom:3 }}>VALOR</div>
                <input type="number" value={txDetectada.valor} onChange={e => setTxDetectada(t => ({...t, valor:parseFloat(e.target.value)||0}))} style={smInput}/>
              </div>
            </div>
            <div>
              <div style={{ fontSize:10, color:"#506050", marginBottom:3 }}>CATEGORIA</div>
              <select value={txDetectada.categoria} onChange={e => setTxDetectada(t => ({...t, categoria:e.target.value}))} style={smInput}>
                {[...CATEGORIAS_GASTO,...CATEGORIAS_RENDA].filter((v,i,a)=>a.indexOf(v)===i).map(c=><option key={c}>{c}</option>)}
              </select>
            </div>
          </div>

          <div style={{ display:"flex", gap:8, marginTop:12 }}>
            <button onClick={cancelar} style={{ ...btn, flex:1, background:"#1a1710", color:"#706050", border:"1px solid #3a3020", fontSize:12 }}>Cancelar</button>
            <button onClick={iniciarGravacao} style={{ ...btn, flex:1, background:"#1a1020", color:"#9060c0", border:"1px solid #5030a0", fontSize:12 }}>🎙 Nova fala</button>
            <button onClick={confirmar} style={{ ...btn, flex:2, background:"#1a3028", color:"#70d090", border:"1px solid #3a7050", fontSize:12 }}>✓ Confirmar</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── App principal ────────────────────────────────────────────────────────────
export default function App() {
  const [transacoes, setTransacoes] = useState(() => {
    try { const s = localStorage.getItem("fin_transacoes"); return s ? JSON.parse(s) : dadosIniciais.transacoes; } catch { return dadosIniciais.transacoes; }
  });
  const [metas, setMetas] = useState(() => {
    try { const s = localStorage.getItem("fin_metas"); return s ? JSON.parse(s) : dadosIniciais.metas; } catch { return dadosIniciais.metas; }
  });
  const [aba, setAba]               = useState("dashboard");
  const [mesSel, setMesSel]         = useState(mesAtual);
  const [anoSel, setAnoSel]         = useState(anoAtual);
  // Persistência automática no localStorage
  useEffect(() => { try { localStorage.setItem("fin_transacoes", JSON.stringify(transacoes)); } catch {} }, [transacoes]);
  useEffect(() => { try { localStorage.setItem("fin_metas", JSON.stringify(metas)); } catch {} }, [metas]);

  const [form, setForm]             = useState({ tipo:"gasto", descricao:"", valor:"", categoria:CATEGORIAS_GASTO[0], mes:mesAtual, ano:anoAtual });
  const [formMeta, setFormMeta]     = useState({ categoria:CATEGORIAS_GASTO[0], limite:"" });
  const [preview, setPreview]       = useState(null);
  const [parcelas, setParcelas]     = useState([]);
  const [importErro, setImportErro] = useState("");
  const [vozToast, setVozToast]     = useState("");
  const fileRef = useRef();

  // ── Derivados ──
  const txMes       = useMemo(() => transacoes.filter(t => t.mes===mesSel && t.ano===anoSel), [transacoes,mesSel,anoSel]);
  const gastosMes   = useMemo(() => txMes.filter(t => t.tipo==="gasto"),  [txMes]);
  const rendasMes   = useMemo(() => txMes.filter(t => t.tipo==="renda"),  [txMes]);
  const totalGastos = useMemo(() => gastosMes.reduce((s,t)=>s+t.valor,0), [gastosMes]);
  const totalRendas = useMemo(() => rendasMes.reduce((s,t)=>s+t.valor,0), [rendasMes]);
  const saldo       = totalRendas - totalGastos;

  const gastosPorCat = useMemo(() => { const m={}; gastosMes.forEach(t=>{m[t.categoria]=(m[t.categoria]||0)+t.valor;}); return Object.entries(m).map(([name,value])=>({name,value})); }, [gastosMes]);
  const rendasPorCat = useMemo(() => { const m={}; rendasMes.forEach(t=>{m[t.categoria]=(m[t.categoria]||0)+t.valor;}); return Object.entries(m).map(([name,value])=>({name,value})); }, [rendasMes]);

  const evolucao = useMemo(() => MESES.map((m,i) => {
    const tx = transacoes.filter(t=>t.mes===i&&t.ano===anoSel);
    return { mes:m, renda:tx.filter(t=>t.tipo==="renda").reduce((s,t)=>s+t.valor,0), gasto:tx.filter(t=>t.tipo==="gasto").reduce((s,t)=>s+t.valor,0) };
  }), [transacoes,anoSel]);

  const alertasMeta = useMemo(() => metas.map(meta => {
    const gasto = gastosMes.filter(t=>t.categoria===meta.categoria).reduce((s,t)=>s+t.valor,0);
    return { ...meta, gasto, pct: meta.limite>0?(gasto/meta.limite)*100:0 };
  }), [metas,gastosMes]);

  // ── Handlers ──
  const addTx = useCallback(() => {
    if (!form.descricao || !form.valor) return;
    setTransacoes(p=>[...p,{...form, id:gerarId(), valor:parseFloat(form.valor)}]);
    setForm(f=>({...f, descricao:"", valor:""}));
  }, [form]);

  const removeTx = useCallback((id) => setTransacoes(p=>p.filter(t=>t.id!==id)), []);

  const addMeta = useCallback(() => {
    if (!formMeta.limite) return;
    setMetas(p => p.find(m=>m.categoria===formMeta.categoria)
      ? p.map(m=>m.categoria===formMeta.categoria?{...m,limite:parseFloat(formMeta.limite)}:m)
      : [...p,{...formMeta,id:gerarId(),limite:parseFloat(formMeta.limite)}]);
    setFormMeta(f=>({...f,limite:""}));
  }, [formMeta]);

  // Transação detectada por voz
  const onTransacaoVoz = useCallback((tx) => {
    setTransacoes(p => [...p, tx]);
    setVozToast(`✓ "${tx.descricao}" adicionada como ${tx.tipo}!`);
    setTimeout(() => setVozToast(""), 3500);
    // Ajusta filtro de mês/ano para o mês da transação detectada
    setMesSel(tx.mes);
    setAnoSel(tx.ano);
  }, []);

  // ── Import extrato ──
  const handleArquivo = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImportErro("");
    const ext = file.name.split(".").pop().toLowerCase();

    const processarLinhas = (linhas) => {
      if (!linhas||linhas.length===0) { setImportErro("Nenhuma transação encontrada. Verifique o arquivo."); return; }
      const comCat = linhas.map(l => {
        const auto = categorizarAuto(l.memo, l.valor);
        return { id:gerarId(), descricao:l.memo, valor:Math.abs(l.valor), mes:isNaN(l.mes)?mesAtual:l.mes, ano:isNaN(l.ano)?anoAtual:l.ano, tipo:auto.tipo, categoria:auto.categoria, selecionado:true };
      });
      setPreview(comCat);
    };

    if (ext === "pdf") {
      // PDF: lê como ArrayBuffer e processa com pdf.js
      const reader = new FileReader();
      reader.onload = async (ev) => {
        try {
          const pdfjsLib = await carregarPdfJS();
          const pdf = await pdfjsLib.getDocument({ data: ev.target.result }).promise;
          let textoTotal = "";
          for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const ct = await page.getTextContent();
            const itens = ct.items.map(it => ({ texto: it.str.trim(), x: Math.round(it.transform[4]), y: Math.round(it.transform[5]) })).filter(it => it.texto.length > 0);
            const lm = {};
            itens.forEach(it => { if (!lm[it.y]) lm[it.y] = []; lm[it.y].push(it); });
            const ls = Object.values(lm).sort((a,b) => b[0].y - a[0].y).map(its => its.sort((a,b)=>a.x-b.x).map(it=>it.texto).join(" "));
            textoTotal += ls.join("\n") + "\n";
          }
          const linhas = await parsearPDF(ev.target.result);
          processarLinhas(linhas);
          // Extrair parcelas futuras se for fatura Itaú
          const fp = extrairParcelasFuturas(textoTotal);
          if (fp.length > 0) {
            const comCat = fp.map(l => { const auto = categorizarAuto(l.memo, l.valor); return { id:gerarId(), descricao:l.memo, valor:l.valor, mes:l.mes, ano:l.ano, tipo:"gasto", categoria:auto.categoria }; });
            setParcelas(prev => [...prev, ...comCat]);
          }
        } catch(err) { setImportErro("Erro ao processar PDF: "+err.message); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const conteudo = ev.target.result;
          let linhas = null;
          if (ext==="ofx"||ext==="ofc") linhas = parsearOFX(conteudo);
          else if (ext==="csv"||ext==="txt") linhas = parsearCSV(conteudo);
          else { setImportErro("Formato não suportado. Use .pdf, .csv, .txt ou .ofx"); return; }
          processarLinhas(linhas);
        } catch(err) { setImportErro("Erro ao processar: "+err.message); }
      };
      reader.readAsText(file,"utf-8");
    }
    e.target.value="";
  }, []);

  const togglePrev  = (id) => setPreview(p=>p.map(x=>x.id===id?{...x,selecionado:!x.selecionado}:x));
  const updatePrev  = (id,campo,val) => setPreview(p=>p.map(x=>x.id===id?{...x,[campo]:val}:x));
  const confirmar   = () => {
    const sel = preview.filter(p=>p.selecionado).map(({selecionado,...r})=>r);
    setTransacoes(p=>[...p,...sel]);
    setPreview(null); setAba("lancamentos");
  };

  const exportarExcel = useCallback(() => {
    const wb = XLSX.utils.book_new();
    const wsTx = XLSX.utils.json_to_sheet(transacoes.map(t=>({ Tipo:t.tipo==="gasto"?"Gasto":"Renda", Descrição:t.descricao, Categoria:t.categoria, Valor:t.valor, Mês:MESES[t.mes], Ano:t.ano })));
    wsTx["!cols"]=[{wch:10},{wch:30},{wch:22},{wch:14},{wch:8},{wch:8}];
    XLSX.utils.book_append_sheet(wb,wsTx,"Transações");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(evolucao.map(r=>({Mês:r.mes,"Renda (R$)":r.renda,"Gastos (R$)":r.gasto,"Saldo (R$)":r.renda-r.gasto}))), "Resumo Mensal");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(alertasMeta.map(m=>({Categoria:m.categoria,"Limite (R$)":m.limite,"Gasto (R$)":m.gasto,"% Usado":m.pct.toFixed(1)+"%",Status:m.pct>100?"ESTOURADO":m.pct>80?"Atenção":"OK"}))), "Metas");
    XLSX.writeFile(wb, `financas_${anoSel}.xlsx`);
  }, [transacoes,evolucao,alertasMeta,anoSel]);

  // ── Modal revisão extrato ──
  if (preview) {
    const sel = preview.filter(p=>p.selecionado);
    return (
      <div style={{ fontFamily:"'Georgia', serif", background:"#0f0e0c", minHeight:"100vh", color:"#e8dcc8", padding:32 }}>
        <div style={{ maxWidth:960, margin:"0 auto" }}>
          <div style={{ marginBottom:20 }}>
            <div style={{ fontSize:20, color:"#c8a96e", fontWeight:"bold", letterSpacing:2 }}>📋 REVISÃO DO EXTRATO</div>
            <div style={{ fontSize:13, color:"#706050", marginTop:4 }}>{preview.length} transações detectadas · {sel.length} selecionadas</div>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:20 }}>
            {[
              {label:"Gastos detectados", val:fmt(sel.filter(p=>p.tipo==="gasto").reduce((s,p)=>s+p.valor,0)), cor:"#e05040"},
              {label:"Rendas detectadas", val:fmt(sel.filter(p=>p.tipo==="renda").reduce((s,p)=>s+p.valor,0)), cor:"#4a9e6a"},
              {label:"Categ. automáticas",val:`${sel.filter(p=>p.categoria!=="Outros").length}/${sel.length}`, cor:"#c8a96e"},
            ].map(c=>(
              <div key={c.label} style={{ background:"#1a1710", border:"1px solid #2a2418", borderRadius:10, padding:"14px 18px" }}>
                <div style={{ fontSize:10, color:"#706050", letterSpacing:1 }}>{c.label.toUpperCase()}</div>
                <div style={{ fontSize:22, color:c.cor, fontWeight:"bold", marginTop:4 }}>{c.val}</div>
              </div>
            ))}
          </div>
          <div style={{ background:"#1a1710", border:"1px solid #2a2418", borderRadius:12, overflow:"hidden", marginBottom:20 }}>
            <div style={{ display:"grid", gridTemplateColumns:"32px 1fr 110px 150px 110px 80px", gap:8, padding:"10px 16px", borderBottom:"1px solid #2a2418", fontSize:10, color:"#706050", letterSpacing:1 }}>
              <span/><span>DESCRIÇÃO</span><span>VALOR</span><span>CATEGORIA</span><span>TIPO</span><span>MÊS</span>
            </div>
            <div style={{ maxHeight:400, overflowY:"auto" }}>
              {preview.map(p=>(
                <div key={p.id} style={{ display:"grid", gridTemplateColumns:"32px 1fr 110px 150px 110px 80px", gap:8, padding:"9px 16px", borderBottom:"1px solid #1a1810", alignItems:"center", opacity:p.selecionado?1:0.3 }}>
                  <input type="checkbox" checked={p.selecionado} onChange={()=>togglePrev(p.id)} style={{ cursor:"pointer", accentColor:"#c8a96e" }}/>
                  <span style={{ fontSize:12, color:"#e8dcc8", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={p.descricao}>{p.descricao}</span>
                  <span style={{ fontSize:12, fontWeight:"bold", color:p.tipo==="renda"?"#4a9e6a":"#e05040" }}>{fmt(p.valor)}</span>
                  <select value={p.categoria} onChange={e=>updatePrev(p.id,"categoria",e.target.value)} style={smInput}>
                    {[...CATEGORIAS_GASTO,...CATEGORIAS_RENDA].filter((v,i,a)=>a.indexOf(v)===i).map(c=><option key={c}>{c}</option>)}
                  </select>
                  <select value={p.tipo} onChange={e=>updatePrev(p.id,"tipo",e.target.value)} style={{...smInput, color:p.tipo==="renda"?"#4a9e6a":"#e05040"}}>
                    <option value="gasto">💸 Gasto</option>
                    <option value="renda">💰 Renda</option>
                  </select>
                  <span style={{ fontSize:11, color:"#a09070" }}>{MESES[p.mes]}/{p.ano}</span>
                </div>
              ))}
            </div>
          </div>
          <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
            <button onClick={()=>setPreview(null)} style={{...btn, background:"#1a1710", color:"#706050", border:"1px solid #3a3020"}}>Cancelar</button>
            <button onClick={()=>setPreview(p=>p.map(x=>({...x,selecionado:true})))} style={{...btn, background:"#2a2418", color:"#c8a96e", border:"1px solid #604818"}}>Selecionar tudo</button>
            <button onClick={confirmar} style={{...btn, background:"#1a3028", color:"#70d090", border:"1px solid #3a7050"}}>✓ Importar {sel.length} transações</button>
          </div>
        </div>
      </div>
    );
  }

  // ── UI principal ──
  return (
    <div style={{ fontFamily:"'Georgia', serif", background:"#0f0e0c", minHeight:"100vh", color:"#e8dcc8" }}>

      {/* Toast de voz */}
      {vozToast && (
        <div style={{ position:"fixed", top:20, right:20, zIndex:9999, background:"#1a3028", border:"1px solid #3a7050", borderRadius:10, padding:"12px 20px", color:"#70d090", fontSize:14, boxShadow:"0 4px 20px #00000060" }}>
          {vozToast}
        </div>
      )}

      {/* Header */}
      <div style={{ background:"linear-gradient(135deg,#1a1710,#241e14)", borderBottom:"1px solid #3a3020", padding:"18px 32px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:10 }}>
        <div>
          <div style={{ fontSize:20, fontWeight:"bold", color:"#c8a96e", letterSpacing:2 }}>CONTROLE FINANCEIRO</div>
          <div style={{ fontSize:11, color:"#a09070", letterSpacing:4, marginTop:2 }}>PESSOAL & AUTOMATIZADO</div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <select value={mesSel} onChange={e=>setMesSel(+e.target.value)} style={sel}>{MESES.map((m,i)=><option key={i} value={i}>{m}</option>)}</select>
          <select value={anoSel} onChange={e=>setAnoSel(+e.target.value)} style={sel}>{[2024,2025,2026].map(a=><option key={a}>{a}</option>)}</select>
          <input type="file" accept=".pdf,.csv,.txt,.ofx,.ofc" ref={fileRef} style={{ display:"none" }} onChange={handleArquivo}/>
          <button onClick={()=>fileRef.current.click()} style={{...btn, background:"#1a1e2e", color:"#7090e0", border:"1px solid #3050a0"}}>📂 Importar Extrato</button>
          <button onClick={exportarExcel} style={{...btn, background:"#1e3028", color:"#70d090", border:"1px solid #3a7050"}}>⬇ Exportar Excel</button>
        </div>
      </div>
      {importErro && <div style={{ background:"#200a0a", borderBottom:"1px solid #6a2020", padding:"10px 32px", fontSize:13, color:"#e08070" }}>⚠ {importErro}</div>}

      {/* Nav */}
      <div style={{ display:"flex", borderBottom:"1px solid #2a2418", background:"#131108" }}>
        {[["dashboard","📊 Dashboard"],["lancamentos","➕ Lançamentos"],["metas","🎯 Metas"],["parcelas","📅 Parcelas"]].map(([key,label])=>(
          <button key={key} onClick={()=>setAba(key)} style={{ padding:"13px 24px", background:"none", border:"none", color:aba===key?"#c8a96e":"#706050", borderBottom:aba===key?"2px solid #c8a96e":"2px solid transparent", cursor:"pointer", fontSize:13, fontFamily:"inherit", letterSpacing:1 }}>{label}</button>
        ))}
      </div>

      <div style={{ padding:"28px 32px", maxWidth:1200, margin:"0 auto" }}>

        {/* DASHBOARD */}
        {aba==="dashboard" && (
          <div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:16, marginBottom:24 }}>
              <Card label="Total de Rendas" valor={fmt(totalRendas)} cor="#4a9e6a" icon="💰"/>
              <Card label="Total de Gastos" valor={fmt(totalGastos)} cor="#c0392b" icon="💸"/>
              <Card label="Saldo do Mês"    valor={fmt(saldo)} cor={saldo>=0?"#4a9e6a":"#c0392b"} icon={saldo>=0?"✅":"⚠️"} destaque/>
            </div>
            {alertasMeta.some(m=>m.pct>80) && (
              <div style={{ background:"#1a1000", border:"1px solid #8a4000", borderRadius:10, padding:"14px 20px", marginBottom:20 }}>
                <div style={{ color:"#e8a030", fontWeight:"bold", marginBottom:8, fontSize:12, letterSpacing:1 }}>⚠ ALERTAS DE META</div>
                {alertasMeta.filter(m=>m.pct>80).map(m=>(
                  <div key={m.id} style={{ display:"flex", justifyContent:"space-between", fontSize:13, color:m.pct>100?"#e05040":"#e0a040", marginTop:4 }}>
                    <span>{m.categoria}</span>
                    <span>{fmt(m.gasto)} / {fmt(m.limite)} — {m.pct.toFixed(0)}% {m.pct>100?"⚠ ESTOURADO":"⚡ atenção"}</span>
                  </div>
                ))}
              </div>
            )}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20, marginBottom:20 }}>
              <GrafCard title={`Gastos — ${MESES[mesSel]}`}>
                {gastosPorCat.length===0?<Vazio/>:<><ResponsiveContainer width="100%" height={190}><PieChart><Pie data={gastosPorCat} dataKey="value" cx="50%" cy="50%" outerRadius={76}>{gastosPorCat.map((_,i)=><Cell key={i} fill={CORES[i%CORES.length]}/>)}</Pie><Tooltip formatter={v=>fmt(v)} contentStyle={tip}/></PieChart></ResponsiveContainer><Legenda dados={gastosPorCat} cores={CORES}/></>}
              </GrafCard>
              <GrafCard title={`Renda — ${MESES[mesSel]}`}>
                {rendasPorCat.length===0?<Vazio/>:<><ResponsiveContainer width="100%" height={190}><PieChart><Pie data={rendasPorCat} dataKey="value" cx="50%" cy="50%" outerRadius={76}>{rendasPorCat.map((_,i)=><Cell key={i} fill={CORES_RENDA[i%CORES_RENDA.length]}/>)}</Pie><Tooltip formatter={v=>fmt(v)} contentStyle={tip}/></PieChart></ResponsiveContainer><Legenda dados={rendasPorCat} cores={CORES_RENDA}/></>}
              </GrafCard>
            </div>
            <GrafCard title={`Evolução Mensal ${anoSel}`}>
              <ResponsiveContainer width="100%" height={210}>
                <BarChart data={evolucao} barGap={4}>
                  <XAxis dataKey="mes" tick={{fill:"#a09070",fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:"#706050",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`}/>
                  <Tooltip formatter={v=>fmt(v)} contentStyle={tip}/>
                  <Bar dataKey="renda" fill="#4a9e6a" radius={[4,4,0,0]} name="Renda"/>
                  <Bar dataKey="gasto"  fill="#c0392b" radius={[4,4,0,0]} name="Gastos"/>
                </BarChart>
              </ResponsiveContainer>
            </GrafCard>
          </div>
        )}

        {/* LANÇAMENTOS */}
        {aba==="lancamentos" && (
          <div style={{ display:"grid", gridTemplateColumns:"320px 1fr", gap:24 }}>
            <div style={card}>
              <div style={cardTit}>Nova Transação Manual</div>
              <div style={{ display:"flex", gap:6, marginBottom:14 }}>
                {["gasto","renda"].map(t=>(
                  <button key={t} onClick={()=>setForm(f=>({...f,tipo:t,categoria:t==="gasto"?CATEGORIAS_GASTO[0]:CATEGORIAS_RENDA[0]}))} style={{ flex:1, padding:"8px", borderRadius:6, border:"none", cursor:"pointer", fontSize:12, fontFamily:"inherit", background:form.tipo===t?(t==="gasto"?"#4a1010":"#0d3a1f"):"#2a2418", color:form.tipo===t?(t==="gasto"?"#e08070":"#70d090"):"#706050" }}>
                    {t==="gasto"?"💸 Gasto":"💰 Renda"}
                  </button>
                ))}
              </div>
              <Lbl>Descrição</Lbl>
              <input value={form.descricao} onChange={e=>setForm(f=>({...f,descricao:e.target.value}))} style={inp} placeholder="Ex: Conta de luz" onKeyDown={e=>e.key==="Enter"&&addTx()}/>
              <Lbl>Valor (R$)</Lbl>
              <input type="number" value={form.valor} onChange={e=>setForm(f=>({...f,valor:e.target.value}))} style={inp} placeholder="0,00"/>
              <Lbl>Categoria</Lbl>
              <select value={form.categoria} onChange={e=>setForm(f=>({...f,categoria:e.target.value}))} style={inp}>
                {(form.tipo==="gasto"?CATEGORIAS_GASTO:CATEGORIAS_RENDA).map(c=><option key={c}>{c}</option>)}
              </select>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                <div><Lbl>Mês</Lbl><select value={form.mes} onChange={e=>setForm(f=>({...f,mes:+e.target.value}))} style={inp}>{MESES.map((m,i)=><option key={i} value={i}>{m}</option>)}</select></div>
                <div><Lbl>Ano</Lbl><select value={form.ano} onChange={e=>setForm(f=>({...f,ano:+e.target.value}))} style={inp}>{[2024,2025,2026].map(a=><option key={a}>{a}</option>)}</select></div>
              </div>
              <button onClick={addTx} style={{...btn, width:"100%", marginTop:8, background:"#2a1e0a", color:"#c8a96e", border:"1px solid #604818"}}>+ Adicionar</button>

              {/* Componente de Voz */}
              <VozInput onTransacaoDetectada={onTransacaoVoz}/>

              <div style={{ borderTop:"1px solid #2a2418", marginTop:22, paddingTop:18 }}>
                <div style={{ fontSize:11, color:"#c8a96e", letterSpacing:2, marginBottom:10 }}>IMPORTAR EXTRATO</div>
                <div style={{ fontSize:12, color:"#706050", lineHeight:1.7, marginBottom:12 }}>
                  Suporta <b style={{color:"#a09070"}}>.PDF</b> (fatura Itaú/Bradesco), <b style={{color:"#a09070"}}>.CSV</b> (Nubank, Itaú, Bradesco, Inter) e <b style={{color:"#a09070"}}>.OFX</b>.
                </div>
                <button onClick={()=>fileRef.current.click()} style={{...btn, width:"100%", background:"#1a1e2e", color:"#7090e0", border:"1px solid #3050a0"}}>📂 Selecionar arquivo</button>
                <input type="file" accept=".pdf,.csv,.txt,.ofx,.ofc" ref={fileRef} style={{display:"none"}} onChange={handleArquivo}/>
                {importErro && <div style={{ fontSize:12, color:"#e08070", marginTop:8 }}>{importErro}</div>}
              </div>
            </div>

            <div style={card}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
                <div style={cardTit}>Transações — {MESES[mesSel]}/{anoSel}</div>
                <span style={{ fontSize:12, color:"#706050" }}>{txMes.length} lançamentos</span>
              </div>
              <div style={{ maxHeight:520, overflowY:"auto" }}>
                {txMes.length===0?<Vazio/>:txMes.map(t=>(
                  <div key={t.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"10px 0", borderBottom:"1px solid #1e1c14" }}>
                    <div>
                      <div style={{ fontSize:13, color:"#e8dcc8" }}>{t.descricao}</div>
                      <div style={{ fontSize:11, color:"#706050" }}>{t.categoria}</div>
                    </div>
                    <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                      <span style={{ fontWeight:"bold", color:t.tipo==="renda"?"#4a9e6a":"#e05040" }}>{t.tipo==="renda"?"+":"-"}{fmt(t.valor)}</span>
                      <button onClick={()=>removeTx(t.id)} style={{ background:"none", border:"none", color:"#604030", cursor:"pointer", fontSize:16 }}>✕</button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* PARCELAS FUTURAS */}
        {aba==="parcelas" && (
          <div style={card}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
              <div style={cardTit}>📅 Parcelas Futuras de Cartão</div>
              <div style={{ display:"flex", gap:8 }}>
                <span style={{ fontSize:12, color:"#706050" }}>{parcelas.length} parcelas · {fmt(parcelas.reduce((s,p)=>s+p.valor,0))} total</span>
                {parcelas.length > 0 && <button onClick={()=>setParcelas([])} style={{...btn, background:"transparent", color:"#705040", border:"1px solid #3a2818", fontSize:11, padding:"4px 10px"}}>Limpar</button>}
              </div>
            </div>
            {parcelas.length === 0 ? (
              <div style={{ textAlign:"center", padding:"32px 0" }}>
                <div style={{ fontSize:32, marginBottom:12 }}>📄</div>
                <div style={{ color:"#706050", fontSize:13 }}>Nenhuma parcela futura detectada.</div>
                <div style={{ color:"#504030", fontSize:12, marginTop:8 }}>Importe uma fatura PDF do Itaú para extrair automaticamente as parcelas das próximas faturas.</div>
              </div>
            ) : (
              <>
                {/* Agrupar por mês */}
                {Object.entries(
                  parcelas.reduce((acc, p) => {
                    const key = `${p.ano}-${String(p.mes).padStart(2,"0")}`;
                    if (!acc[key]) acc[key] = { label: `${MESES[p.mes]}/${p.ano}`, items:[], total:0 };
                    acc[key].items.push(p);
                    acc[key].total += p.valor;
                    return acc;
                  }, {})
                ).sort(([a],[b]) => a.localeCompare(b)).map(([key, grupo]) => (
                  <div key={key} style={{ marginBottom:20 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", padding:"8px 0", borderBottom:"1px solid #2a2418", marginBottom:8 }}>
                      <span style={{ fontSize:12, color:"#c8a96e", letterSpacing:1 }}>{grupo.label.toUpperCase()}</span>
                      <span style={{ fontSize:12, color:"#e05040", fontWeight:"bold" }}>{fmt(grupo.total)}</span>
                    </div>
                    {grupo.items.map(p => (
                      <div key={p.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"7px 0", borderBottom:"1px solid #1a1810" }}>
                        <div>
                          <div style={{ fontSize:13, color:"#e8dcc8" }}>{p.descricao}</div>
                          <div style={{ fontSize:11, color:"#706050" }}>{p.categoria}</div>
                        </div>
                        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                          <span style={{ color:"#e05040", fontWeight:"bold", fontSize:13 }}>-{fmt(p.valor)}</span>
                          <button onClick={()=>setParcelas(prev=>prev.filter(x=>x.id!==p.id))} style={{ background:"none", border:"none", color:"#604030", cursor:"pointer", fontSize:14 }}>✕</button>
                        </div>
                      </div>
                    ))}
                  </div>
                ))}
                <div style={{ borderTop:"1px solid #2a2418", paddingTop:14, marginTop:8 }}>
                  <button
                    onClick={() => {
                      const novas = parcelas.map(({id,...p}) => ({ ...p, id:gerarId(), tipo:"gasto", descricao:p.descricao }));
                      setTransacoes(prev => [...prev, ...novas]);
                      setParcelas([]);
                      setAba("lancamentos");
                    }}
                    style={{...btn, background:"#2a1e0a", color:"#c8a96e", border:"1px solid #604818"}}
                  >
                    ✓ Lançar todas como transações futuras
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {/* METAS */}
        {aba==="metas" && (
          <div style={{ display:"grid", gridTemplateColumns:"300px 1fr", gap:24 }}>
            <div style={card}>
              <div style={cardTit}>Definir Meta</div>
              <Lbl>Categoria</Lbl>
              <select value={formMeta.categoria} onChange={e=>setFormMeta(f=>({...f,categoria:e.target.value}))} style={inp}>{CATEGORIAS_GASTO.map(c=><option key={c}>{c}</option>)}</select>
              <Lbl>Limite Mensal (R$)</Lbl>
              <input type="number" value={formMeta.limite} onChange={e=>setFormMeta(f=>({...f,limite:e.target.value}))} style={inp} placeholder="0,00"/>
              <button onClick={addMeta} style={{...btn, width:"100%", marginTop:6, background:"#1a2e20", color:"#70d090", border:"1px solid #3d7a50"}}>Salvar Meta</button>
              <div style={{ fontSize:11, color:"#605040", marginTop:10 }}>Categoria já cadastrada? O limite será atualizado.</div>
            </div>
            <div style={card}>
              <div style={cardTit}>Acompanhamento — {MESES[mesSel]}/{anoSel}</div>
              {alertasMeta.length===0?<Vazio/>:alertasMeta.map(m=>{
                const cor = m.pct>100?"#c0392b":m.pct>80?"#e09020":"#4a9e6a";
                return (
                  <div key={m.id} style={{ marginBottom:20 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                      <span style={{ fontSize:14, color:"#e8dcc8" }}>{m.categoria}</span>
                      <span style={{ fontSize:13, color:cor, fontWeight:"bold" }}>{fmt(m.gasto)} / {fmt(m.limite)}</span>
                    </div>
                    <div style={{ height:10, background:"#2a2418", borderRadius:6, overflow:"hidden" }}>
                      <div style={{ height:"100%", width:`${Math.min(m.pct,100)}%`, background:cor, borderRadius:6, transition:"width 0.4s" }}/>
                    </div>
                    <div style={{ display:"flex", justifyContent:"space-between", marginTop:4, fontSize:11, color:"#706050" }}>
                      <span>{m.pct.toFixed(0)}% utilizado</span>
                      <span>{m.pct>100?`⚠ Estourou em ${fmt(m.gasto-m.limite)}`:`Disponível: ${fmt(m.limite-m.gasto)}`}</span>
                    </div>
                  </div>
                );
              })}
              <div style={{ borderTop:"1px solid #2a2418", paddingTop:12 }}>
                <button onClick={()=>setMetas(p=>p.slice(0,-1))} style={{...btn, background:"transparent", color:"#705040", border:"1px solid #3a2818", fontSize:12}}>Remover última meta</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function Card({ label, valor, cor, icon, destaque }) {
  return (
    <div style={{ background:destaque?`${cor}18`:"#1a1710", border:`1px solid ${destaque?cor+"50":"#2a2418"}`, borderRadius:12, padding:"20px 24px" }}>
      <div style={{ fontSize:11, color:"#706050", letterSpacing:2, marginBottom:8 }}>{icon} {label.toUpperCase()}</div>
      <div style={{ fontSize:26, fontWeight:"bold", color:cor }}>{valor}</div>
    </div>
  );
}
function GrafCard({ title, children }) { return <div style={card}><div style={cardTit}>{title}</div>{children}</div>; }
function Legenda({ dados, cores }) {
  return (
    <div style={{ display:"flex", flexWrap:"wrap", gap:"5px 14px", marginTop:10 }}>
      {dados.map((d,i)=><div key={d.name} style={{ display:"flex", alignItems:"center", gap:5, fontSize:11, color:"#a09070" }}><div style={{ width:8, height:8, borderRadius:2, background:cores[i%cores.length] }}/>{d.name}</div>)}
    </div>
  );
}
function Lbl({ children }) { return <div style={{ fontSize:11, color:"#706050", letterSpacing:1, marginBottom:4, marginTop:10 }}>{children}</div>; }
function Vazio() { return <div style={{ textAlign:"center", color:"#504030", padding:"24px 0", fontSize:13 }}>Sem dados para este período</div>; }

const card    = { background:"#1a1710", border:"1px solid #2a2418", borderRadius:12, padding:"20px 24px" };
const cardTit = { fontSize:11, color:"#c8a96e", letterSpacing:2, marginBottom:16, textTransform:"uppercase" };
const inp     = { width:"100%", background:"#111009", border:"1px solid #3a3020", borderRadius:6, padding:"9px 12px", color:"#e8dcc8", fontSize:13, fontFamily:"inherit", boxSizing:"border-box", marginBottom:2 };
const smInput = { background:"#111009", border:"1px solid #3a3020", borderRadius:5, padding:"5px 8px", color:"#e8dcc8", fontFamily:"inherit", width:"100%", fontSize:11 };
const btn     = { padding:"10px 18px", borderRadius:7, border:"none", cursor:"pointer", fontSize:13, fontFamily:"inherit", letterSpacing:1 };
const sel     = { background:"#1a1710", border:"1px solid #3a3020", borderRadius:6, padding:"7px 12px", color:"#c8a96e", fontSize:13, fontFamily:"inherit" };
const tip     = { background:"#1a1710", border:"1px solid #3a3020", borderRadius:6, color:"#e8dcc8", fontSize:12 };
