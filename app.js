// =====================================================================
// GESTÃO DE FROTAS — telas (login, Veículos, Usuários)
// Dados na planilha Frota.xlsx (SharePoint). O site conversa com o fluxo HTTP
// "Frota - API" do Power Automate, que roda o Office Script "Frota - API".
// Login, permissões e o que cada um vê são conferidos no script, não nesta tela.
// =====================================================================
(function () {
  "use strict";
  const $ = s => document.querySelector(s);
  const $$ = s => Array.from(document.querySelectorAll(s));
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const sem = s => String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
  const brl = v => v == null || v === "" ? "—" : Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const msg = (el, t, tipo) => { el.textContent = t || ""; el.className = "msg" + (tipo ? " " + tipo : ""); };

  const CFG = window.FROTA_CONFIG || {};
  const configurado = /^https:\/\/.+/.test(CFG.API_URL || "");
  const CHAVE = "frotaSessao_v1";
  let SESSAO = null;   // { email, token }
  try { SESSAO = JSON.parse(localStorage.getItem(CHAVE) || "null"); } catch (e) { SESSAO = null; }
  const guardarSessao = s => { SESSAO = s; try { s ? localStorage.setItem(CHAVE, JSON.stringify(s)) : localStorage.removeItem(CHAVE); } catch (e) { } };
  let ocupado = 0;
  const carregando = on => { ocupado += on ? 1 : -1; document.body.classList.toggle("ocupado", ocupado > 0); };
  // chama o fluxo do Power Automate. Devolve a resposta (já com os dados novos) ou lança o erro em texto.
  async function api(acao, extra) {
    const corpo = Object.assign({ acao }, SESSAO && acao !== "login" && acao !== "primeiro_acesso" ? { email: SESSAO.email, token: SESSAO.token } : {}, extra || {});
    carregando(true);
    try {
      const resp = await fetch(CFG.API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(corpo) });
      const txt = await resp.text();
      let r; try { r = JSON.parse(txt); } catch (e) { throw new Error("O fluxo respondeu " + resp.status + (txt ? ": " + txt.slice(0, 160) : "")); }
      if (typeof r === "string") { try { r = JSON.parse(r); } catch (e) { } }
      if (r && r.result && typeof r.result === "string") { try { r = JSON.parse(r.result); } catch (e) { } }
      if (!r || !r.ok) {
        if (r && r.codigo === "SESSAO") { guardarSessao(null); EU = null; mostrarLogin(r.erro, "erro"); }
        const pa = r && r.error ? (r.error.message || r.error.code || "") : "";
        throw new Error((r && r.erro) || (pa ? "o fluxo falhou (" + resp.status + "): " + String(pa).slice(0, 200) : "o fluxo respondeu " + resp.status + " sem dados: " + txt.slice(0, 160)));
      }
      if (r.dados) aplicar(r.dados);
      return r;
    } finally { carregando(false); }
  }

  let EU = null;            // linha da tabela usuarios
  let VEICULOS = [];
  const filtro = { contrato: "", gestor: "", status: "", busca: "", garagem: "" };
  // TODAS as colunas do veículo, na ordem da planilha (tabela, cadastro, carga inicial e Excel usam esta lista)
  //   t: num = número/R$ · int = ano · date = data · teste = como reconhecer o título na planilha (sem acento, minúsculo)
  const CAMPOS = [
    { c: "placa", r: "Placa", teste: h => h === "placa" },
    { c: "valor_seguro", r: "Valor seguro", t: "num", teste: h => h.startsWith("valor seguro") || h.startsWith("valor do seguro") },
    { c: "valor_rastreador", r: "Valor rastreador", t: "num", teste: h => h.startsWith("valor rastreador") || h.startsWith("valor do rastreador") },
    { c: "tipo_uso", r: "Tipo de uso", teste: h => h.startsWith("tipo de uso") },
    { c: "matricula_sra", r: "Matrícula base SRA", teste: h => h.startsWith("matricula") },
    { c: "cpf_condutor", r: "CPF", teste: h => h === "cpf" || h.startsWith("cpf ") },
    { c: "condutor", r: "Condutor do veículo", teste: h => h.startsWith("condutor") && !h.includes("mail") },
    { c: "status", r: "Status", teste: h => h === "status" || h.startsWith("status") },
    { c: "tipo_frota", r: "Tipo de frota", teste: h => h.startsWith("tipo de frota") },
    { c: "tipo_veiculo", r: "Tipo de veículo", teste: h => h.startsWith("tipo de veiculo") },
    { c: "proprietario", r: "Proprietário / Locadora", teste: h => h.startsWith("proprietario") || h.includes("locadora") },
    { c: "km_mensal", r: "KM mensal contratado", t: "num", teste: h => h.startsWith("km") },
    { c: "marca", r: "Marca", teste: h => h === "marca" },
    { c: "modelo", r: "Modelo", teste: h => h === "modelo" },
    { c: "cor", r: "Cor", teste: h => h === "cor" },
    { c: "combustivel", r: "Combustível", teste: h => h.startsWith("combustiv") },
    { c: "tipo", r: "Tipo", teste: h => h === "tipo" },
    { c: "centro_custo", r: "Novo centro de custo", teste: h => h.includes("centro de custo") || h.includes("centro custo") },
    { c: "operacao", r: "Operação", teste: h => h === "operacao" || h === "contrato" },
    { c: "gestor_operacao", r: "Gestor operação", teste: h => h.startsWith("gestor") },
    { c: "vertical", r: "Vertical", teste: h => h === "vertical" },
    { c: "ano_fabricacao", r: "Ano de fabricação", t: "int", teste: h => h.startsWith("ano de fab") || h.startsWith("ano fab") },
    { c: "ano_modelo", r: "Ano do modelo", t: "int", teste: h => h.startsWith("ano do mod") || h.startsWith("ano mod") },
    { c: "contrato_inicio", r: "Data início contrato", t: "date", teste: h => h.startsWith("data inicio") || h.startsWith("inicio do contrato") },
    { c: "contrato_fim", r: "Data fim do contrato", t: "date", teste: h => h.startsWith("data fim") || h.startsWith("fim do contrato") },
    { c: "periodo_troca", r: "Período troca frota", teste: h => h.startsWith("periodo") },
    { c: "valor_locacao", r: "Valor locação / compra (R$)", t: "num", teste: h => h.startsWith("valor") && (h.includes("locac") || h.includes("compra")) }
  ];
  const EXTRA_IMPORT = [
    { c: "supervisor", teste: h => h.startsWith("supervisor") || h === "supervisao" },
    { c: "coordenador", teste: h => h.startsWith("coordenador") },
    { c: "condutor_email", teste: h => h.includes("e-mail") || h.includes("email") }
  ];
  const dataBr = d => { if (!d) return ""; const m = String(d).match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : String(d); };
  const mostrar = (f, v) => v == null || v === "" ? "—" : f.t === "num" ? (f.c === "km_mensal" ? Number(v).toLocaleString("pt-BR") : brl(v)) : f.t === "date" ? dataBr(v) : esc(v);
  let OCORRENCIAS = [];
  const fofo = { status: "", tipo: "", busca: "", placas: null, rotulo: "" };
  const TIPOS_OCOR = ["Batida", "Avaria", "Multa", "Roubo/Furto", "Pane mecânica", "Desvio de rota", "Uso fora do horário", "Uso indevido", "Abastecimento irregular", "Outro"];
  const GARAGENS = ["Própria", "Alugada", "Casa do condutor", "Cliente"];
  const dataHora = v => v ? new Date(v).toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "—";
  // data do campo (só tem minutos): se for o minuto de agora, usa a hora exata
  const doCampo = v => { if (!v) return new Date(); const d = new Date(v); return Math.abs(d.getTime() - Date.now()) < 90000 ? new Date() : d; };
  const paraInput = d => { const x = new Date(d.getTime() - d.getTimezoneOffset() * 60000); return x.toISOString().slice(0, 16); };
  const podeGerir = () => EU && (EU.perfil === "admin" || EU.perfil === "gerente");
  const podeSup = () => EU && EU.perfil !== "supervisor";      // supervisor do carro: gerente, coordenador ou admin

  // ------------------------------------------------------------------ login
  function mostrarLogin(texto, tipo) {
    $("#tela-app").hidden = true; $("#tela-login").hidden = false;
    msg($("#login-msg"), texto, tipo);
  }
  if (!configurado) {
    mostrarLogin("Falta colocar o endereço do fluxo (API_URL) no arquivo config.js.", "erro");
    $("#bt-entrar").disabled = true;
    return;
  }

  $("#form-login").addEventListener("submit", async e => {
    e.preventDefault();
    const email = $("#login-email").value.trim().toLowerCase(), senha = $("#login-senha").value;
    msg($("#login-msg"), "Entrando… (pode levar alguns segundos)");
    try { const r = await api("login", { email, senha }); guardarSessao({ email, token: r.token }); await entrar(); }
    catch (err) { msg($("#login-msg"), err.message, "erro"); }
  });
  $("#bt-primeiro").addEventListener("click", async () => {
    const email = $("#login-email").value.trim().toLowerCase(), senha = $("#login-senha").value;
    if (!email || senha.length < 6) return msg($("#login-msg"), "Digite o e-mail e a senha que quer usar (mínimo 6 caracteres) e clique de novo.", "erro");
    msg($("#login-msg"), "Criando a senha…");
    try { const r = await api("primeiro_acesso", { email, senha }); guardarSessao({ email, token: r.token }); await entrar(); }
    catch (err) { msg($("#login-msg"), err.message, "erro"); }
  });
  $("#bt-esqueci").addEventListener("click", () => {
    msg($("#login-msg"), "Peça ao administrador para zerar a sua senha (aba Usuários). Depois entre em Primeiro acesso e crie uma nova.", "ok");
  });
  $("#bt-sair").addEventListener("click", async () => {
    try { await api("sair"); } catch (e) { }
    guardarSessao(null); EU = null; mostrarLogin("");
  });

  async function entrar() {
    $("#tela-login").hidden = true; $("#tela-app").hidden = false;
    try { await api("dados"); }
    catch (err) { if (EU) msg($("#imp-msg"), err.message, "erro"); else mostrarLogin(err.message, "erro"); }
  }
  // tudo o que vem do script (só os carros que esse login pode ver)
  let USUARIOS = [];
  function aplicar(d) {
    const primeiro = !EU;
    EU = d.usuario;
    // guarda tudo e depois aplica o filtro de equipe (barra 👥 do topo) — todas as telas usam as listas filtradas
    TODOS.veic = (d.veiculos || []).filter(v => v.placa).sort((a, b) => String(a.placa).localeCompare(String(b.placa)));
    TODOS.hist = (d.historico || []).filter(h => h.placa).sort((a, b) => String(b.inicio).localeCompare(String(a.inicio)));
    TODOS.ocor = (d.ocorrencias || []).filter(o => o.placa).map(o => Object.assign(o, { id: Number(o.id) })).sort((a, b) => String(b.data_hora).localeCompare(String(a.data_hora)));
    USUARIOS = d.usuarios || [];
    TODOS.hier = (d.hierarquia || []).filter(h => h.colaborador).sort((a, b) => sem(a.colaborador).localeCompare(sem(b.colaborador)));
    TODOS.comb = (d.combustivel || []).filter(x => x.placa && x.mes);
    EXTRATO = d.extrato || [];
    TODOS.cart = (d.cartoes || []).filter(x => x.cartao || x.placa);
    TRAT_M = new Map((d.multas_trat || []).filter(m => m.chave).map(m => [String(m.chave), m]));
    filtrarEquipe();
    if (primeiro) {
      $("#user-nome").textContent = EU.nome;
      $("#user-perfil").textContent = { admin: "Administrador", gerente: "Gerente", coordenador: "Coordenador", supervisor: "Supervisor" }[EU.perfil] || EU.perfil;
      $$("[data-so-admin]").forEach(el => { el.hidden = EU.perfil !== "admin"; });
      $$("[data-so-gerente]").forEach(el => { el.hidden = EU.perfil !== "admin" && EU.perfil !== "gerente"; });
      aplicarAcessos();
      let ult = ""; try { ult = localStorage.getItem("frotaAba_v1") || ""; } catch (e) { }
      if (ult && document.getElementById("aba-" + ult)) setTimeout(() => irPara(ult), 0);
      $("#tela-login").hidden = true; $("#tela-app").hidden = false;
    }
    redesenharTudo();
  }
  function redesenharTudo() {
    montarFiltros(); desenhar();
    if (!$("#aba-rank").hidden) desenharRanking();
    if (!$("#aba-alertas").hidden) desenharAlertasComb();
    $("#q-placa").innerHTML = VEICULOS.map(v => `<option>${esc(v.placa)}</option>`).join("");
    $("#h-q-placa").innerHTML = $("#q-placa").innerHTML;
    if (!$("#q-data").value) $("#q-data").value = paraInput(new Date());
    if (!$("#h-q-data").value) $("#h-q-data").value = paraInput(new Date());
    $("#dl-pessoas").innerHTML += Array.from(new Set(HIST.map(h => h.condutor))).map(c => `<option value="${esc(c)}">`).join("");
    desenharOcorrencias(); desenharHistorico(); desenharInicio();
    if (!CZ.linhas.length) montarMultasSalvas(); else pintarTrat();
    montarFiltrosComb(); desenharComb(); desenharHier();
    if (EU.perfil === "admin") desenharUsuarios();
  }
  // ------------------------------------------------------------------ filtro de equipe (vale para o sistema inteiro)
  const TODOS = { veic: [], hist: [], ocor: [], comb: [], cart: [], hier: [] };
  // combustível: só os carros destes contratos (Operação na base de veículos)
  const CONTRATOS_FROTA = ["Claro Infra Sp", "RATEIO", "Volante", "Equinix Sp4", "Claro Arcon", "Claro Infra Rj", "Comercial", "Cix - Poupatempo Pr",
    "Suprimentos", "Cix - Poupatempo Sp", "Femsa Norte", "Sistemas/gestão", "Operações", "Backoffice", "Cmoc", "Bradesco Spot", "Industrial", "Siemens",
    "Controladoria", "CEO", "Care Plus", "Bresco Vira Copos", "Motorola"];
  const chaveCt = c => sem(c).replace(/[^a-z0-9]+/g, " ").trim();
  const CT_SET = new Set(CONTRATOS_FROTA.map(chaveCt));
  let ESCOPO = true;
  try { ESCOPO = localStorage.getItem("frotaEscopo_v1") !== "0"; } catch (e) { }
  let PLACAS_ESCOPO = new Set();
  const noEscopo = placa => !ESCOPO || PLACAS_ESCOPO.has(String(placa || "").toUpperCase());
  const EQ = { ger: "", coord: "", sup: "" };
  try { Object.assign(EQ, JSON.parse(localStorage.getItem("frotaEquipe_v1") || "{}")); } catch (e) { }
  const equipeAtiva = () => !!(EQ.ger || EQ.coord || EQ.sup);
  // gerente / coordenador / supervisor do carro: cadastro do carro; o que faltar vem da hierarquia do condutor
  function equipeDoCarro(v, mapaHier) {
    const h = mapaHier.get(sem(v.condutor)) || {};
    return { ger: h.gerente || v.gestor_operacao || "", coord: v.coordenador || h.coordenador || "", sup: v.supervisor || h.supervisor || "" };
  }
  function filtrarEquipe() {
    PLACAS_ESCOPO = new Set(TODOS.veic.filter(v => CT_SET.has(chaveCt(v.operacao))).map(v => v.placa));
    const mapaHier = new Map(TODOS.hier.map(h => [sem(h.colaborador), h]));
    const eqs = TODOS.veic.map(v => [v, equipeDoCarro(v, mapaHier)]);
    const bate = (e, comCoord, comSup) => (!EQ.ger || sem(e.ger) === sem(EQ.ger)) && (!comCoord || !EQ.coord || sem(e.coord) === sem(EQ.coord)) && (!comSup || !EQ.sup || sem(e.sup) === sem(EQ.sup));
    const hierEq = h => ({ ger: h.gerente || "", coord: h.coordenador || "", sup: h.supervisor || "" });
    // opções em cascata: coordenadores do gerente escolhido, supervisores do coordenador escolhido
    const todasEq = eqs.map(x => x[1]).concat(TODOS.hier.map(hierEq));
    const lista = (k, f) => Array.from(new Set(todasEq.filter(f).map(e => e[k]).filter(Boolean).map(x => x.replace(/\s+/g, " ").trim()))).sort((a, b) => sem(a).localeCompare(sem(b)));
    const pre = (id, l, todos, k) => { const el = $(id); el.innerHTML = `<option value="">${todos}</option>` + l.map(x => `<option>${esc(x)}</option>`).join(""); el.value = l.find(x => sem(x) === sem(EQ[k])) || ""; EQ[k] = el.value; };
    pre("#eq-ger", lista("ger", () => true), "Todos os gerentes", "ger");
    pre("#eq-coord", lista("coord", e => bate(e, false, false)), "Todos os coordenadores", "coord");
    pre("#eq-sup", lista("sup", e => bate(e, true, false)), "Todos os supervisores", "sup");
    try { localStorage.setItem("frotaEquipe_v1", JSON.stringify(EQ)); } catch (e) { }
    if (!equipeAtiva()) {
      VEICULOS = TODOS.veic; HIST = TODOS.hist; OCORRENCIAS = TODOS.ocor; COMB = TODOS.comb.filter(x => noEscopo(x.placa)); CARTOES = TODOS.cart.filter(x => noEscopo(x.placa)); HIER = TODOS.hier;
    } else {
      VEICULOS = eqs.filter(x => bate(x[1], true, true)).map(x => x[0]);
      const placas = new Set(VEICULOS.map(v => v.placa));
      HIST = TODOS.hist.filter(h => placas.has(h.placa));
      OCORRENCIAS = TODOS.ocor.filter(o => placas.has(o.placa));
      COMB = TODOS.comb.filter(x => placas.has(x.placa));
      CARTOES = TODOS.cart.filter(x => placas.has(String(x.placa || "").toUpperCase()) && noEscopo(x.placa));
      COMB = COMB.filter(x => noEscopo(x.placa));
      HIER = TODOS.hier.filter(h => bate(hierEq(h), true, true));
    }
    $("#eq-limpar").hidden = !equipeAtiva();
    $("#barra-equipe").classList.toggle("on", equipeAtiva());
    $("#eq-info").textContent = equipeAtiva() ? `${VEICULOS.length} carros · ${HIER.length} pessoas na hierarquia` : "";
  }
  [["#eq-ger", "ger"], ["#eq-coord", "coord"], ["#eq-sup", "sup"]].forEach(([id, k]) => $(id).addEventListener("change", e => {
    EQ[k] = e.target.value;
    if (k === "ger") { EQ.coord = ""; EQ.sup = ""; } else if (k === "coord") EQ.sup = "";
    filtrarEquipe(); redesenharTudo();
  }));
  $("#eq-limpar").addEventListener("click", () => { EQ.ger = EQ.coord = EQ.sup = ""; filtrarEquipe(); redesenharTudo(); });

  if (SESSAO && SESSAO.token) entrar(); else mostrarLogin("");

  // ------------------------------------------------------------------ abas
  // menu lateral: grupos abrem/fecham; itens com data-aba trocam a tela; data-acao executa
  $$(".mg").forEach(b => b.addEventListener("click", () => b.parentElement.classList.toggle("aberto")));
  // ícone antes de cada item do menu
  const ICO = {
    inicio: '<path d="M3 11l9-7 9 7"/><path d="M5 10v10h5v-6h4v6h5V10"/>',
    veiculos: '<path d="M5 16l1.5-5.5A2 2 0 0 1 8.4 9h7.2a2 2 0 0 1 1.9 1.5L19 16"/><rect x="3" y="16" width="18" height="4" rx="1"/><circle cx="7.5" cy="20" r="1.5"/><circle cx="16.5" cy="20" r="1.5"/>',
    "nova-placa": '<rect x="3" y="7" width="18" height="10" rx="2"/><path d="M12 9.5v5M9.5 12h5"/>',
    usuarios: '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/><path d="M16 5a3 3 0 0 1 0 6M18 14c2 .8 3 2.8 3 6"/>',
    carga: '<path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 16v4h16v-4"/>',
    hierarquia: '<rect x="9" y="3" width="6" height="5" rx="1"/><rect x="3" y="16" width="6" height="5" rx="1"/><rect x="15" y="16" width="6" height="5" rx="1"/><path d="M12 8v4M6 16v-4h12v4"/>',
    organograma: '<circle cx="12" cy="6" r="3"/><path d="M6 21v-2a6 6 0 0 1 12 0v2"/><path d="M17 11l2 2 3-3"/>',
    trocar: '<path d="M4 7h13l-3-3M20 17H7l3 3"/>',
    historico: '<circle cx="12" cy="12" r="8"/><path d="M12 8v4l3 2"/>',
    ocorrencias: '<path d="M12 4l9 16H3z"/><path d="M12 10v4M12 17v.5"/>',
    cruzar: '<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 8h6M9 12h6M9 16h3"/>',
    combustivel: '<rect x="4" y="4" width="10" height="16" rx="1"/><path d="M4 10h10M14 8l4 3v7a2 2 0 0 1-4 0v-3"/>',
    rank: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    alertas: '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 8 3 8H3s3-1 3-8"/><path d="M10 20a2 2 0 0 0 4 0"/>',
    exportar: '<path d="M12 4v12M7 11l5 5 5-5"/><path d="M4 20h16"/>'
  };
  $$("#menu .mi, #menu .ms").forEach(b => {
    const k = b.dataset.aba || b.dataset.acao || "";
    const txt = b.textContent.trim();
    b.title = txt;
    b.innerHTML = `<i class="mi-ico" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">${ICO[k] || ICO.inicio}</svg></i><span class="mi-txt">${esc(txt)}</span>`;
  });
  // ☰: no celular abre/fecha o menu por cima; no computador esconde/mostra o menu (e lembra a escolha)
  try { if (localStorage.getItem("frotaMenuFechado") === "1") document.body.classList.add("menu-fechado"); } catch (e) { }
  $("#bt-menu").addEventListener("click", () => {
    if (window.innerWidth > 900) {
      const f = document.body.classList.toggle("menu-fechado");
      try { localStorage.setItem("frotaMenuFechado", f ? "1" : "0"); } catch (e) { }
      if (window.Chart && GRAF) setTimeout(() => GRAF.resize(), 50);
    } else document.body.classList.toggle("menu-aberto");
  });
  $("#menu-x").addEventListener("click", () => {
    if (window.innerWidth > 900) {
      const f = document.body.classList.toggle("menu-fechado");
      try { localStorage.setItem("frotaMenuFechado", f ? "1" : "0"); } catch (e) { }
      if (window.Chart && GRAF) setTimeout(() => GRAF.resize(), 50);
    } else document.body.classList.remove("menu-aberto");
  });
  $$(".menu [data-acao]").forEach(b => b.addEventListener("click", () => {
    document.body.classList.remove("menu-aberto");
    const a = b.dataset.acao;
    if (a === "nova-placa") $("#bt-novo-veic").click();
    else if (a === "carga") $("#arq-frota").click();
    else if (a === "organograma") $("#bt-org").click();
    else if (a === "trocar") $("#bt-trocar").click();
    else if (a === "exportar") $("#bt-exportar").click();
  }));
  $$(".aba").forEach(b => b.addEventListener("click", () => {
    document.body.classList.remove("menu-aberto");
    irPara(b.dataset.aba);
  }));

  // ------------------------------------------------------------------ veículos
  $("#th-veiculos").innerHTML = "<tr>" + CAMPOS.map(f => `<th${f.t === "num" ? ' class="num"' : ""}>${esc(f.r)}</th>`).join("") + "<th>Coordenador</th><th>Supervisor</th><th>Garagem</th></tr>";
  const COR_STATUS = s => {
    const t = sem(s);
    if (t.startsWith("ativo")) return "#15803d";
    if (t.includes("manut")) return "#4338ca";
    if (t.includes("funil")) return "#b45309";
    if (t.includes("dispon")) return "#b91c1c";
    return "#475569";
  };

  // lista de status para escolher (os padrões + os que já existem na base)
  const STATUS_PADRAO = ["Ativo", "Manutenção", "Funilaria", "Disponível", "Devolvido"];
  function opcoesStatus(atual) {
    const l = STATUS_PADRAO.slice();
    VEICULOS.forEach(v => { const s = String(v.status || "").trim(); if (s && !l.some(x => sem(x) === sem(s))) l.push(s); });
    const at = l.find(x => sem(x) === sem(atual || "")) || atual;
    if (at && l.indexOf(at) < 0) l.unshift(at);
    return l.map(s => `<option ${s === at ? "selected" : ""}>${esc(s)}</option>`).join("");
  }
  function trocarStatus(placa) {
    const v = VEICULOS.find(x => x.placa === placa); if (!v) return;
    abrirModal(`<h2>Status · ${esc(v.placa)}</h2>
      <p style="margin:0 0 10px">${esc([v.marca, v.modelo].filter(Boolean).join(" "))} · condutor: <b>${esc(v.condutor) || "—"}</b></p>
      <p style="margin:0 0 10px">Status atual: ${v.status ? `<span class="st" style="background:${COR_STATUS(v.status)}">${esc(v.status)}</span>` : "—"}</p>
      <div class="linha-bt"><select id="ts-status">${opcoesStatus(v.status)}</select>
        <button class="bt bt-pri" id="ts-salva">Alterar status</button><span id="ts-msg" class="msg"></span></div>`);
    $("#ts-salva").addEventListener("click", async () => {
      const st = $("#ts-status").value;
      if (st === v.status) return msg($("#ts-msg"), "O carro já está com esse status.", "erro");
      msg($("#ts-msg"), "Gravando…");
      try { await api("salvar_status", { placa: v.placa, status: st }); fecharModal(); }
      catch (err) { msg($("#ts-msg"), err.message, "erro"); }
    });
  }

  function opcoes(sel, valores, primeiro) {
    const atual = sel.value;
    sel.innerHTML = `<option value="">${esc(primeiro)}</option>` + valores.map(v => `<option>${esc(v)}</option>`).join("");
    sel.value = valores.includes(atual) ? atual : "";
  }
  // gestor × contrato em cascata: escolhendo o gestor, a lista de contratos mostra só os dele (e vice-versa)
  const uniqDe = (lista, campo) => Array.from(new Set(lista.map(v => String(v[campo] || "").trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, "pt-BR"));
  const contratosDo = g => uniqDe(VEICULOS.filter(v => !g || String(v.gestor_operacao || "").trim() === g), "operacao");
  const gestoresDo = c => uniqDe(VEICULOS.filter(v => !c || String(v.operacao || "").trim() === c), "gestor_operacao");
  const unicos = campo => Array.from(new Set(VEICULOS.map(v => (v[campo] || "").trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, "pt-BR"));
  function montarFiltros() {
    opcoes($("#f-contrato"), contratosDo($("#f-gestor").value), "Todos os contratos");
    opcoes($("#f-gestor"), gestoresDo($("#f-contrato").value), "Todos os gestores");
    filtro.contrato = $("#f-contrato").value; filtro.gestor = $("#f-gestor").value;
    opcoes($("#f-status"), unicos("status"), "Todos os status");
    const dl = (id, campos, extra) => { $(id).innerHTML = Array.from(new Set([].concat(...campos.map(unicos), extra || []))).sort((a, b) => a.localeCompare(b, "pt-BR")).map(v => `<option value="${esc(v)}">`).join(""); };
    dl("#dl-operacao", ["operacao"]); dl("#dl-gestor", ["gestor_operacao"]);
    dl("#dl-status", ["status"], ["Ativo", "Manutenção", "Funilaria", "Disponível", "Devolvido"]);
    dl("#dl-tipo", ["tipo"]); dl("#dl-comb", ["combustivel"], ["Flex", "Diesel", "Gasolina", "Etanol", "Elétrico"]);
    dl("#dl-pessoas", ["supervisor", "coordenador", "gestor_operacao"]);
    $("#lista-equipes").innerHTML = Array.from(new Set(unicos("gestor_operacao").concat(unicos("coordenador"), unicos("supervisor")))).map(v => `<option value="${esc(v)}">`).join("");
  }
  [["#f-contrato", "contrato"], ["#f-gestor", "gestor"], ["#f-status", "status"]].forEach(([id, k]) =>
    $(id).addEventListener("change", e => { filtro[k] = e.target.value; if (k !== "status") montarFiltros(); desenhar(); }));
  $("#f-busca").addEventListener("input", e => { filtro.busca = sem(e.target.value); desenhar(); });
  $("#f-garagem").addEventListener("change", e => { filtro.garagem = e.target.value; desenhar(); });

  // ---------- oficina: condutor "OFICINA" (ou "Oficina - nome", "Manutenção") = carro parado em manutenção
  const ehOficina = nome => /\boficina\b|manutenc/.test(sem(nome));
  const diasDesde = iso => { const t = new Date(iso).getTime(); return isNaN(t) ? null : Math.max(0, Math.floor((Date.now() - t) / 86400000)); };
  // períodos em que cada placa esteve na oficina (pela linha do tempo)
  function periodosOficina() {
    const m = new Map();
    TODOS.hist.filter(h => ehOficina(h.condutor)).forEach(h => {
      const ini = new Date(h.inicio).getTime(), fim = h.fim ? new Date(h.fim).getTime() : Date.now();
      if (isNaN(ini)) return;
      if (!m.has(h.placa)) m.set(h.placa, []);
      m.get(h.placa).push([ini, fim, h.condutor]);
    });
    return m;
  }
  function filtrados(ignorarStatus) {
    return VEICULOS.filter(v =>
      (!filtro.contrato || v.operacao === filtro.contrato) &&
      (!filtro.gestor || v.gestor_operacao === filtro.gestor) &&
      (ignorarStatus || !filtro.status || (filtro.status === "__oficina" ? ehOficina(v.condutor) : v.status === filtro.status)) &&
      (!filtro.garagem || (filtro.garagem === "sem" ? !v.garagem_tipo : v.garagem_tipo === filtro.garagem)) &&
      (!filtro.busca || sem([v.placa, v.condutor, v.modelo, v.marca, v.operacao, v.gestor_operacao, v.supervisor, v.coordenador, v.garagem_endereco].join(" ")).includes(filtro.busca)));
  }

  function desenhar() {
    const base = filtrados(true);
    const porStatus = {};
    base.forEach(v => { const s = v.status || "Sem status"; porStatus[s] = (porStatus[s] || 0) + 1; });
    const kp = [{ r: "Total de veículos", v: base.length, cor: "#0e7c86", s: "" }]
      .concat(Object.keys(porStatus).sort((a, b) => porStatus[b] - porStatus[a]).map(s => ({ r: s, v: porStatus[s], cor: COR_STATUS(s), s })))
      .concat([{ r: "🔧 Na oficina", v: base.filter(v => ehOficina(v.condutor)).length, cor: "#7c3aed", s: "__oficina" }]);
    $("#kpis").innerHTML = kp.map(k => `<div class="kpi${filtro.status === k.s && k.s ? " on" : ""}" data-s="${esc(k.s)}" style="border-left-color:${k.cor}"><small>${esc(k.r)}</small><b>${k.v.toLocaleString("pt-BR")}</b></div>`).join("");
    $$("#kpis .kpi").forEach(el => el.addEventListener("click", () => {
      filtro.status = filtro.status === el.dataset.s ? "" : el.dataset.s;
      $("#f-status").value = filtro.status; desenhar();
    }));
    // ALERTAS: carro sem garagem (supervisor preenche) e sem supervisor/coordenador (gerente preenche)
    const semGaragem = VEICULOS.filter(v => !v.garagem_tipo).length;
    const semEquipe = VEICULOS.filter(v => !v.supervisor || (podeGerir() && !v.coordenador)).length;
    $("#alertas").innerHTML =
      (semGaragem ? `<div class="alerta" data-alerta="garagem">⚠ <b>${semGaragem}</b> ${semGaragem === 1 ? "carro está" : "carros estão"} sem o local onde fica guardado (garagem própria ou alugada + endereço). Clique para ver e preencher.</div>` : "") +
      (semEquipe && podeSup() ? `<div class="alerta amarelo" data-alerta="equipe">⚠ <b>${semEquipe}</b> ${semEquipe === 1 ? "carro está" : "carros estão"} sem supervisor ou coordenador. Abra o carro e preencha.</div>` : "");
    $$("#alertas .alerta").forEach(a => a.addEventListener("click", () => {
      if (a.dataset.alerta === "garagem") { filtro.garagem = "sem"; $("#f-garagem").value = "sem"; }
      else { filtro.busca = ""; $("#f-busca").value = ""; }
      desenhar();
      if (a.dataset.alerta === "equipe") { const v = filtrados(false).find(x => !x.supervisor || !x.coordenador); if (v) abrirVeiculo(v.placa); }
    }));
    const lista = filtrados(false);
    $("#f-info").textContent = `${lista.length} de ${VEICULOS.length} veículos`;
    $("#tb-veiculos").innerHTML = lista.length ? lista.map(v => `<tr class="clica" data-placa="${esc(v.placa)}" title="Clique para abrir o carro">` +
      CAMPOS.map(f => {
        if (f.c === "placa") return `<td><b>${esc(v.placa)}</b></td>`;
        if (f.c === "condutor" && ehOficina(v.condutor)) { const d = diasDesde(v.condutor_desde); return `<td><span class="tag-oficina">🔧 ${esc(v.condutor)}</span>${d != null ? `<small class="resp">há ${d} ${d === 1 ? "dia" : "dias"}</small>` : ""}</td>`; }
        if (f.c === "status") return `<td><span class="st st-troca" data-st="${esc(v.placa)}" title="Clique para alterar o status" style="background:${v.status ? COR_STATUS(v.status) : "#94a3b8"}">${esc(v.status) || "definir"} ▾</span></td>`;
        return `<td${f.t === "num" ? ' class="num"' : ""}>${mostrar(f, v[f.c])}</td>`;
      }).join("") +
      `<td>${esc(v.coordenador) || '<span style="color:#b45309;font-weight:700">preencher</span>'}</td>
      <td>${esc(v.supervisor) || '<span style="color:#b45309;font-weight:700">preencher</span>'}</td>
      <td>${v.garagem_tipo ? `${esc(v.garagem_tipo)}<small style="display:block;color:#475569">${esc(v.garagem_endereco || "")}</small>` : '<span style="color:#b91c1c;font-weight:800">⚠ informar</span>'}</td></tr>`).join("")
      : `<tr><td colspan="${CAMPOS.length + 3}" style="text-align:center;color:#475569;padding:24px">${VEICULOS.length ? "Nenhum veículo com esses filtros." : (EU && EU.perfil === "admin" ? "Nenhum veículo ainda. Clique em Carga inicial (planilha) para trazer a base de hoje, ou em + Nova placa." : "Nenhum veículo da sua equipe cadastrado ainda.")}</td></tr>`;
  }

  $("#tb-veiculos").addEventListener("click", e => {
    const st = e.target.closest("[data-st]"); if (st) { trocarStatus(st.dataset.st); return; }
    const tr = e.target.closest("tr[data-placa]"); if (tr) abrirVeiculo(tr.dataset.placa);
  });

  // ------------------------------------------------------------------ início (resumo)
  function alertaTransInicio(mes, placas, g, ct) {
    if (!TRANS_C[mes]) { garantirTrans(mes, () => desenharInicio()); return ""; }
    const rows = TRANS_C[mes].filter(t => noEscopo(t.placa) && ((!g && !ct && !equipeAtiva()) || placas.has(t.placa)));
    const al = alertasTrans(rows);
    const mais2 = al.filter(a => a.dia2).length, acima = al.reduce((s, a) => s + a.acima, 0);
    const ofic = al.reduce((s, a) => s + (a.oficina || 0), 0);
    if (!mais2 && !acima && !ofic) return "";
    return `<div class="alerta" data-ir="combustivel" data-acesso="combustivel">⛽ <b>${mais2}</b> ${mais2 === 1 ? "vez" : "vezes"} um carro abasteceu 2+ vezes no mesmo dia e <b>${acima}</b> ${acima === 1 ? "abastecimento" : "abastecimentos"} acima de 80 L${ofic ? ` e <b>${ofic}</b> com o carro na oficina` : ""} em ${nomeMes(mes)}. Clique para ver.</div>`;
  }
  function desenharInicio() {
    $("#ini-data").textContent = "Atualizado em " + new Date().toLocaleString("pt-BR");
    // filtro de gestor e contrato vale para o Início inteiro (carros, ocorrências, multas, trocas e combustível)
    const g = opSel("#ic-gestor", gestoresDo($("#ic-contrato").value), "Todos os gestores");
    const ct = opSel("#ic-contrato", contratosDo(g), "Todos os contratos");
    const V = VEICULOS.filter(v => (!g || v.gestor_operacao === g) && (!ct || v.operacao === ct));
    const placas = new Set(V.map(v => v.placa));
    const O = (g || ct) ? OCORRENCIAS.filter(o => placas.has(o.placa)) : OCORRENCIAS;
    const H = (g || ct) ? HIST.filter(h => placas.has(h.placa)) : HIST;
    const semGar = V.filter(v => !v.garagem_tipo).length;
    const semEq = V.filter(v => !v.supervisor || !v.coordenador).length;
    const abertas = O.filter(o => o.status !== "Encerrada");
    const trinta = Date.now() - 30 * 86400000;
    const trocas = H.filter(h => new Date(h.inicio).getTime() >= trinta && h.origem !== "planilha");
    const ativos = V.filter(v => sem(v.status).startsWith("ativo")).length;
    const hoje = new Date(), mesAt = hoje.getFullYear() + "-" + String(hoje.getMonth() + 1).padStart(2, "0");
    const doMesO = O.filter(o => String(o.data_hora || "").slice(0, 7) === mesAt);
    const multas = doMesO.filter(o => o.tipo === "Multa").length, batidas = doMesO.filter(o => o.tipo === "Batida" || o.tipo === "Avaria").length;
    $("#ini-alertas").innerHTML =
      (semGar ? `<div class="alerta" data-ir="garagem">⚠ <b>${semGar}</b> carros sem o local onde ficam guardados. Clique para preencher.</div>` : "") +
      (semEq && podeSup() ? `<div class="alerta amarelo" data-ir="veiculos">⚠ <b>${semEq}</b> carros sem supervisor ou coordenador.</div>` : "") +
      (abertas.length ? `<div class="alerta amarelo" data-ir="ocorrencias">⚠ <b>${abertas.length}</b> ocorrências abertas aguardando tratativa.</div>` : "") + alertaTransInicio(mesAt, placas, g, ct);
    // ao abrir Veículos ou Ocorrências, leva o mesmo filtro
    const levar = destino => {
      if (destino === "oficina") { filtro.gestor = g; filtro.contrato = ct; $("#f-gestor").value = g; $("#f-contrato").value = ct; filtro.status = "__oficina"; irPara("veiculos"); desenhar(); return; }
      if (destino === "veiculos" || destino === "garagem") {
        filtro.gestor = g; filtro.contrato = ct; $("#f-gestor").value = g; $("#f-contrato").value = ct;
        if (destino === "garagem") { filtro.garagem = "sem"; $("#f-garagem").value = "sem"; }
        irPara("veiculos"); desenhar();
      } else if (destino === "combustivel") { fcb.mes = mesAt; $("#cb-gestor").value = g; $("#cb-contrato").value = ct; montarFiltrosComb(); desenharComb(); irPara("combustivel"); setTimeout(() => { const el = $("#cb-trans-caixa"); if (el && !el.hidden) el.scrollIntoView({ behavior: "smooth" }); }, 100);
      } else if (destino === "ocorrencias") { fofo.placas = (g || ct) ? placas : null; fofo.rotulo = [g, ct].filter(Boolean).join(" · "); irPara("ocorrencias"); desenharOcorrencias(); }
      else irPara(destino);
    };
    $$("#ini-alertas .alerta").forEach(a => a.addEventListener("click", () => levar(a.dataset.ir)));
    const naOficina = V.filter(v => ehOficina(v.condutor));
    const oficinaLonga = naOficina.filter(v => (diasDesde(v.condutor_desde) || 0) > 7);
    if (oficinaLonga.length) $("#ini-alertas").insertAdjacentHTML("beforeend", `<div class="alerta amarelo" data-ir="oficina">🔧 <b>${oficinaLonga.length}</b> ${oficinaLonga.length === 1 ? "carro está" : "carros estão"} na oficina há mais de 7 dias.</div>`);
    $$("#ini-alertas .alerta[data-ir=oficina]").forEach(a => a.addEventListener("click", () => levar("oficina")));
    const k = [["Veículos", V.length, "#0e7c86", "veiculos"], ["Ativos", ativos, "#15803d", "veiculos"], ["🔧 Na oficina", naOficina.length, naOficina.length ? "#7c3aed" : "#15803d", "oficina"],
      ["Sem garagem", semGar, semGar ? "#b91c1c" : "#15803d", "garagem"], ["Ocorrências abertas", abertas.length, abertas.length ? "#b45309" : "#15803d", "ocorrencias"],
      ["Batidas/avarias no mês", batidas, batidas ? "#b45309" : "#15803d", "ocorrencias"],
      ["Trocas de condutor (30 dias)", trocas.length, "#4338ca", "historico"]];
    $("#ini-kpis").innerHTML = k.map(x => `<div class="kpi" data-ir="${x[3]}" style="border-left-color:${x[2]}"><small>${x[0]}</small><b>${x[1].toLocaleString("pt-BR")}</b></div>`).join("");
    $$("#ini-kpis .kpi").forEach(el => el.addEventListener("click", () => levar(el.dataset.ir)));
    const porC = {}; V.forEach(v => { const c = v.operacao || "Sem contrato"; porC[c] = (porC[c] || 0) + 1; });
    const cs = Object.keys(porC).sort((a, b) => porC[b] - porC[a]); const mx = Math.max(1, ...cs.map(c => porC[c]));
    $("#ini-contratos").innerHTML = cs.length ? cs.map(c => `<div class="barra-c"><span title="${esc(c)}">${esc(c)}</span><div class="b" style="width:${porC[c] / mx * 100}%"></div><b>${porC[c]}</b></div>`).join("") : '<p class="info">Nenhum veículo com esse filtro.</p>';
    $("#ini-trocas").innerHTML = H.filter(h => h.origem !== "planilha").slice(0, 8).map(h => `<li><b>${esc(h.placa)}</b><span>${esc(h.condutor)}</span><span class="info">${dataHora(h.inicio)}</span></li>`).join("") || '<li class="info">Nenhuma troca ainda.</li>';
    desenharInicioComb();
    $("#ini-ocor").innerHTML = abertas.slice(0, 8).map(o => `<li><b>${esc(o.placa)} · ${esc(o.tipo)}</b><span>${esc(o.condutor || "condutor não identificado")}</span><span class="info">${dataHora(o.data_hora)}</span></li>`).join("") || '<li class="info">Nenhuma ocorrência aberta.</li>';
  }

  // ------------------------------------------------------------------ o que cada usuário pode acessar
  const ACESSOS = { inicio: "Início", veiculos: "Veículos", trocar: "Trocar condutor", historico: "Histórico", ocorrencias: "Ocorrências", multas: "Multas e batidas",
    combustivel: "Combustível", hierarquia: "Hierarquia", exportar: "Baixar Excel" };
  const ABA_ACESSO = { inicio: "inicio", rank: "combustivel", alertas: "combustivel", veiculos: "veiculos", historico: "historico", ocorrencias: "ocorrencias", cruzar: "multas", combustivel: "combustivel", hierarquia: "hierarquia" };
  const listaAcessos = u => String((u && u.acessos) || "").split(",").map(x => x.trim()).filter(Boolean);
  const pode = k => !EU || EU.perfil === "admin" || !listaAcessos(EU).length || listaAcessos(EU).indexOf(k) >= 0;
  function aplicarAcessos() {
    Object.keys(ACESSOS).forEach(k => document.body.classList.toggle("sem-" + k, !pode(k)));
    // grupo do menu sem nenhum item liberado some
    $$("#menu .grupo").forEach(g => { g.hidden = !Array.from(g.querySelectorAll(".sub button")).some(b => getComputedStyle(b).display !== "none" && !b.hidden); });
    const ativa = $$(".aba-conteudo").find(c => !c.hidden);
    const id = ativa ? ativa.id.replace("aba-", "") : "inicio";
    if (ABA_ACESSO[id] && !pode(ABA_ACESSO[id])) irPara(primeiraAba());
  }
  // primeira tela que a pessoa pode abrir (na ordem do menu)
  function primeiraAba() {
    const ordem = ["inicio", "veiculos", "hierarquia", "historico", "ocorrencias", "cruzar", "combustivel"];
    return ordem.find(a => pode(ABA_ACESSO[a])) || "inicio";
  }
  function irPara(aba) {
    if (ABA_ACESSO[aba] && !pode(ABA_ACESSO[aba])) aba = primeiraAba();
    $$(".aba").forEach(x => x.classList.toggle("on", x.dataset.aba === aba));
    $$(".aba-conteudo").forEach(c => { c.hidden = c.id !== "aba-" + aba; });
    // título do topo e da aba do navegador = tela aberta
    const bt = document.querySelector(`#menu [data-aba="${aba}"]`);
    const nome = bt ? bt.textContent.trim() : "Início";
    $("#topo-titulo").textContent = nome; document.title = nome + " · Frota";
    try { localStorage.setItem("frotaAba_v1", aba); } catch (e) { }   // ao atualizar a página, volta para a mesma tela
    if (aba === "rank") desenharRanking();
    if (aba === "alertas") desenharAlertasComb();
  }

  // ------------------------------------------------------------------ janela
  function abrirModal(html) { $("#modal-corpo").innerHTML = html; $("#modal").hidden = false; }
  function fecharModal() { $("#modal").hidden = true; $("#modal-corpo").innerHTML = ""; }
  $("#modal-x").addEventListener("click", fecharModal);
  $("#modal").addEventListener("click", e => { if (e.target.id === "modal") fecharModal(); });
  document.addEventListener("keydown", e => { if (e.key === "Escape" && !$("#modal").hidden) fecharModal(); });

  // ------------------------------------------------------------------ detalhe do veículo
  async function abrirVeiculo(placa) {
    const v = VEICULOS.find(x => x.placa === placa); if (!v) return;
    const gerir = podeGerir();
    abrirModal(`<h2>${esc(v.placa)} · ${esc([v.marca, v.modelo].filter(Boolean).join(" "))}</h2>
      <div class="dados">
        <div><small>Contrato</small>${esc(v.operacao) || "—"}</div><div><small>Gestor operação</small>${esc(v.gestor_operacao) || "—"}</div>
        <div><small>Status</small>${v.status ? `<span class="st" style="background:${COR_STATUS(v.status)}">${esc(v.status)}</span>` : "—"}</div><div><small>Tipo</small>${esc(v.tipo) || "—"}</div>
      </div>
      <h3>Status</h3>
      <div class="linha-bt"><select id="m-status">${opcoesStatus(v.status)}</select>
        <button class="bt bt-pri" id="m-salva-status">Alterar status</button><span id="m-msg-st" class="msg"></span></div>
      ${EU.perfil === "admin" ? `<h3>Dados do veículo</h3>${camposVeiculo(v, false)}
      <div class="linha-bt"><button class="bt bt-pri" id="m-salva-dados">Salvar dados do veículo</button><span id="m-msg0" class="msg"></span></div>` : ""}
      <h3>Condutor</h3>
      <p style="margin:0 0 8px">Com o carro agora: <b>${esc(v.condutor) || "ninguém registrado"}</b>${v.condutor_desde ? " · desde " + dataHora(v.condutor_desde) : ""}</p>
      <div class="campos">
        <label>Novo condutor<input id="m-condutor" placeholder="Nome do motorista"></label>
        <label>Está com o carro desde<input id="m-desde" type="datetime-local" value="${paraInput(new Date())}"></label>
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="m-salva-condutor">Registrar troca de condutor</button><span id="m-msg1" class="msg"></span></div>
      <h3>Supervisor e coordenador ${gerir ? "" : '<small style="font-weight:500;color:#475569">(o gerente preenche' + (podeSup() ? "; o coordenador preenche o supervisor" : "") + ')</small>'}</h3>
      <div class="campos">
        <label>Supervisor<input id="m-sup" value="${esc(v.supervisor || "")}" ${podeSup() ? "" : "disabled"}></label>
        <label>Coordenador<input id="m-coord" value="${esc(v.coordenador || "")}" ${gerir ? "" : "disabled"}></label>
      </div>
      ${podeSup() ? '<div class="linha-bt"><button class="bt bt-pri" id="m-salva-equipe">Salvar ' + (gerir ? "supervisor e coordenador" : "supervisor") + '</button><span id="m-msg2" class="msg"></span></div>' : ""}
      <h3>Onde o carro fica guardado</h3>
      <div class="campos">
        <label>Garagem<select id="m-gar"><option value="">— escolha —</option>${GARAGENS.map(g => `<option ${v.garagem_tipo === g ? "selected" : ""}>${g}</option>`).join("")}</select></label>
        <label class="larga">Endereço<input id="m-end" value="${esc(v.garagem_endereco || "")}" placeholder="Rua, número, bairro, cidade"></label>
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="m-salva-gar">Salvar garagem</button><span id="m-msg3" class="msg"></span>
        ${v.garagem_por ? `<span class="info">última alteração: ${esc(v.garagem_por)} · ${dataHora(v.garagem_em)}</span>` : ""}</div>
      <h3>Linha do tempo do condutor</h3><ul class="tempo" id="m-tempo"><li>Carregando…</li></ul>
      <h3>Ocorrências deste carro</h3><div id="m-ocor">Carregando…</div>
      <div class="linha-bt"><button class="bt bt-sec" id="m-nova-ocor">+ Registrar ocorrência deste carro</button></div>`);

    if (EU.perfil === "admin") $("#m-salva-dados").addEventListener("click", async () => {
      const r = lerCamposVeiculo(false); if (typeof r === "string") return msg($("#m-msg0"), r, "erro");
      try { await api("salvar_veiculo", { dados: Object.assign({ placa: v.placa }, r) }); msg($("#m-msg0"), "✓ Salvo", "ok"); }
      catch (err) { msg($("#m-msg0"), err.message, "erro"); }
    });
    $("#m-salva-status").addEventListener("click", async () => {
      const st = $("#m-status").value;
      if (st === v.status) return msg($("#m-msg-st"), "O carro já está com esse status.", "erro");
      msg($("#m-msg-st"), "Gravando…");
      try { await api("salvar_status", { placa: v.placa, status: st }); abrirVeiculo(v.placa); }
      catch (err) { msg($("#m-msg-st"), err.message, "erro"); }
    });
    $("#m-salva-condutor").addEventListener("click", async () => {
      const nome = $("#m-condutor").value.trim(), desde = $("#m-desde").value;
      if (!nome) return msg($("#m-msg1"), "Digite o nome do novo condutor.", "erro");
      msg($("#m-msg1"), "Gravando…");
      try { await api("trocar_condutor", { placa: v.placa, condutor: nome, desde: doCampo(desde).toISOString() }); abrirVeiculo(v.placa); }
      catch (err) { msg($("#m-msg1"), err.message, "erro"); }
    });
    if (podeSup()) $("#m-salva-equipe").addEventListener("click", async () => {
      const mud = { supervisor: $("#m-sup").value.trim() || null };
      if (gerir) mud.coordenador = $("#m-coord").value.trim() || null;
      try { await api("salvar_equipe", { placa: v.placa, dados: mud }); msg($("#m-msg2"), "✓ Salvo", "ok"); }
      catch (err) { msg($("#m-msg2"), err.message, "erro"); }
    });
    $("#m-salva-gar").addEventListener("click", async () => {
      const tipo = $("#m-gar").value, end = $("#m-end").value.trim();
      if (!tipo || !end) return msg($("#m-msg3"), "Escolha a garagem e escreva o endereço.", "erro");
      try { await api("salvar_garagem", { placa: v.placa, dados: { garagem_tipo: tipo, garagem_endereco: end } }); msg($("#m-msg3"), "✓ Salvo", "ok"); }
      catch (err) { msg($("#m-msg3"), err.message, "erro"); }
    });
    $("#m-nova-ocor").addEventListener("click", () => novaOcorrencia(v.placa));

    if (!$("#m-tempo")) return;
    const h = HIST.filter(x => x.placa === v.placa);
    $("#m-tempo").innerHTML = h.length ? h.map(x => `<li><b>${esc(x.condutor)}</b><span>${dataHora(x.inicio)} → ${x.fim ? dataHora(x.fim) : '<span class="atual">com o carro agora</span>'}</span><span class="info">registrado por ${esc(x.registrado_por || "—")}</span></li>`).join("")
      : "<li>Nenhuma troca registrada ainda.</li>";
    const o = OCORRENCIAS.filter(x => x.placa === v.placa);
    $("#m-ocor").innerHTML = o.length ? `<ul class="tempo">${o.map(x => `<li><b>${dataHora(x.data_hora)} · ${esc(x.tipo)}</b><span>${esc(x.descricao)} — condutor: ${esc(x.condutor || "não identificado")}</span><span class="info">${esc(x.status)}</span></li>`).join("")}</ul>` : "Nenhuma ocorrência.";
  }

  // ------------------------------------------------------------------ cadastro de veículo (admin)
  const LISTAS = { operacao: "dl-operacao", gestor_operacao: "dl-gestor", status: "dl-status", tipo: "dl-tipo", combustivel: "dl-comb" };
  const OBRIG = ["placa", "operacao", "gestor_operacao", "status"];
  function camposVeiculo(v, novo) {
    v = v || {};
    const val = f => { const x = v[f.c]; if (x == null) return f.c === "status" && novo ? "Ativo" : ""; return f.t === "num" ? String(x).replace(".", ",") : String(x); };
    return `<div class="campos">` + CAMPOS.filter(f => novo || f.c !== "placa").filter(f => f.c !== "condutor" || novo).map(f =>
      `<label>${esc(f.r)}${OBRIG.includes(f.c) ? " *" : ""}<input id="cv-${f.c}" value="${esc(val(f))}"
        ${f.t === "date" ? 'type="date"' : f.t ? 'inputmode="decimal"' : ""} ${LISTAS[f.c] ? `list="${LISTAS[f.c]}"` : ""}
        ${f.c === "placa" ? 'placeholder="ABC1D23" maxlength="8" style="text-transform:uppercase"' : ""}></label>`).join("") +
      (novo ? `<label>Supervisor<input id="cv-supervisor" list="dl-pessoas"></label><label>Coordenador<input id="cv-coordenador" list="dl-pessoas"></label>` : "") +
      `</div>` + (novo ? "" : '<p class="info" style="margin:6px 0 0">O condutor é trocado em "Condutor", logo abaixo, para ficar na linha do tempo.</p>');
  }
  function lerCamposVeiculo(novo) {
    const r = {};
    for (const f of CAMPOS) {
      if (!novo && (f.c === "placa" || f.c === "condutor")) continue;
      const el = $("#cv-" + f.c); if (!el) continue;
      const t = el.value.trim();
      r[f.c] = !t ? null : f.t === "num" ? numero(t) : f.t === "int" ? (parseInt(t, 10) || null) : t;
    }
    if (!r.operacao || !r.gestor_operacao || !r.status) return "Preencha Operação, Gestor operação e Status.";
    if (novo) {
      r.placa = String(r.placa || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
      if (!/^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$/.test(r.placa)) return "Placa inválida (ex.: ABC1D23 ou ABC1234).";
      if (VEICULOS.some(v => v.placa === r.placa)) return "Essa placa já está cadastrada.";
      r.supervisor = $("#cv-supervisor").value.trim() || null; r.coordenador = $("#cv-coordenador").value.trim() || null;
    }
    return r;
  }
  $("#bt-novo-veic").addEventListener("click", () => {
    abrirModal(`<h2>Novo veículo</h2>${camposVeiculo(null, true)}
      <div class="linha-bt"><button class="bt bt-pri" id="cv-salva">Cadastrar</button><span id="cv-msg" class="msg"></span></div>`);
    $("#cv-salva").addEventListener("click", async () => {
      const r = lerCamposVeiculo(true); if (typeof r === "string") return msg($("#cv-msg"), r, "erro");
      msg($("#cv-msg"), "Gravando…");
      try { await api("novo_veiculo", { dados: r }); abrirVeiculo(r.placa); }
      catch (err) { msg($("#cv-msg"), err.message, "erro"); }
    });
  });

  // ------------------------------------------------------------------ histórico de uso (linha do tempo de todos os carros)
  let HIST = [];   // (preenchido por aplicar)
  const fh = { busca: "", de: "", ate: "" };
  // quem estava com o carro (usa o histórico já carregado — os mesmos dados que o banco libera para o login)
  function quemLocal(placa, quando) {
    const t = quando.getTime();
    const h = HIST.filter(x => x.placa === placa && new Date(x.inicio).getTime() <= t && (!x.fim || new Date(x.fim).getTime() > t))
      .sort((a, b) => String(b.inicio).localeCompare(String(a.inicio)))[0];
    return h ? h.condutor : "";
  }
  $("#h-busca").addEventListener("input", e => { fh.busca = sem(e.target.value); desenharHistorico(); });
  $("#h-de").addEventListener("change", e => { fh.de = e.target.value; desenharHistorico(); });
  $("#h-ate").addEventListener("change", e => { fh.ate = e.target.value; desenharHistorico(); });
  function desenharHistorico() {
    const de = fh.de ? new Date(fh.de + "T00:00:00").getTime() : -Infinity, ate = fh.ate ? new Date(fh.ate + "T23:59:59").getTime() : Infinity;
    const l = HIST.filter(h => (!fh.busca || sem(h.placa + " " + h.condutor).includes(fh.busca)) &&
      new Date(h.inicio).getTime() <= ate && (!h.fim || new Date(h.fim).getTime() >= de));
    $("#h-info").textContent = `${l.length} registros`;
    $("#tb-historico").innerHTML = l.length ? l.map(h => {
      const dias = Math.max(0, Math.round(((h.fim ? new Date(h.fim) : new Date()) - new Date(h.inicio)) / 86400000));
      return `<tr class="clica" data-placa="${esc(h.placa)}"><td><b>${esc(h.placa)}</b></td><td>${esc(h.condutor)}</td><td>${dataHora(h.inicio)}</td>
        <td>${h.fim ? dataHora(h.fim) : '<span style="color:#15803d;font-weight:800">com o carro agora</span>'}</td><td class="num">${dias}</td><td>${esc(h.registrado_por) || "—"}</td></tr>`;
    }).join("") : '<tr><td colspan="6" style="text-align:center;color:#475569;padding:24px">Nenhuma troca registrada.</td></tr>';
  }
  $("#tb-historico").addEventListener("click", e => { const tr = e.target.closest("tr[data-placa]"); if (tr) abrirVeiculo(tr.dataset.placa); });
  $("#bt-h-quem").addEventListener("click", () => {
    const placa = $("#h-q-placa").value, q = $("#h-q-data").value; if (!placa || !q) return;
    const c = quemLocal(placa, new Date(q));
    $("#h-q-resp").innerHTML = c ? `👤 <b>${esc(c)}</b> estava com o ${esc(placa)} em ${dataHora(q)}` : `Nenhum condutor registrado para o ${esc(placa)} nessa data.`;
  });
  $("#bt-trocar").addEventListener("click", () => {
    abrirModal(`<h2>🔄 Trocar condutor</h2>
      <div class="campos">
        <label>Placa<select id="tc-placa">${VEICULOS.map(v => `<option value="${esc(v.placa)}">${esc(v.placa)} — ${esc(v.condutor || "sem condutor")}</option>`).join("")}</select></label>
        <label>Novo condutor<input id="tc-nome" list="dl-pessoas" placeholder="Nome do motorista"></label>
        <label>Está com o carro desde<input id="tc-desde" type="datetime-local" value="${paraInput(new Date())}"></label>
        <label class="larga">🔧 Ou mandar para a oficina (nome da oficina, opcional)<input id="tc-oficina" placeholder="Ex.: Auto Center Silva"></label>
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="tc-salva">Registrar troca</button><button class="bt bt-sec" id="tc-ofi">🔧 Mandar para a oficina</button><span id="tc-msg" class="msg"></span></div>
      <p class="nota">Quando o carro voltar, é só registrar a troca para o motorista de novo. Os dias parados ficam no histórico.</p>`);
    $("#tc-ofi").addEventListener("click", () => { const o = $("#tc-oficina").value.trim(); $("#tc-nome").value = "OFICINA" + (o ? " - " + o : ""); $("#tc-salva").click(); });
    $("#tc-salva").addEventListener("click", async () => {
      const nome = $("#tc-nome").value.trim(); if (!nome) return msg($("#tc-msg"), "Digite o nome do novo condutor.", "erro");
      msg($("#tc-msg"), "Gravando…");
      try { await api("trocar_condutor", { placa: $("#tc-placa").value, condutor: nome, desde: doCampo($("#tc-desde").value).toISOString() }); fecharModal(); }
      catch (err) { msg($("#tc-msg"), err.message, "erro"); }
    });
  });

  // ------------------------------------------------------------------ multas e batidas: cruza a base com o histórico
  let CZ = { cab: [], linhas: [], iPlaca: -1, quando: [], chaves: [], sit: "" };
  let TRAT_M = new Map();   // tratativa das multas gravada na planilha (chave → situação, datas, observação)
  const SIT_COR = { "Não tratada": "#b91c1c", "Em tratativa": "#b45309", "Tratada": "#15803d" };
  const dataCurta = s => { const m = String(s || "").match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : ""; };
  const situacaoDe = i => { const t = TRAT_M.get(CZ.chaves[i]); return (t && t.situacao) || "Não tratada"; };
  function celTrat(i) {
    const t = TRAT_M.get(CZ.chaves[i]) || {}, s = situacaoDe(i);
    const datas = [t.data_envio ? "enviada " + dataCurta(t.data_envio) : "", t.data_tratativa ? "tratada " + dataCurta(t.data_tratativa) : ""].filter(Boolean).join(" · ");
    return `<span class="st" style="background:${SIT_COR[s]}">${esc(s)}</span>${datas ? `<small class="resp">${datas}</small>` : ""}` +
      (t.observacao ? `<small class="resp" title="${esc(t.observacao)}">📝 ${esc(String(t.observacao).slice(0, 40))}${String(t.observacao).length > 40 ? "…" : ""}</small>` : "") +
      ` <button class="bt bt-sec bt-mini" data-edita="${i}">✎ Editar</button>`;
  }
  // repinta a coluna de tratativa e o quadro "Tratadas" (depois de salvar ou de recarregar os dados)
  function pintarTrat() {
    if (!CZ.linhas.length) return;
    $$("#tb-cruzar tr").forEach((tr, i) => { const td = tr.querySelector("td.cz-trat"); if (td) td.innerHTML = celTrat(i); });
    const n = { "Não tratada": 0, "Em tratativa": 0, "Tratada": 0 };
    CZ.linhas.forEach((_, i) => { n[situacaoDe(i)]++; });
    const k = $("#cz-trat-kpi");
    if (k) k.innerHTML = `<small>Tratadas</small><b>${n["Tratada"]} de ${CZ.linhas.length}</b><span class="resp">${n["Em tratativa"]} em tratativa · ${n["Não tratada"]} não tratadas</span>`;
    filtrarMultas();
  }
  function editarMulta(i) {
    const l = CZ.linhas[i], t = TRAT_M.get(CZ.chaves[i]) || {};
    const placa = String(l[CZ.iPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, "");
    const hoje = new Date().toISOString().slice(0, 10);
    abrirModal(`<h2>Multa · ${esc(placa)}</h2>
      <div class="dados">
        <div><small>Data da multa</small>${CZ.quando[i] ? CZ.quando[i].toLocaleString("pt-BR") : "—"}</div>
        <div><small>Condutor</small>${esc(CZ.cond[i]) || "—"}</div>
        <div><small>Valor</small>${CZ.valores[i] ? "R$ " + fmtN(CZ.valores[i], 2) : "—"}</div>
        <div><small>Faixa</small>${esc(CZ.faixas[i]) || "—"}</div>
        ${CZ.iAuto >= 0 ? `<div><small>${esc(CZ.cab[CZ.iAuto])}</small>${esc(l[CZ.iAuto])}</div>` : ""}
      </div>
      <h3>Tratativa</h3>
      <div class="campos">
        <label>Situação<select id="mt-sit">${["Não tratada", "Em tratativa", "Tratada"].map(s => `<option ${situacaoDe(i) === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>
        <label>Data de envio ao condutor<input id="mt-envio" type="date" value="${esc(String(t.data_envio || "").slice(0, 10))}"></label>
        <label>Data da tratativa<input id="mt-data" type="date" value="${esc(String(t.data_tratativa || "").slice(0, 10))}"></label>
        <label class="larga">Observação<input id="mt-obs" value="${esc(t.observacao || "")}" placeholder="Ex.: condutor indicado no Detran, desconto em folha, recurso…"></label>
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="mt-salva">Salvar</button><span id="mt-msg" class="msg"></span>
        ${t.tratado_por ? `<span class="info">última alteração: ${esc(t.tratado_por)} · ${dataHora(t.tratado_em)}</span>` : ""}</div>`);
    $("#mt-sit").addEventListener("change", e => { if (e.target.value === "Tratada" && !$("#mt-data").value) $("#mt-data").value = hoje; });
    $("#mt-salva").addEventListener("click", async () => {
      const sit = $("#mt-sit").value, dt = $("#mt-data").value;
      if (sit === "Tratada" && !dt) return msg($("#mt-msg"), "Informe a data da tratativa.", "erro");
      msg($("#mt-msg"), "Gravando…");
      const q = CZ.quando[i];
      try {
        await api("salvar_multa", { dados: { chave: CZ.chaves[i], placa, data_multa: q ? paraInput(q).replace("T", " ") : String(l[CZ.iData] || ""),
          auto: CZ.iAuto >= 0 ? String(l[CZ.iAuto] || "") : "", valor: CZ.valores[i] ? CZ.valores[i].toFixed(2) : "", condutor: CZ.cond[i] || "", faixa: CZ.faixas[i] || "",
          situacao: sit, data_envio: $("#mt-envio").value, data_tratativa: dt, observacao: $("#mt-obs").value.trim() } });
        fecharModal();
      } catch (err) { msg($("#mt-msg"), err.message, "erro"); }
    });
  }
  const dataDe = (v, hora) => {
    let d = null;
    if (v instanceof Date) d = new Date(v.getTime());
    else if (typeof v === "number" && v > 20000 && v < 80000) d = new Date(Math.round((v - 25569) * 86400000) + new Date().getTimezoneOffset() * 60000);
    else {
      const m = String(v || "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[ T]+(\d{1,2}):(\d{2}))?/);
      if (m) d = new Date(Number(m[3].length === 2 ? "20" + m[3] : m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0));
      else { const iso = String(v || "").match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/); if (iso) d = new Date(+iso[1], +iso[2] - 1, +iso[3], +(iso[4] || 0), +(iso[5] || 0)); }
    }
    if (d && hora != null && hora !== "") {
      if (typeof hora === "number" && hora < 1) { const min = Math.round(hora * 1440); d.setHours(Math.floor(min / 60), min % 60, 0, 0); }
      else { const h = String(hora).match(/(\d{1,2}):(\d{2})/); if (h) d.setHours(+h[1], +h[2], 0, 0); }
    }
    return d && !isNaN(d) ? d : null;
  };
  $("#arq-cruzar").addEventListener("change", async e => {
    const arq = e.target.files[0]; e.target.value = ""; if (!arq) return;
    const el = $("#cz-msg");
    try {
      const wb = XLSX.read(await arq.arrayBuffer(), { type: "array", cellDates: true });
      let m = null, r = -1;
      for (const nome of wb.SheetNames) {
        const mm = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: true });
        const rr = mm.slice(0, 40).findIndex(l => l.some(c => sem(c) === "placa") && l.some(c => sem(c).startsWith("data")));
        if (rr >= 0) { m = mm; r = rr; break; }
      }
      if (!m) return msg(el, "Não achei uma aba com as colunas Placa e Data.", "erro");
      const cab = m[r].map(c => String(c).trim());
      const iP = cab.map(sem).indexOf("placa");
      // datas viram texto dd/mm/aaaa hh:mm (assim a base fica igual quando volta da planilha)
      const p2 = n => String(n).padStart(2, "0");
      const txt = v => !(v instanceof Date) || isNaN(v) ? v : v.getFullYear() < 1901 ? p2(v.getHours()) + ":" + p2(v.getMinutes())
        : `${p2(v.getDate())}/${p2(v.getMonth() + 1)}/${v.getFullYear()}` + (v.getHours() || v.getMinutes() ? ` ${p2(v.getHours())}:${p2(v.getMinutes())}` : "");
      const linhas = m.slice(r + 1).filter(l => String(l[iP] || "").trim()).map(l => cab.map((_, j) => txt(l[j] === undefined ? "" : l[j])));
      processarMultas(cab, linhas);
      // grava na planilha (Frota.xlsx, aba Multas) para a base continuar aqui depois de atualizar a página
      msg(el, `Gravando ${linhas.length} multas na planilha…`);
      const regs = CZ.linhas.map((l, i) => {
        const o = {}; cab.forEach((c, j) => { if (c) o[c] = l[j]; });
        const q = CZ.quando[i];
        return { chave: CZ.chaves[i], placa: String(l[CZ.iPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, ""), data_multa: q ? paraInput(q).replace("T", " ") : String(l[CZ.iData] || ""),
          auto: CZ.iAuto >= 0 ? String(l[CZ.iAuto] || "") : "", valor: CZ.valores[i] ? CZ.valores[i].toFixed(2) : "", condutor: CZ.cond[i] || "", faixa: CZ.faixas[i] || "", dados: JSON.stringify(o) };
      });
      const r2 = await api("subir_multas", { regs });
      const rr2 = (r2 && r2.r) || {};
      msg(el, `✓ Base gravada: ${rr2.novas != null ? rr2.novas + " multas novas, " + rr2.atualizadas + " já estavam na base" : linhas.length + " multas"}. Ela continua aqui quando você atualizar a página.`, "ok");
    } catch (err) { msg(el, "Erro ao ler a base: " + (err.message || err), "erro"); }
  });
  // monta a tela de multas a partir de cabeçalho + linhas (vem do arquivo ou da planilha gravada)
  function processarMultas(cab, linhas) {
      const el = $("#cz-msg");
      cab = cab.map(c => String(c).trim()); const cs = cab.map(sem);
      const iPlaca = cs.indexOf("placa");
      const iData = cs.findIndex(c => c === "data" || c.startsWith("data da") || c.startsWith("data do") || c.startsWith("data/hora") || c.startsWith("data e hora") || c.startsWith("data"));
      const iHora = cs.findIndex(c => c === "hora" || c.startsWith("hora da") || c.startsWith("hora do"));
      // valor da multa: coluna com "valor" (de preferência "valor da multa"), "vlr" ou "R$"
      let iValor = cs.findIndex(c => /valor|vlr/.test(c) && /multa|infra/.test(c));
      if (iValor < 0) iValor = cs.findIndex(c => /^valor|^vlr|r\$/.test(c) && !/desconto|pago|juros/.test(c));
      let achou = 0, semData = 0;
      const quando = [], cond = [], origem = [];
      linhas.forEach(l => {
        const placa = String(l[iPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, "");
        const d = dataDe(l[iData], iHora >= 0 ? l[iHora] : null);
        quando.push(d);
        // quem estava no dia (linha do tempo); se não tiver registro nessa data, usa o condutor da base de veículos
        let c = d ? quemLocal(placa, d) : "", orig = c ? "dia" : "";
        if (!c) { const v = TODOS.veic.find(x => x.placa === placa); if (v && v.condutor) { c = v.condutor; orig = "base"; } }
        if (!d) semData++;
        if (c) achou++;
        cond.push(c); origem.push(orig);
      });
      // condutor atual (base de veículos) e reincidência: placa com 2+ multas na base, ou que já tinha multa registrada
      const veic = new Map(TODOS.veic.map(v => [v.placa, v]));
      const placaDe = l => String(l[iPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, "");
      const naBase = {}; linhas.forEach(l => { const p = placaDe(l); naBase[p] = (naBase[p] || 0) + 1; });
      const antes = {}; TODOS.ocor.filter(o => o.tipo === "Multa").forEach(o => { antes[o.placa] = (antes[o.placa] || 0) + 1; });
      const condBase = linhas.map(l => (veic.get(placaDe(l)) || {}).condutor || "");
      const valores = linhas.map(l => iValor >= 0 ? n0(numBR(l[iValor])) : 0);
      // faixa (gravidade): coluna Gravidade/Natureza; senão pelo valor tabelado do CTB
      const iGrav = cs.findIndex(c => /gravidade|natureza|classifica/.test(c));
      const faixaDoValor = v => {
        if (!v) return "Sem valor";
        const perto = (a, b) => Math.abs(a - b) < 0.6;
        if (perto(v, 88.38)) return "Leve"; if (perto(v, 130.16)) return "Média"; if (perto(v, 195.23)) return "Grave";
        const f = v / 293.47; if (Math.abs(f - Math.round(f)) < 0.01 && Math.round(f) >= 1) return "Gravíssima" + (Math.round(f) > 1 ? ` (${Math.round(f)}x)` : "");
        return "Outro valor";
      };
      const faixaDoTexto = t => { const x = sem(t); return /gravissima/.test(x) ? "Gravíssima" : /grave/.test(x) ? "Grave" : /media/.test(x) ? "Média" : /leve/.test(x) ? "Leve" : ""; };
      const faixas = linhas.map((l, i) => (iGrav >= 0 && faixaDoTexto(l[iGrav])) || faixaDoValor(valores[i]));
      const alertas = linhas.map(l => { const p = placaDe(l), n = naBase[p], a = antes[p] || 0;
        return n >= 2 ? `⚠ Reincidente (${n + a}x)` : a ? `⚠ Reincidente (${a + 1}x)` : ""; });
      // identificação da multa: nº do auto de infração (AIT); se não tiver, placa + data/hora + valor
      const iAuto = cs.findIndex(c => /(^|[^a-z])(ait|auto)([^a-z]|$)/.test(c) || /auto de infra|n(umero|º|o)? ?(da|de)? ?infra/.test(c));
      const chavesM = linhas.map((l, i) => { const au = iAuto >= 0 ? String(l[iAuto] || "").trim().toUpperCase() : "";
        return au ? "AIT:" + au : placaDe(l) + "|" + (quando[i] ? paraInput(quando[i]) : String(l[iData] || "").trim()) + "|" + valores[i].toFixed(2); });
      CZ = { cab, linhas, iPlaca, iData, iAuto, quando, cond, origem, condBase, valores, alertas, iValor, faixas, chaves: chavesM, so: false, faixa: "", sit: "" };
      const reinc = Object.keys(naBase).filter(p => naBase[p] >= 2 || antes[p]);
      const tot = valores.reduce((a, v) => a + v, 0);
      const kp = (t, v, c) => `<div class="kpi" style="border-left-color:${c}"><small>${t}</small><b>${v}</b></div>`;
      $("#cz-kpis").innerHTML = kp("Valor das multas", iValor >= 0 ? "R$ " + fmtN(tot, 2) : "sem coluna de valor", "#b91c1c") + kp("Multas na base", linhas.length.toLocaleString("pt-BR"), "#0e7c86")
        + kp("Placas reincidentes", reinc.length, reinc.length ? "#b91c1c" : "#15803d").replace('class="kpi"', 'class="kpi" id="cz-reinc" title="Clique para ver só as reincidentes"') + kp("Com condutor", `${achou} de ${linhas.length}`, "#4338ca")
        + `<div class="kpi" id="cz-trat-kpi" style="border-left-color:#15803d"></div>`;
      $("#cz-reinc").addEventListener("click", () => { CZ.so = !CZ.so; $("#cz-reinc").classList.toggle("on", CZ.so); filtrarMultas(); });
      // quadro por faixa: quantidade e valor; clique filtra a tabela
      const ORDEM = ["Leve", "Média", "Grave", "Gravíssima"], COR = { "Leve": "#16a34a", "Média": "#ca8a04", "Grave": "#ea580c", "Gravíssima": "#b91c1c" };
      const grupo = f => f.startsWith("Gravíssima") ? "Gravíssima" : f;
      const porF = {};
      faixas.forEach((f, i) => { const g = grupo(f); porF[g] = porF[g] || { n: 0, v: 0 }; porF[g].n++; porF[g].v += valores[i]; });
      const chaves = ORDEM.filter(k => porF[k]).concat(Object.keys(porF).filter(k => ORDEM.indexOf(k) < 0));
      $("#cz-faixas").innerHTML = chaves.map(k => `<div class="faixa" data-faixa="${esc(k)}" style="border-top-color:${COR[k] || "#64748b"}"><small>${esc(k)}</small><b>${porF[k].n}</b><span>R$ ${fmtN(porF[k].v, 2)}</span><i style="width:${Math.round(porF[k].n / linhas.length * 100)}%;background:${COR[k] || "#64748b"}"></i></div>`).join("");
      $$("#cz-faixas .faixa").forEach(el => el.addEventListener("click", () => { CZ.faixa = CZ.faixa === el.dataset.faixa ? "" : el.dataset.faixa; $$("#cz-faixas .faixa").forEach(x => x.classList.toggle("on", x.dataset.faixa === CZ.faixa)); filtrarMultas(); }));
      // ordem: Placa, Condutor e Alerta primeiro; depois as outras colunas da planilha
      const outras = cab.map((_, j) => j).filter(j => j !== iPlaca);
      $("#cz-cab").innerHTML = `<tr><th>${esc(cab[iPlaca])}</th><th style="background:#dcfce7">Condutor</th><th style="background:#fee2e2">Alerta</th><th>Faixa</th><th style="background:#e0f2fe">Tratativa</th>` + outras.map(j => `<th>${esc(cab[j])}</th>`).join("") + "</tr>";
      const fmt = v => v instanceof Date ? v.toLocaleString("pt-BR") : esc(v);
      $("#tb-cruzar").innerHTML = linhas.map((l, i) => `<tr><td><b>${fmt(l[iPlaca])}</b></td>` +
        `<td style="font-weight:800;background:#f0fdf4">${cond[i] ? esc(cond[i]) + (origem[i] === "base" ? '<small class="resp">base de veículos</small>' : '<small class="resp">no dia da multa</small>') : '<span style="color:#b91c1c">placa fora da base</span>'}</td>` +
        `<td style="color:#b91c1c;font-weight:800;white-space:nowrap">${esc(alertas[i])}</td><td><span class="tag-faixa f-${sem(faixas[i]).replace(/[^a-z]/g, "").slice(0, 10)}">${esc(faixas[i])}</span></td><td class="cz-trat" style="white-space:nowrap"></td>` +
        outras.map(j => `<td>${fmt(l[j])}</td>`).join("") + "</tr>").join("");
      $("#cz-busca").value = ""; $("#cz-filtro").hidden = false;
      $$("#tb-cruzar tr").forEach((tr, i) => { if (alertas[i]) tr.classList.add("tem-alerta"); });
      pintarTrat();
      $("#bt-cruzar-exp").hidden = false;   // multas: só baixar (batidas vão em Ocorrências)
  }
  // depois de atualizar a página: remonta a tela com a base de multas gravada na planilha
  function montarMultasSalvas() {
    const salvas = Array.from(TRAT_M.values()).filter(m => m.dados);
    if (!salvas.length) return false;
    const objs = salvas.map(m => { try { return JSON.parse(m.dados); } catch (e) { return null; } }).filter(Boolean);
    const cab = [];
    objs.forEach(o => Object.keys(o).forEach(k => { if (cab.indexOf(k) < 0) cab.push(k); }));
    if (!cab.some(c => sem(c) === "placa")) return false;
    processarMultas(cab, objs.map(o => cab.map(c => o[c] == null ? "" : o[c])));
    $("#cz-msg").textContent = `${objs.length} multas gravadas na planilha. Para incluir novas, clique em Subir base.`;
    $("#cz-msg").className = "msg";
    return true;
  }
  // busca por placa ou condutor (junto com o filtro de reincidentes)
  function filtrarMultas() {
    const b = sem($("#cz-busca").value).replace(/[^a-z0-9 ]/g, "");
    let n = 0;
    $$("#tb-cruzar tr").forEach((tr, i) => {
      const placa = String(CZ.linhas[i] ? CZ.linhas[i][CZ.iPlaca] : "").toUpperCase().replace(/[^A-Z0-9]/g, "");
      const ok = (!CZ.sit || situacaoDe(i) === CZ.sit) && (!CZ.so || CZ.alertas[i]) && (!CZ.faixa || String(CZ.faixas[i] || "").startsWith(CZ.faixa)) && (!b || sem(placa).includes(b.replace(/ /g, "")) || sem(CZ.cond[i] || "").includes(b));
      tr.hidden = !ok; if (ok) n++;
    });
    $("#cz-info").textContent = CZ.linhas.length ? `${n} de ${CZ.linhas.length}` : "";
  }
  $("#cz-busca").addEventListener("input", filtrarMultas);
  $("#cz-sit").addEventListener("change", e => { CZ.sit = e.target.value; filtrarMultas(); });
  $("#tb-cruzar").addEventListener("click", e => { const b = e.target.closest("[data-edita]"); if (b) editarMulta(Number(b.dataset.edita)); });
  $("#bt-cruzar-exp").addEventListener("click", () => {
    const aoa = [CZ.cab.concat("Condutor", "De onde veio o condutor", "Alerta", "Faixa", "Situação", "Data de envio ao condutor", "Data da tratativa", "Observação", "Alterado por")].concat(CZ.linhas.map((l, i) => { const t = TRAT_M.get(CZ.chaves[i]) || {};
      return CZ.cab.map((_, j) => l[j]).concat(CZ.cond[i] || "", CZ.origem[i] === "base" ? "base de veículos" : CZ.origem[i] === "dia" ? "no dia da multa" : "placa fora da base", CZ.alertas[i] || "", CZ.faixas[i] || "",
        situacaoDe(i), dataCurta(t.data_envio), dataCurta(t.data_tratativa), t.observacao || "", t.tratado_por || ""); }));
    const wb = XLSX.utils.book_new(); const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: "dd/mm/yyyy hh:mm" });
    XLSX.utils.book_append_sheet(wb, ws, "Com condutor"); XLSX.writeFile(wb, "Base com condutor no dia.xlsx");
  });
  $("#bt-cruzar-ocor").addEventListener("click", async () => {
    const tipo = $("#cz-tipo").value;
    const placas = new Set(VEICULOS.map(v => v.placa));
    const regs = CZ.linhas.map((l, i) => ({ l, i, placa: String(l[CZ.iPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, "") }))
      .filter(x => CZ.quando[x.i] && placas.has(x.placa))
      .map(x => ({ placa: x.placa, data_hora: CZ.quando[x.i].toISOString(), tipo, condutor: CZ.cond[x.i] || null,
        descricao: CZ.cab.map((c, j) => j === CZ.iPlaca || j === CZ.iData || x.l[j] === "" ? "" : c + ": " + (x.l[j] instanceof Date ? x.l[j].toLocaleString("pt-BR") : x.l[j])).filter(Boolean).join(" · ") || tipo }));
    if (!regs.length) return msg($("#cz-msg"), "Nenhuma linha com placa da sua equipe e data válida.", "erro");
    if (!window.confirm(`Registrar ${regs.length} ocorrências do tipo "${tipo}"?`)) return;
    msg($("#cz-msg"), "Gravando…");
    try {
      for (let i = 0; i < regs.length; i += 100) await api("nova_ocorrencia", { ocorrencias: regs.slice(i, i + 100) });
      msg($("#cz-msg"), `✓ ${regs.length} ocorrências registradas.`, "ok");
    } catch (err) { msg($("#cz-msg"), "Erro: " + err.message, "erro"); }
  });

  // ------------------------------------------------------------------ consumo de combustível
  // relatório da Ticket Log: uma linha por placa e combustível no mês. Cada placa fica ligada ao condutor
  // do mês (linha do tempo); carros coringa ficam numa tabela à parte.
  let COMB = [], EXTRATO = [], CARTOES = [];
  const fcb = { mes: "", gestor: "", contrato: "", comb: "", cc: "", busca: "", alerta: false, difnome: false };
  const numBR = v => {
    if (typeof v === "number") return v;
    let t = String(v == null ? "" : v).replace(/R\$|\s/g, "");
    if (!t || t === "-" || t === "—") return null;
    if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
    else if (/^-?\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, "");
    const n = Number(t); return isFinite(n) ? n : null;
  };
  const n0 = v => { const n = Number(v); return isFinite(n) ? n : 0; };
  const fmtN = (v, c = 0) => v == null || !isFinite(v) ? "-" : v.toLocaleString("pt-BR", { minimumFractionDigits: c, maximumFractionDigits: c });
  const nomeMes = m => { if (!m) return ""; const [a, b] = m.split("-"); return ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"][Number(b) - 1] + "/" + a; };
  // coringa: cartão/carro sem condutor fixo (modelo "CARTAO CORINGA", placa COR0001…, ou condutor/tipo de uso/status "coringa")
  const ehCoringa = (r, v) => /^COR\d{4}$/.test(String(r.placa || "")) || /coringa/.test(sem(r.condutores) + " " + sem(r.modelo) + " " + sem(r.tipo_uso) + " " + (v ? sem(v.condutor) + " " + sem(v.modelo) + " " + sem(v.tipo_uso) + " " + sem(v.status) : ""));
  (() => { const d = new Date(); $("#cb-mes-up").value = d.toISOString().slice(0, 7); })();

  // pode escolher os dois arquivos de uma vez (relatório de consumo + Extrato Simplificado): cada um tem a sua mensagem
  $("#arq-comb").addEventListener("change", async e => {
    const arqs = Array.from(e.target.files || []); e.target.value = "";
    if (!arqs.length) return;
    msg($("#cb-msg"), "");
    const caixa = $("#cb-msgs"); caixa.innerHTML = "";
    for (const arq of arqs) {
      const el = document.createElement("p"); el.className = "msg"; caixa.appendChild(el);
      await subirArquivoComb(arq, el);
    }
  });
  async function subirArquivoComb(arq, el) {
    try {
      msg(el, "Lendo " + arq.name + "…");
      const buf = await arq.arrayBuffer();
      let wb;
      const inicio = new TextDecoder("windows-1252").decode(buf.slice(0, 4000));
      if (/<\s*(table|tr|td|html|main)\b/i.test(inicio)) {
        // o ".xls" da Ticket Log é uma página HTML: lê a tabela respeitando as células mescladas
        let txt; try { txt = new TextDecoder("utf-8", { fatal: true }).decode(buf); } catch (e2) { txt = new TextDecoder("windows-1252").decode(buf); }
        wb = { SheetNames: ["html"], Sheets: { html: XLSX.utils.aoa_to_sheet(htmlParaMatriz(txt)) } };
      } else if (/\.(csv|txt)$/i.test(arq.name)) {
        let txt; try { txt = new TextDecoder("utf-8", { fatal: true }).decode(buf); } catch (e2) { txt = new TextDecoder("windows-1252").decode(buf); }
        const sep = (txt.split("\n").slice(0, 10).join("").match(/;/g) || []).length > (txt.split("\n").slice(0, 10).join("").match(/\t/g) || []).length ? ";" : (/\t/.test(txt) ? "\t" : ";");
        const linhasTxt = txt.replace(/^\uFEFF/, "").split(/\r?\n/).map(l => l.split(sep).map(c => c.replace(/^"|"$/g, "").trim()));
        wb = { SheetNames: ["csv"], Sheets: { csv: XLSX.utils.aoa_to_sheet(linhasTxt) } };
      } else wb = XLSX.read(buf, { type: "array", cellDates: false, raw: true });
      // Extrato Simplificado (limite do contrato e saldo de cada cartão)? → outra leitura
      for (const nome of wb.SheetNames) {
        const t = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: true });
        if (/saldo contratual|extrato simplificado/.test(sem(t.slice(0, 60).map(l => l.join(" ")).join(" ")))) return await subirExtrato(t, el);
        if (t.slice(0, 40).some(l => l.some(c => sem(c).replace(/\s+/g, " ").trim() === "codigo transacao"))) return await subirTransacoes(t, el);
      }
      let m = null, r0 = -1;
      for (const nome of wb.SheetNames) {
        const t = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: true });
        r0 = t.slice(0, 40).findIndex(l => l.some(c => sem(c) === "placa"));
        if (r0 >= 0) { m = t; break; }
      }
      if (!m) return msg(el, "Não achei a coluna Placa na planilha.", "erro");
      // período escrito no topo do relatório (ex.: 01/10/2026 a 31/10/2026) → sugere o mês
      const topo = m.slice(0, r0).map(l => l.join(" ")).join(" ");
      const dt = topo.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      const datas = topo.match(/\d{2}\/\d{2}\/\d{4}/g) || [];
      const periodo = datas.length >= 2 ? `de ${datas[0].slice(0, 5)} até ${datas[1].slice(0, 5)}` : "";
      let mes = $("#cb-mes-up").value;
      if (dt && `${dt[3]}-${dt[2]}` !== mes && window.confirm(`A planilha fala de ${dt[1]}/${dt[2]}/${dt[3]}. Gravar como ${nomeMes(dt[3] + "-" + dt[2])}?`)) { mes = `${dt[3]}-${dt[2]}`; $("#cb-mes-up").value = mes; }
      if (!mes) return msg(el, "Escolha o mês da planilha.", "erro");
      // títulos sem acento e sem símbolos (aceita "Combust�vel" de arquivo com acento quebrado)
      const limpa = h => sem(h).replace(/[^a-z0-9/ ]/g, "").replace(/\s+/g, " ").trim();
      const cab = m[r0].map(limpa), cab2 = (m[r0 + 1] || []).map(limpa);
      const col = teste => cab.findIndex(teste);
      const ix = {
        placa: col(h => h === "placa"), modelo: col(h => h === "modelo"), marca: col(h => h === "marca"), ano: col(h => h === "ano"),
        combustivel: col(h => h.includes("combust")), ultima_km: col(h => /ltima km/.test(h)), km_rodados: col(h => h.startsWith("km rodado")),
        litros: col(h => h === "litros"), valor_litro: col(h => /^valor m/.test(h)), km_litro: col(h => /^km p/.test(h)),
        total: col(h => h.startsWith("total")),
      };
      const iCartao = col(h => /^cart/.test(h));
      ix.limite = cab2.indexOf("limite") >= 0 ? cab2.indexOf("limite") : iCartao;
      ix.saldo = cab2.indexOf("saldo") >= 0 ? cab2.indexOf("saldo") : (iCartao >= 0 ? iCartao + 1 : -1);
      const faltam = ["litros", "total"].filter(k => ix[k] < 0);
      if (faltam.length) return msg(el, "Não achei as colunas: " + faltam.join(", "), "erro");
      const regs = [], pendentes = [], vistos = new Set();
      let totalCC = 0, temCC = false, totalComb = 0, repetidas = 0;
      m.slice(r0 + 1).forEach(l => {
        const c0 = String(l[0] == null ? "" : l[0]);
        const cc = sem(c0).match(/^centro de custo\s*:?\s*(.*)$/);
        if (/^combust/.test(sem(c0))) { totalComb += n0(numBR(l[ix.total])); return; }
        if (cc) { temCC = true; totalCC += n0(numBR(l[ix.total])); const nomeCc = c0.split(":").slice(1).join(":").replace(/\s+/g, " ").trim(); pendentes.forEach(r => { r.centro_custo = nomeCc; }); pendentes.length = 0; return; }
        const placa = String(l[ix.placa] || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
        if (!/^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$/.test(placa)) return;
        const txt = k => ix[k] >= 0 ? String(l[ix[k]] == null ? "" : l[ix[k]]).trim() : "";
        const num = k => ix[k] >= 0 ? numBR(l[ix[k]]) : null;
        const r = { placa, modelo: txt("modelo"), marca: txt("marca"), ano: txt("ano"), combustivel: txt("combustivel"),
          limite: num("limite"), saldo: num("saldo"), ultima_km: num("ultima_km"), km_rodados: num("km_rodados"), litros: num("litros"),
          valor_litro: num("valor_litro"), km_litro: num("km_litro"), total: num("total"), centro_custo: "" };
        Object.keys(r).forEach(k => { if (r[k] === null) r[k] = ""; });
        // a mesma linha repetida no relatório (mesma placa, combustível e valores) entra uma vez só
        const chave = [placa, r.combustivel, r.litros, r.total, r.km_rodados].join("|");
        if (vistos.has(chave)) { repetidas++; return; }
        vistos.add(chave);
        regs.push(r); pendentes.push(r);
      });
      if (!regs.length) return msg(el, "Nenhuma placa encontrada na planilha.", "erro");
      if (COMB.some(x => x.mes === mes) && !window.confirm(`${nomeMes(mes)} já tem consumo gravado. Substituir pelos dados desta planilha?`)) return msg(el, "");
      msg(el, `Enviando ${regs.length} linhas de ${nomeMes(mes)}…`);
      await api("subir_combustivel", { mes, periodo, regs });
      fcb.mes = mes; montarFiltrosComb(); desenharComb(); desenharInicioComb();
      const fora = COMB.filter(x => x.mes === mes && x.na_frota === "NAO").length;
      // confere com o total que o próprio relatório traz (linhas CENTRO DE CUSTO / Combustível)
      const lido = regs.reduce((a, r) => a + n0(r.total), 0);
      const doRel = temCC ? totalCC : totalComb;
      const confere = doRel ? (Math.abs(doRel - lido) < 1 ? ` Total confere com o relatório: R$ ${fmtN(lido, 2)}.` : ` ⚠ Total lido R$ ${fmtN(lido, 2)} × total do relatório R$ ${fmtN(doRel, 2)} (me mande o arquivo para eu conferir).`) : ` Total lido: R$ ${fmtN(lido, 2)}.`;
      msg(el, `✓ ${arq.name}: ${regs.length} linhas de ${nomeMes(mes)} gravadas.` + confere + (repetidas ? ` ${repetidas} linha(s) repetida(s) ignorada(s).` : "") + (fora ? ` ${fora} placa(s) não estão cadastradas na frota.` : ""), doRel && Math.abs(doRel - lido) >= 1 ? "erro" : "ok");
    } catch (err) { msg(el, arq.name + ": erro ao subir: " + (err.message || err), "erro"); }
  }

  function montarFiltrosComb() {
    const meses = Array.from(new Set(COMB.map(x => x.mes).concat(TRANS_MESES))).filter(Boolean).sort().reverse();
    if (!fcb.mes || meses.indexOf(fcb.mes) < 0) fcb.mes = meses[0] || "";
    $("#cb-mes").innerHTML = meses.length ? meses.map(x => `<option value="${x}">${nomeMes(x)}</option>`).join("") : "<option value=''>Nenhum mês</option>";
    $("#cb-mes").value = fcb.mes;
    const op = (id, lista, todos) => { const v = $(id).value; $(id).innerHTML = `<option value="">${todos}</option>` + lista.map(x => `<option>${esc(x)}</option>`).join(""); $(id).value = lista.indexOf(v) >= 0 ? v : ""; };
    const doMes = fcb.mes ? fonteComb().rows : [];
    op("#cb-gestor", gestoresDo($("#cb-contrato").value), "Todos os gestores");
    op("#cb-contrato", contratosDo($("#cb-gestor").value), "Todos os contratos");
    fcb.contrato = $("#cb-contrato").value;
    op("#cb-comb", Array.from(new Set(doMes.map(x => x.combustivel).filter(Boolean))).sort(), "Todos os combustíveis");
    atualizarListaCC();
    fcb.gestor = $("#cb-gestor").value; fcb.comb = $("#cb-comb").value; fcb.cc = ccDoCampo($("#cb-cc-txt").value);
  }
  $("#cb-mes").addEventListener("change", e => { fcb.mes = e.target.value; montarFiltrosComb(); desenharComb(); if (!TRANS_C[fcb.mes]) garantirTrans(fcb.mes, () => { montarFiltrosComb(); desenharComb(); }); });
  $("#cb-gestor").addEventListener("change", e => { fcb.gestor = e.target.value; montarFiltrosComb(); desenharComb(); });
  $("#cb-contrato").addEventListener("change", e => { fcb.contrato = e.target.value; montarFiltrosComb(); desenharComb(); });
  $("#cb-comb").addEventListener("change", e => { fcb.comb = e.target.value; desenharComb(); });
  $("#cb-cc-txt").addEventListener("input", e => { fcb.cc = ccDoCampo(e.target.value); desenharComb(); });
  $("#tb-cc").addEventListener("click", e => { const tr = e.target.closest("tr[data-cc]"); if (!tr) return; fcb.cc = fcb.cc === tr.dataset.cc ? "" : tr.dataset.cc; $("#cb-cc-txt").value = fcb.cc; desenharComb(); });
  $("#cb-busca").addEventListener("input", e => { fcb.busca = sem(e.target.value); desenharComb(); });
  $("#cb-alerta").addEventListener("change", e => { fcb.alerta = e.target.checked; desenharComb(); });
  $("#cb-difnome").addEventListener("change", e => { fcb.difnome = e.target.checked; desenharComb(); });
  $("#cb-escopo").checked = ESCOPO;
  $("#cb-escopo").parentElement.title = "Contratos: " + CONTRATOS_FROTA.join(", ");
  $("#cb-escopo").addEventListener("change", e => {
    ESCOPO = e.target.checked; try { localStorage.setItem("frotaEscopo_v1", ESCOPO ? "1" : "0"); } catch (x) { }
    filtrarEquipe(); redesenharTudo();
  });
  // responsável do cartão (Ticket Log) diferente do condutor da base de veículos
  const nomeDiferente = p => { const q = (p.motoristas && p.motoristas[0]) || (p.cartao && p.cartao.responsavel); return !!(q && p.condutor && sem(q) !== sem(p.condutor)); };

  // junta as linhas de cada placa (gasolina + etanol) e marca alerta de consumo
  // de onde vêm os números do mês: as TRANSAÇÕES (mês inteiro, igual ao portal) quando já subiu; senão o relatório de consumo
  function fonteComb() {
    const tr = TRANS_C[fcb.mes];
    if (tr && tr.length) {
      const ccPl = new Map(CARTOES.map(c => [String(c.placa || "").toUpperCase(), c.centro_custo]));
      const ccCo = new Map(COMB.filter(x => x.centro_custo).map(x => [x.placa, x.centro_custo]));
      const vis = new Set(VEICULOS.map(v => v.placa));
      const rows = tr.filter(t => noEscopo(t.placa) && (!equipeAtiva() || vis.has(t.placa))).map(t => {
        const ok = !ehPedagio(t) && kmValido(t);
        return { placa: String(t.placa || "").toUpperCase(), mes: fcb.mes, combustivel: ehPedagio(t) ? "" : String(t.combustivel || "").trim().toUpperCase(),
          km_rodados: ok ? n0(numBR(t.km_rodados)) : 0, litros: ehPedagio(t) ? 0 : n0(numBR(t.litros)), litrosKm: ok ? n0(numBR(t.litros)) : 0, total: n0(numBR(t.valor)),
          centro_custo: ccPl.get(String(t.placa || "").toUpperCase()) || ccCo.get(t.placa) || "", modelo: t.modelo || "", ultima_km: n0(numBR(t.hodometro)), condutores: "", data: String(t.data || "") };
      });
      const ult = rows.reduce((m, t) => t.data > m ? t.data : m, "");
      return { trans: true, rows, periodo: `de 01/${fcb.mes.slice(5, 7)} até ${ult ? ult.slice(8, 10) + "/" + ult.slice(5, 7) : "hoje"} · transações` };
    }
    const per = (COMB.find(x => x.mes === fcb.mes && x.periodo) || {}).periodo;
    return { trans: false, rows: COMB.filter(x => x.mes === fcb.mes), periodo: per ? per + " · relatório de consumo" : "" };
  }
  function consumoPorPlaca() {
    const veic = new Map(VEICULOS.map(v => [v.placa, v]));
    const mp = new Map();
    fonteComb().rows.filter(x => bateDono(x.placa, x, fcb.gestor, fcb.contrato) && (!fcb.comb || x.combustivel === fcb.comb) && bateCC(x.centro_custo, fcb.cc)).forEach(x => {
      let p = mp.get(x.placa);
      if (!p) { const v = veic.get(x.placa); const dn = dono(x.placa, x); p = { placa: x.placa, condutor: x.condutores || dn.condutor || "", modelo: x.modelo || (v && v.modelo) || "", gestor: dn.gestor, operacao: dn.contrato,
        cc: x.centro_custo || "", combs: [], km: 0, litros: 0, total: 0, limite: x.limite, saldo: x.saldo, coringa: ehCoringa(x, v), naFrota: x.na_frota !== "NAO" }; mp.set(x.placa, p); }
      if (x.combustivel && p.combs.indexOf(x.combustivel) < 0) p.combs.push(x.combustivel);
      p.km += n0(x.km_rodados); p.litros += n0(x.litros); p.total += n0(x.total); p.litrosKm = (p.litrosKm || 0) + n0(x.litrosKm != null ? x.litrosKm : x.litros); p.odo = Math.max(p.odo || 0, n0(x.ultima_km));
      if (!p.cc && x.centro_custo) p.cc = x.centro_custo;
    });
    const lista = Array.from(mp.values());
    lista.forEach(p => { p.kml = p.litrosKm > 0 && p.km > 0 ? p.km / p.litrosKm : null; p.rsl = p.litros > 0 ? p.total / p.litros : null; });
    // referência: mediana do km/L do mesmo modelo (precisa de 3 carros ou mais)
    const porModelo = {};
    lista.forEach(p => { if (p.kml) (porModelo[sem(p.modelo)] = porModelo[sem(p.modelo)] || []).push(p.kml); });
    lista.forEach(p => {
      const ks = (porModelo[sem(p.modelo)] || []).slice().sort((a, b) => a - b);
      p.ref = ks.length >= 3 ? ks[Math.floor(ks.length / 2)] : null;
      p.alerta = p.litros > 0 && p.km <= 0 ? "Abasteceu sem km válido" : p.ref && p.kml && p.kml < p.ref * 0.75 ? `Km/L abaixo do normal do ${p.modelo} (${fmtN(p.ref, 1)})` : "";
    });
    return lista.filter(p => (!fcb.busca || sem([p.placa, p.condutor, p.modelo, p.gestor, p.cc].join(" ")).includes(fcb.busca)) && (!fcb.alerta || p.alerta));
  }
  const CAB_CB = c => `<tr><th>${c}</th><th>Placa</th><th>Modelo</th><th>Combustível</th><th>Gestor operação</th><th class="num">Km rodados</th><th class="num">Litros</th><th class="num">R$/litro</th><th class="num">Km/L</th><th class="num">Total R$</th><th class="num">Saldo cartão</th><th>Centro de custo</th><th>Alerta</th></tr>`;
  function linhasComb(lista, vazio) {
    if (!lista.length) return `<tr><td colspan="13" class="info">${vazio}</td></tr>`;
    const s = lista.reduce((a, p) => ({ km: a.km + p.km, l: a.l + p.litros, t: a.t + p.total }), { km: 0, l: 0, t: 0 });
    return lista.map(p => `<tr class="clica${p.alerta ? " tem-alerta" : ""}" data-placa="${esc(p.placa)}"><td class="cond">${esc(p.condutor || (p.coringa ? "Cartão coringa" : "Sem condutor na linha do tempo"))}</td><td><b>${esc(p.placa)}</b>${p.naFrota ? "" : ' <span class="tag">fora da frota</span>'}</td><td>${esc(p.modelo)}</td><td>${esc(p.combs.join(" + "))}</td><td>${esc(p.gestor)}</td>
      <td class="num">${fmtN(p.km)}</td><td class="num">${fmtN(p.litros, 2)}</td><td class="num">${fmtN(p.rsl, 2)}</td><td class="num${p.alerta ? " alerta-km" : ""}">${fmtN(p.kml, 2)}</td><td class="num">${fmtN(p.total, 2)}</td><td class="num">${p.saldo === "" || p.saldo == null ? "-" : fmtN(n0(p.saldo), 2)}</td><td>${esc(p.cc)}</td><td class="info">${esc(p.alerta)}</td></tr>`).join("")
      + `<tr><td colspan="5"><b>Total (${lista.length} ${lista.length === 1 ? "carro" : "carros"})</b></td><td class="num"><b>${fmtN(s.km)}</b></td><td class="num"><b>${fmtN(s.l, 2)}</b></td><td class="num"><b>${fmtN(s.l ? s.t / s.l : null, 2)}</b></td><td class="num"><b>${fmtN(s.l ? s.km / s.l : null, 2)}</b></td><td class="num"><b>${fmtN(s.t, 2)}</b></td><td colspan="3"></td></tr>`;
  }
  // junta, por placa, o consumo do mês (relatório) e o cartão (Extrato Simplificado)
  function linhasUnificadas() {
    const lista = COMB.length ? consumoPorPlaca() : [];
    const porPlaca = new Map(lista.map(p => [p.placa, Object.assign({ temConsumo: true }, p)]));
    const veic = new Map(VEICULOS.map(v => [v.placa, v]));
    CARTOES.filter(c => bateCC(c.centro_custo, fcb.cc) && bateDono(c.placa, null, fcb.gestor, fcb.contrato)).forEach(c => {
      const placa = String(c.placa || "").toUpperCase();
      let p = porPlaca.get(placa);
      if (!p) {
        if (fcb.comb || fcb.alerta) return;   // filtro de combustível/alerta: só quem tem consumo
        const d = dono(placa);
        if (fcb.busca && !sem([placa, c.responsavel, c.centro_custo, c.modelo, d.condutor].join(" ")).includes(fcb.busca)) return;
        p = { placa, condutor: d.condutor, modelo: c.modelo, gestor: d.gestor, operacao: d.contrato, cc: "", combs: [], km: 0, litros: 0, total: 0, kml: null, rsl: null, alerta: "",
          coringa: ehCoringa({ placa, modelo: c.modelo }, veic.get(placa)), naFrota: !!d.v, temConsumo: false };
        porPlaca.set(placa, p);
      }
      p.cartao = c; p.limiteC = n0(c.limite_atual); p.utilC = n0(c.utilizado); p.saldoC = n0(c.saldo);
      p.pct = p.limiteC > 0 ? p.utilC / p.limiteC * 100 : 0;
      if (!p.modelo) p.modelo = c.modelo;
    });
    const motM = motoristaDoMes(fcb.mes);
    porPlaca.forEach(p => { const ms = motM.get(p.placa); if (ms && ms.length) p.motoristas = ms; });
    const todas = Array.from(porPlaca.values()).filter(p => !fcb.difnome || nomeDiferente(p));
    const nomeP = p => (p.motoristas && p.motoristas[0]) || (p.cartao && p.cartao.responsavel) || p.condutor || "";
    const cond = todas.filter(p => !p.coringa).sort((a, b) => (nomeP(a) ? 0 : 1) - (nomeP(b) ? 0 : 1) || sem(nomeP(a)).localeCompare(sem(nomeP(b))) || a.placa.localeCompare(b.placa));
    const cor = todas.filter(p => p.coringa).sort((a, b) => a.placa.localeCompare(b.placa));
    return { lista, cond, cor };
  }
  const CAB_UNI = () => `<tr><th title="Responsável do cartão na Ticket Log; embaixo, o condutor da base de veículos quando for outra pessoa">Quem abastece</th><th>Placa</th><th>Modelo · combustível</th><th>Gestor · contrato</th><th>Centro de custo</th>
    <th class="num" title="Km rodados no mês (embaixo: odômetro)">Km</th><th class="num">Litros</th><th class="num">R$/L</th><th class="num">Km/L</th><th class="num">Consumo</th>
    <th class="num" title="Utilizado no cartão (embaixo: limite)">Cartão</th><th class="num" title="Saldo do cartão">Saldo</th><th class="num">% uso</th></tr>`;
  const NCOL = 13;
  function linhasUni(l, vazio) {
    if (!l.length) return `<tr><td colspan="${NCOL}" class="info">${vazio}</td></tr>`;
    const dif = (a, b) => a && b && sem(a) !== sem(b);
    const s = l.reduce((a, p) => ({ km: a.km + p.km, l: a.l + p.litros, t: a.t + p.total, li: a.li + (p.limiteC || 0), u: a.u + (p.utilC || 0), sa: a.sa + (p.saldoC || 0) }), { km: 0, l: 0, t: 0, li: 0, u: 0, sa: 0 });
    return l.map(p => {
      const c = p.cartao || {};
      // nome principal = quem abastece (responsável do cartão na Ticket Log); embaixo, o condutor da base se for outro
      // quem abastece: motorista das transações do mês; senão o responsável do cartão; embaixo, o condutor da base se for outro
      const quem = (p.motoristas && p.motoristas[0]) || c.responsavel || "";
      const outros = p.motoristas && p.motoristas.length > 1 ? `<small class="resp">também: ${esc(p.motoristas.slice(1, 3).join(", "))}${p.motoristas.length > 3 ? "…" : ""}</small>` : "";
      const resp = outros + (quem && p.condutor && dif(p.condutor, quem) ? `<small class="resp">condutor na base: ${esc(p.condutor)}</small>` : "");
      const nome = quem || p.condutor || (p.coringa ? "Cartão coringa" : "Sem condutor na base");
      const ccC = c.centro_custo || "", ccR = p.cc || "";
      const ccTxt = dif(ccR, ccC) ? `${esc(ccR)}<small class="resp">no extrato: ${esc(ccC)}</small>` : esc(ccR || ccC);
      const pct = p.cartao ? p.pct : null;
      return `<tr class="clica${p.alerta || (pct != null && pct >= 90) ? " tem-alerta" : ""}" data-placa="${esc(p.placa)}"${c.cartao ? ` title="Cartão ${esc(c.cartao)}"` : ""}>
        <td class="cond">${esc(nome)}${resp}${p.naFrota === false ? '<small class="resp">fora da base de veículos</small>' : ""}${p.alerta ? `<small class="resp alerta-txt">⚠ ${esc(p.alerta)}</small>` : ""}</td>
        <td><b>${esc(p.placa)}</b></td><td>${esc(p.modelo)}${p.combs.length ? `<small class="resp">${esc(p.combs.join(" + "))}</small>` : ""}</td>
        <td>${esc(p.gestor)}${p.operacao ? `<small class="resp">${esc(p.operacao)}</small>` : ""}</td><td>${ccTxt}</td>
        <td class="num">${p.temConsumo ? fmtN(p.km) : "—"}${p.odo ? `<small class="resp">odôm. ${fmtN(p.odo)}</small>` : ""}</td><td class="num">${p.temConsumo ? fmtN(p.litros, 2) : "—"}</td><td class="num">${fmtN(p.rsl, 2)}</td>
        <td class="num${p.alerta ? " alerta-km" : ""}">${fmtN(p.kml, 2)}</td><td class="num">${p.temConsumo ? fmtN(p.total, 2) : "—"}</td>
        <td class="num">${p.cartao ? `${fmtN(p.utilC, 2)}<small class="resp">de ${fmtN(p.limiteC, 2)}</small>` : "—"}</td><td class="num"><b>${p.cartao ? fmtN(p.saldoC, 2) : "—"}</b></td>
        <td class="num">${pct == null ? "—" : `<span class="uso"><i style="width:${Math.min(100, pct)}%;background:${pct >= 90 ? "#b91c1c" : pct >= 70 ? "#ea580c" : "#0e7c86"}"></i></span> ${fmtN(pct, 0)}%`}</td></tr>`;
    }).join("") + `<tr class="linha-total"><td colspan="5"><b>Total (${l.length} ${l.length === 1 ? "carro" : "carros"})</b></td><td class="num"><b>${fmtN(s.km)}</b></td><td class="num"><b>${fmtN(s.l, 2)}</b></td>
      <td class="num"><b>${fmtN(s.l ? s.t / s.l : null, 2)}</b></td><td class="num"><b>${fmtN(s.l ? s.km / s.l : null, 2)}</b></td><td class="num"><b>${fmtN(s.t, 2)}</b></td>
      <td class="num"><b>${fmtN(s.u, 2)}</b><small class="resp">de ${fmtN(s.li, 2)}</small></td><td class="num"><b>${fmtN(s.sa, 2)}</b></td><td class="num"><b>${s.li ? fmtN(s.u / s.li * 100, 0) + "%" : "—"}</b></td></tr>`;
  }
  function desenharComb() {
    $("#cb-cab").innerHTML = CAB_UNI();
    desenharTrans();
    const FT = fonteComb();
    if (!FT.rows.length && !CARTOES.length) { $("#cb-kpis").innerHTML = ""; $("#tb-comb").innerHTML = `<tr><td colspan="${NCOL}" class="info">Nenhum consumo gravado ainda.${EU && EU.perfil === "admin" ? " Suba o relatório de consumo e o Extrato Simplificado." : ""}</td></tr>`; $("#cb-info").textContent = ""; $("#tb-cc").innerHTML = ""; return; }
    const { lista, cond, cor } = linhasUnificadas();
    const tot = lista.reduce((a, p) => ({ km: a.km + p.km, l: a.l + p.litros, lk: a.lk + (p.litrosKm || 0), t: a.t + p.total }), { km: 0, l: 0, lk: 0, t: 0 });
    const k = (t, v, cor2) => `<div class="kpi"${cor2 ? ` style="border-left-color:${cor2}"` : ""}><small>${t}</small><b>${v}</b></div>`;
    $("#cb-kpis").innerHTML = k("Gasto no mês", "R$ " + fmtN(tot.t, 2)) + k("Litros", fmtN(tot.l)) + k("Km rodados", fmtN(tot.km)) + k("Km/L médio", fmtN(tot.lk ? tot.km / tot.lk : null, 2))
      + k("Carros com condutor", cond.length) + k("Nome diferente da base", `${cond.filter(nomeDiferente).length} de ${cond.filter(p => p.cartao && p.condutor).length}`, "#ea580c") + k("Carros coringa", cor.length, "#7c3aed") + k("Com alerta", cond.concat(cor).filter(p => p.alerta || (p.cartao && p.pct >= 90)).length, "#b91c1c");
    // uma tabela só: condutores e, no fim, a seção dos coringas
    $("#tb-comb").innerHTML = linhasUni(cond, "Nenhum carro com esse filtro.")
      + (cor.length ? `<tr class="secao"><td colspan="${NCOL}">Coringa · ${cor.length} ${cor.length === 1 ? "cartão" : "cartões"}</td></tr>` + linhasUni(cor, "") : "");
    $("#cb-uni-info").textContent = `${cond.length} carros` + (CARTOES.length ? ` · extrato de ${CARTOES[0].data_ref || "-"}` : "");
    $("#cb-cor-info").textContent = cor.length ? `· coringa: ${cor.length} · consumo R$ ${fmtN(cor.reduce((a, p) => a + p.total, 0), 2)} · saldo R$ ${fmtN(cor.reduce((a, p) => a + (p.saldoC || 0), 0), 2)}` : "";
    if (!FT.rows.length) { $("#cb-info").textContent = ""; $("#tb-cc").innerHTML = '<tr><td colspan="7" class="info">Suba as transações (ou o relatório de consumo) para ver o gasto por centro de custo.</td></tr>'; return; }
    $("#cb-info").textContent = `${nomeMes(fcb.mes)}${FT.periodo ? " (" + FT.periodo + ")" : ""} · ${lista.length} carros`;
    // gasto por centro de custo (respeita mês, gestor e combustível)
    const g = {};
    FT.rows.filter(x => bateDono(x.placa, x, fcb.gestor, fcb.contrato) && (!fcb.comb || x.combustivel === fcb.comb)).forEach(x => {
      const c = x.centro_custo || "Sem centro de custo";
      const o = g[c] = g[c] || { placas: new Set(), l: 0, lk: 0, km: 0, t: 0 };
      o.placas.add(x.placa); o.l += n0(x.litros); o.lk += n0(x.litrosKm != null ? x.litrosKm : x.litros); o.km += n0(x.km_rodados); o.t += n0(x.total);
    });
    const ccs = Object.keys(g).sort((a, b) => g[b].t - g[a].t), totG = ccs.reduce((a, c) => a + g[c].t, 0);
    $("#tb-cc").innerHTML = ccs.map(c => `<tr class="clica${fcb.cc && bateCC(c, fcb.cc) ? " cc-on" : ""}" data-cc="${esc(c === "Sem centro de custo" ? "" : c)}"><td><b>${esc(c)}</b></td><td class="num">${g[c].placas.size}</td><td class="num">${fmtN(g[c].l, 2)}</td><td class="num">${fmtN(g[c].km)}</td><td class="num">${fmtN(g[c].lk ? g[c].km / g[c].lk : null, 2)}</td><td class="num"><b>${fmtN(g[c].t, 2)}</b></td><td class="num">${totG ? fmtN(g[c].t / totG * 100, 1) + "%" : "-"}</td></tr>`).join("") || '<tr><td colspan="7" class="info">Sem dados.</td></tr>';
  }
  ["#tb-comb"].forEach(id => $(id).addEventListener("click", e => { const tr = e.target.closest("tr[data-placa]"); if (tr && VEICULOS.some(v => v.placa === tr.dataset.placa)) abrirVeiculo(tr.dataset.placa); }));

  $("#bt-comb-exp").addEventListener("click", () => {
    const { cond, cor } = linhasUnificadas();
    const lista = cond.concat(cor);
    if (!lista.length) return msg($("#cb-msg"), "Nada para baixar com esse filtro.", "erro");
    const r2 = v => v == null ? null : Math.round(v * 100) / 100;
    const linha = p => { const c = p.cartao || {}; return { "Mês": nomeMes(fcb.mes), "Quem abastece (Ticket Log)": c.responsavel || "", "Condutor na base de veículos": p.condutor, "Placa": p.placa, "Modelo": p.modelo,
      "Combustível": p.combs.join(" + "), "Gestor operação": p.gestor, "Contrato": p.operacao, "Centro de custo (consumo)": p.cc, "Centro de custo (extrato)": c.centro_custo || "",
      "Odômetro (última km)": p.odo || null, "Km rodados": p.temConsumo ? p.km : null, "Litros": p.temConsumo ? r2(p.litros) : null, "R$/litro": r2(p.rsl), "Km/L": r2(p.kml), "Consumo R$": p.temConsumo ? r2(p.total) : null,
      "Cartão": c.cartao || "", "Limite cartão": p.cartao ? p.limiteC : null, "Utilizado cartão": p.cartao ? p.utilC : null, "Saldo cartão": p.cartao ? p.saldoC : null,
      "% usado": p.cartao ? r2(p.pct) : null, "Alerta": p.alerta }; };
    const wb = XLSX.utils.book_new();
    [["Por condutor", cond], ["Coringa", cor]].forEach(([nome, l]) => {
      const ws = XLSX.utils.json_to_sheet(l.length ? l.map(linha) : [{ "Placa": "" }]);
      ws["!cols"] = Object.keys(linha(lista[0])).map(k2 => ({ wch: Math.max(12, k2.length + 4) }));
      XLSX.utils.book_append_sheet(wb, ws, nome);
    });
    const det = COMB.filter(x => x.mes === fcb.mes).map(x => ({ "Mês": nomeMes(x.mes), "Placa": x.placa, "Condutores no mês": x.condutores, "Modelo": x.modelo, "Combustível": x.combustivel,
      "Km rodados": numBR(x.km_rodados), "Litros": numBR(x.litros), "R$/litro": numBR(x.valor_litro), "Km/L": numBR(x.km_litro), "Total R$": numBR(x.total), "Limite": numBR(x.limite), "Saldo": numBR(x.saldo), "Centro de custo": x.centro_custo }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(det), "Detalhe por combustível");
    XLSX.writeFile(wb, `Combustivel_${fcb.mes || "cartoes"}.xlsx`);
  });

  // liga a placa à base de veículos: condutor, gestor operação e contrato (operação) vêm do cadastro do carro
  let MAPA_VEIC = new Map(), MAPA_VEIC_N = -1;
  function dono(placa, x) {
    if (MAPA_VEIC_N !== VEICULOS.length || !MAPA_VEIC.size) { MAPA_VEIC = new Map(VEICULOS.map(v => [v.placa, v])); MAPA_VEIC_N = VEICULOS.length; }
    const v = MAPA_VEIC.get(String(placa || "").toUpperCase()) || null;
    return { v, condutor: v ? (v.condutor || "") : "", gestor: (v && v.gestor_operacao) || (x && x.gestor_operacao) || "", contrato: (v && v.operacao) || (x && x.operacao) || "" };
  }
  // NOVO (25/11/2026): transação de placa fora da base → acha o gestor pelo NOME DO MOTORISTA
  // (condutor de algum carro da base; se não, a hierarquia: colaborador → gerente)
  let MOT_MAPA = null, MOT_N = -1, MOT_H = null;
  const chavesMot = s => { const p = sem(s).replace(/[^a-z ]/g, " ").split(/\s+/).filter(x => x && ["de", "da", "do", "das", "dos", "e"].indexOf(x) < 0); return p.length ? [p.join(" "), p.length > 1 ? p[0] + " " + p[p.length - 1] : ""].filter(Boolean) : []; };
  function donoPorMotorista(nome) {
    if (!nome) return null;
    if (MOT_N !== VEICULOS.length || MOT_H !== TODOS.hier) {
      MOT_MAPA = new Map(); MOT_N = VEICULOS.length; MOT_H = TODOS.hier;
      const pon = (k, d) => { if (!MOT_MAPA.has(k)) MOT_MAPA.set(k, d); else if (MOT_MAPA.get(k) && MOT_MAPA.get(k).gestor !== d.gestor) MOT_MAPA.set(k, null); };
      VEICULOS.filter(v => v.condutor && v.gestor_operacao).forEach(v => chavesMot(v.condutor).forEach(k => pon(k, { gestor: v.gestor_operacao, contrato: v.operacao || "", via: "motorista" })));
      (TODOS.hier || []).filter(h => h.colaborador && h.gerente).forEach(h => chavesMot(h.colaborador).forEach(k => { if (!MOT_MAPA.has(k)) MOT_MAPA.set(k, { gestor: h.gerente, contrato: "", via: "hierarquia" }); }));
    }
    const ks = chavesMot(nome);
    return (ks[0] && MOT_MAPA.get(ks[0])) || (ks[1] && MOT_MAPA.get(ks[1])) || null;
  }
  function donoTrans(t) {
    const d = dono(t.placa);
    if (d.gestor) return d;
    const m = donoPorMotorista(t.motorista) || donoPorMotorista(t.responsavel);
    return m ? { v: null, condutor: "", gestor: m.gestor, contrato: d.contrato || m.contrato, via: m.via } : d;
  }
  const bateDono = (placa, x, g, ct) => { if (!g && !ct) return true; const d = dono(placa, x); return (!g || d.gestor === g) && (!ct || d.contrato === ct); };
  const opSel = (id, lista, todos) => { const el = $(id), v = el.value; el.innerHTML = `<option value="">${todos}</option>` + lista.map(x => `<option>${esc(x)}</option>`).join(""); el.value = lista.indexOf(v) >= 0 ? v : ""; return el.value; };

  // centro de custo: valor exato, ou "grupo:CLARO" = todos os que começam com CLARO
  const primeiraPalavra = c => (sem(c).match(/^[a-z0-9]+/) || [""])[0];
  // "texto:claro sp" = centros que têm todas as palavras (cada palavra pode ser o começo: "sp" pega SP e SPC)
  const bateCC = (valor, filtro) => {
    if (!filtro) return true;
    if (filtro.startsWith("grupo:")) return primeiraPalavra(valor) === filtro.slice(6);
    if (filtro.startsWith("texto:")) {
      const ps = sem(valor).split(/[^a-z0-9]+/).filter(Boolean);
      return sem(filtro.slice(6)).split(/\s+/).filter(Boolean).every(w => ps.some(p => p.startsWith(w)));
    }
    return valor === filtro;
  };
  const nomeCC = f => f.startsWith("grupo:") ? f.slice(6).toUpperCase() + " (todos)" : f.startsWith("texto:") ? "“" + f.slice(6) + "”" : f;
  // um campo só: nome exato da lista → aquele centro; qualquer outro texto → soma os que têm as palavras
  let CCS_TODOS = [];
  function ccDoCampo(v) {
    const t = String(v || "").trim();
    if (!t) return "";
    const exato = CCS_TODOS.find(c => sem(c) === sem(t));
    return exato || "texto:" + t;
  }
  function atualizarListaCC() {
    CCS_TODOS = Array.from(new Set(COMB.map(x => x.centro_custo).concat(CARTOES.map(x => x.centro_custo)).filter(Boolean))).sort();
    $("#dl-cc").innerHTML = CCS_TODOS.map(c => `<option value="${esc(c)}">`).join("");
  }
  function opcoesCC(ccs) {
    const cont = {};
    ccs.forEach(c => { const w = primeiraPalavra(c); if (w.length >= 3 && !/^\d+$/.test(w)) cont[w] = (cont[w] || 0) + 1; });
    const grupos = Object.keys(cont).filter(w => cont[w] >= 2).sort();
    return '<option value="">Todos os centros de custo</option>'
      + (grupos.length ? '<optgroup label="Grupos">' + grupos.map(w => `<option value="grupo:${w}">${esc(w.toUpperCase())} (todos · ${cont[w]} centros)</option>`).join("") + "</optgroup>" : "")
      + '<optgroup label="Centros de custo">' + ccs.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join("") + "</optgroup>";
  }
  // tabela HTML → linhas e colunas (colspan/rowspan viram células vazias, como a mesclagem do Excel)
  function htmlParaMatriz(txt) {
    const doc = new DOMParser().parseFromString(txt, "text/html");
    const out = [];
    doc.querySelectorAll("table").forEach(tb => {
      const ocupado = [];   // ocupado[linha][coluna] por rowspan de cima
      Array.from(tb.rows).forEach((tr, ri) => {
        const linha = []; ocupado[ri] = ocupado[ri] || [];
        let c = 0;
        Array.from(tr.cells).forEach(td => {
          while (ocupado[ri][c]) { linha[c] = ""; c++; }
          const v = (td.textContent || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
          const cs = Math.max(1, Number(td.getAttribute("colspan")) || 1), rs = Math.max(1, Number(td.getAttribute("rowspan")) || 1);
          for (let k = 0; k < cs; k++) {
            linha[c + k] = k === 0 ? v : "";
            for (let r = 1; r < rs; r++) { ocupado[ri + r] = ocupado[ri + r] || []; ocupado[ri + r][c + k] = true; }
          }
          c += cs;
        });
        for (let k = 0; k < linha.length; k++) if (linha[k] === undefined) linha[k] = "";
        out.push(linha);
      });
    });
    return out;
  }

  // ------------------------------------------------------------------ transações (cada abastecimento: data, motorista, litros)
  const TRANS_C = {};
  let TRANS_MESES = [];            // mês → lista de transações (carrega quando precisa)
  const transPedidas = {};
  function garantirTrans(mes, depois) {
    if (!mes || TRANS_C[mes]) return;
    if (transPedidas[mes]) { if (depois) transPedidas[mes].push(depois); return; }
    transPedidas[mes] = depois ? [depois] : [];
    api("transacoes", { mes }).then(r => {
      TRANS_C[mes] = (r.transacoes || []).filter(t => t.placa);
      if (r.meses) TRANS_MESES = r.meses;
      const fs = transPedidas[mes] || []; delete transPedidas[mes];
      fs.forEach(f => { try { f(); } catch (e) { } });
    }).catch(() => { delete transPedidas[mes]; });
  }
  const TANQUE = [["mobi", 47], ["kwid", 38], ["onix", 44], ["prisma", 54], ["hb20", 50], ["argo", 48], ["cronos", 48], ["gol", 55], ["voyage", 55], ["polo", 52],
    ["virtus", 52], ["saveiro", 55], ["strada", 55], ["montana", 44], ["fiorino", 50], ["kangoo", 60], ["partner", 60], ["doblo", 60], ["ranger", 80], ["hilux", 80],
    ["s10", 76], ["l200", 75], ["amarok", 80], ["tracker", 44], ["kicks", 41], ["yaris", 45], ["logan", 50], ["sandero", 50], ["duster", 50], ["spin", 53],
    ["cobalt", 54], ["corolla", 50], ["renegade", 60], ["compass", 60], ["master", 80], ["ducato", 90], ["sprinter", 75], ["daily", 90], [" up", 50]];
  const tanqueDe = modelo => { const m = " " + sem(modelo); const t = TANQUE.find(x => m.includes(x[0])); return t ? t[1] : 60; };
  // dias com mais de 2 abastecimentos na mesma placa, ou abastecimento acima do tanque
  // ---------- regras de alerta das transações (iguais à planilha de análise)
  //  km inválido · consumo < 5 km/l (FROTA/RAC) · 2+ abastecimentos da placa no dia · > 80 L ·
  //  preço/litro > 15% acima da mediana do mesmo combustível no mesmo estado · aditivado. GMG fora de km, dia e litros.
  // NUMERO FROTA: o da transação; se não veio, o do cartão (Extrato Simplificado) pelo nº do cartão ou pela placa
  let NRF_MAPA = null, NRF_BASE = null;
  function nrFrotaDe(t) {
    const prop = String(t.nr_frota || "").trim();
    if (prop) return prop.toUpperCase();
    if (NRF_BASE !== TODOS.cart) {
      NRF_BASE = TODOS.cart; NRF_MAPA = new Map();
      (TODOS.cart || []).forEach(c => { const n = String(c.nr_frota || "").trim().toUpperCase(); if (!n) return;
        const ct = String(c.cartao || "").replace(/\D/g, ""); if (ct) NRF_MAPA.set("c" + ct, n);
        const pl = String(c.placa || "").toUpperCase().replace(/[^A-Z0-9]/g, ""); if (pl && !NRF_MAPA.has("p" + pl)) NRF_MAPA.set("p" + pl, n); });
    }
    const ct = String(t.cartao || "").replace(/\D/g, "");
    return (ct && NRF_MAPA.get("c" + ct)) || NRF_MAPA.get("p" + String(t.placa || "").toUpperCase()) || "";
  }
  const ehGMG = t => /gerador|\bgmg\b/.test(sem([t.modelo, t.familia, t.placa, t.servico, t.tipo_frota, nrFrotaDe(t)].join(" ")));
  const precoLitro = t => { const p = n0(numBR(t.vl_litro)); if (p > 0) return p; const l = n0(numBR(t.litros)); return l > 0 ? n0(numBR(t.valor)) / l : 0; };
  const mediana = l => { const a = l.filter(x => x > 0).sort((x, y) => x - y); return a.length ? a[Math.floor(a.length / 2)] : 0; };
  function analisarTrans(rows) {
    const veic = new Map(TODOS.veic.map(v => [v.placa, v]));
    const ofi = periodosOficina();
    const base = rows.filter(t => !ehPedagio(t));
    // mediana do preço por combustível + UF (usa o mês inteiro, não só o filtro)
    const grupos = {};
    base.forEach(t => { const k = String(t.combustivel || "").toUpperCase() + "|" + String(t.uf || "").toUpperCase(); (grupos[k] = grupos[k] || []).push(precoLitro(t)); });
    const med = {}; Object.keys(grupos).forEach(k => { med[k] = mediana(grupos[k]); });
    const porDia = {};
    base.forEach(t => { if (!ehGMG(t)) { const k = t.placa + "|" + String(t.data).slice(0, 10); porDia[k] = (porDia[k] || 0) + 1; } });
    return base.map(t => {
      const g = ehGMG(t), km = n0(numBR(t.km_rodados)), l = n0(numBR(t.litros)), kml = l > 0 ? km / l : 0;
      const v = veic.get(t.placa) || {};
      const classif = String(t.tipo_frota || v.tipo_frota || "").toUpperCase();
      const f = {};
      f.kmInv = !g && !(km > 0 && km <= 1500 && (l <= 0 || kml <= 25));
      f.consumo = !g && !f.kmInv && l > 0 && kml < 5 && (!classif || /FROTA|RAC/.test(classif));
      f.dia2 = !g && porDia[t.placa + "|" + String(t.data).slice(0, 10)] >= 2;
      f.l80 = !g && l > 80;
      const m = med[String(t.combustivel || "").toUpperCase() + "|" + String(t.uf || "").toUpperCase()];
      f.preco = m > 0 && precoLitro(t) > m * 1.15;
      f.aditivado = /aditiv/.test(sem(t.combustivel));
      const ps = ofi.get(t.placa), x = new Date(String(t.data).replace(" ", "T")).getTime();
      f.oficina = !!(ps && ps.some(p => x >= p[0] && x < p[1]));
      return { t, g, km, l, kml, preco: precoLitro(t), medPreco: m, f, alerta: Object.keys(f).some(k => f[k]) };
    });
  }
  const NOMES_ALERTA = { kmInv: "km inválido", consumo: "consumo < 5 km/l", dia2: "2+ no dia", l80: "> 80 L", preco: "preço alto", aditivado: "aditivado", oficina: "carro na oficina" };
  // por condutor: contagens, pontos e "o que olhar"
  function alertasPorCondutor(an) {
    const veic = new Map(TODOS.veic.map(v => [v.placa, v]));
    const por = new Map();
    an.forEach(a => {
      const nome = (a.t.motorista || a.t.responsavel || "").replace(/\s+/g, " ").trim().toUpperCase();
      if (!nome) return;
      if (!por.has(nome)) por.set(nome, []);
      por.get(nome).push(a);
    });
    const out = [];
    por.forEach((l, nome) => {
      const c = k => l.filter(a => a.f[k]).length;
      // hodômetro digitado: o mesmo km rodado repetido em muitos abastecimentos
      const kms = l.filter(a => !a.g && a.km > 0).map(a => Math.round(a.km));
      const cnt = {}; kms.forEach(k => { cnt[k] = (cnt[k] || 0) + 1; });
      const moda = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
      const hodo = moda && cnt[moda] >= 3 && cnt[moda] / l.length >= 0.4 ? { km: Number(moda), n: cnt[moda] } : null;
      l.forEach(a => { a.hodo = !!(hodo && Math.round(a.km) === hodo.km); });
      const comAlerta = l.filter(a => a.alerta || a.hodo);
      const n = { consumo: c("consumo"), kmInv: c("kmInv"), dia2: c("dia2"), l80: c("l80"), preco: c("preco"), aditivado: c("aditivado"), oficina: c("oficina") };
      const pontos = n.consumo + n.kmInv + n.dia2 + n.l80 * 2 + n.preco + n.aditivado * 0.5 + n.oficina * 2 + (hodo ? hodo.n * 0.5 : 0);
      const kmV = l.filter(a => !a.g && !a.f.kmInv).reduce((s, a) => s + a.km, 0), litV = l.filter(a => !a.g && !a.f.kmInv).reduce((s, a) => s + a.l, 0);
      const olhar = [];
      if (hodo) olhar.push(`Hodômetro digitado: +${fmtN(hodo.km)} km fixo em ${hodo.n} de ${l.length} abast.`);
      if (n.consumo) olhar.push(`consumo abaixo de 5 km/l em ${n.consumo}`);
      if (n.dia2) olhar.push(`${n.dia2} abast. em dias com 2+`);
      if (n.l80) olhar.push(`${n.l80} acima de 80 L`);
      if (n.preco) olhar.push(`preço alto em ${n.preco}`);
      if (n.aditivado) olhar.push(`aditivado em ${n.aditivado}`);
      if (n.kmInv) olhar.push(`km inválido em ${n.kmInv}`);
      if (n.oficina) olhar.push(`${n.oficina} com o carro na oficina`);
      const vs = l.map(a => veic.get(a.t.placa) || {});
      // modelo e combustível vêm das colunas MODELO VEICULO e TIPO COMBUSTIVEL da base de transações
      const modelos = Array.from(new Set(l.map((a, i) => String(a.t.modelo || vs[i].modelo || "").trim().toUpperCase()).filter(Boolean)));
      const litT = l.reduce((s2, a) => s2 + a.l, 0), porC = {};
      l.forEach(a => { const cb = String(a.t.combustivel || "").trim().toUpperCase() || "NÃO INFORMADO"; porC[cb] = (porC[cb] || 0) + a.l; });
      const cbs = Object.keys(porC).sort((a, b) => porC[b] - porC[a]);
      const cap = x => x.charAt(0) + x.slice(1).toLowerCase();
      const combTxt = cbs.map(c2 => cbs.length > 1 && litT ? `${cap(c2)} ${Math.round(porC[c2] / litT * 100)}%` : cap(c2)).join(" · ");
      out.push(Object.assign({ nome, modelo: modelos.join(" / "), combTxt, gestor: maisFrequente(vs.map(v => v.gestor_operacao)), contrato: maisFrequente(vs.map(v => v.operacao)), abast: l.length, comAlerta: comAlerta.length,
        pontos, kml: litV ? kmV / litV : null, gasto: l.reduce((s, a) => s + n0(numBR(a.t.valor)), 0), gastoAlerta: comAlerta.reduce((s, a) => s + n0(numBR(a.t.valor)), 0),
        olhar: olhar.join("; "), hodo, itens: l }, n));
    });
    return out;
  }
  // consumo de combustível / Início: dias por placa com algum alerta
  function alertasTrans(rows) {
    const an = analisarTrans(rows);
    const dias = new Map();
    an.forEach(a => { const k = a.t.placa + "|" + String(a.t.data).slice(0, 10); if (!dias.has(k)) dias.set(k, []); dias.get(k).push(a); });
    const veic = new Map(TODOS.veic.map(v => [v.placa, v]));
    const out = [];
    dias.forEach((l, k) => {
      if (!l.some(a => a.alerta)) return;
      l.sort((a, b) => String(a.t.data).localeCompare(String(b.t.data)));
      const cont = {}; l.forEach(a => Object.keys(a.f).forEach(f => { if (a.f[f]) cont[f] = (cont[f] || 0) + 1; }));
      const tipos = Object.keys(cont).map(f => f === "dia2" ? `${l.length} abastecimentos no dia` : `${NOMES_ALERTA[f]}${cont[f] > 1 ? " (" + cont[f] + ")" : ""}`);
      out.push({ dia: k.split("|")[1], placa: l[0].t.placa, modelo: l[0].t.modelo || (veic.get(l[0].t.placa) || {}).modelo || "", itens: l.map(a => Object.assign({ _a: a }, a.t)),
        n: l.length, dia2: !!cont.dia2, acima: cont.l80 || 0, oficina: cont.oficina || 0, tipos, motoristas: Array.from(new Set(l.map(a => a.t.motorista || a.t.responsavel).filter(Boolean))),
        litros: l.reduce((s, a) => s + a.l, 0), valor: l.reduce((s, a) => s + n0(numBR(a.t.valor)), 0) });
    });
    return out.sort((a, b) => b.dia.localeCompare(a.dia) || a.placa.localeCompare(b.placa));
  }
  // mesmas regras de filtro da tela (equipe, gestor, contrato, centro de custo, busca)
  function transFiltradas(mes, opts) {
    const rows = TRANS_C[mes] || [];
    const veicVis = new Set(VEICULOS.map(v => v.placa));
    const ccCartao = new Map(CARTOES.map(c => [String(c.placa || "").toUpperCase(), c.centro_custo]));
    return rows.filter(t => noEscopo(t.placa) && (!equipeAtiva() || veicVis.has(t.placa))
      && (!opts || ((!opts.g && !opts.ct) || (d => (!opts.g || d.gestor === opts.g) && (!opts.ct || d.contrato === opts.ct))(donoTrans(t))) && (!opts.cc || bateCC(ccCartao.get(t.placa) || "", opts.cc))
        && (!opts.busca || sem([t.placa, t.motorista, t.responsavel, t.estabelecimento, t.cidade].join(" ")).includes(opts.busca))));
  }
  function desenharTrans() {
    const caixa = $("#cb-trans-caixa");
    const hoje = new Date();
    const mes = fcb.mes || (hoje.getFullYear() + "-" + String(hoje.getMonth() + 1).padStart(2, "0"));
    if (!TRANS_C[mes]) { garantirTrans(mes, () => { montarFiltrosComb(); desenharComb(); }); caixa.hidden = true; return; }
    const rows = transFiltradas(mes, { g: fcb.gestor, ct: fcb.contrato, cc: fcb.cc, busca: fcb.busca });
    caixa.hidden = !TRANS_C[mes].length;
    const al = alertasTrans(rows);
    const mais2 = al.filter(a => a.dia2).length, acima = al.reduce((s, a) => s + a.acima, 0);
    $("#cb-trans-info").textContent = `${nomeMes(mes)} · ${rows.length} abastecimentos · ${fmtN(rows.reduce((s, t) => s + n0(numBR(t.litros)), 0))} litros`;
    $("#cb-trans-kpis").innerHTML = `<span class="chip vermelho">🔴 ${mais2} ${mais2 === 1 ? "dia" : "dias"} com 2+ abastecimentos</span><span class="chip laranja">🟠 ${acima} acima de 80 L</span>` + (al.some(a => a.oficina) ? `<span class="chip roxo">🔧 ${al.reduce((s, a) => s + (a.oficina || 0), 0)} com o carro na oficina</span>` : "");
    $("#tb-trans").innerHTML = al.slice(0, 500).map(a => `<tr class="${a.dia2 ? "tem-alerta" : ""}">
      <td>${a.dia.split("-").reverse().join("/")}</td><td><b>${esc(a.placa)}</b><small class="resp">${esc(a.modelo)}</small></td>
      <td class="cond">${esc(a.motoristas.join(" / ") || "—")}</td><td class="num"><b>${a.n}</b></td><td class="num">${fmtN(a.litros, 2)}</td><td class="num">${fmtN(a.valor, 2)}</td>
      <td class="det">${a.itens.map(t => `<span class="${t._a && t._a.alerta ? "acima" : ""}">${esc(String(t.data).slice(11, 16))} · ${fmtN(n0(numBR(t.litros)), 2)} L · ${esc(t.estabelecimento || "")}${t.cidade ? " (" + esc(t.cidade) + ")" : ""}</span>`).join("")}</td>
      <td class="alerta-txt">${a.tipos.map(esc).join("<br>")}</td></tr>`).join("") || '<tr><td colspan="8" class="info">Nenhum alerta nesse período. 👍</td></tr>';
  }
  // ---------- Gasto x Km rodado por técnico (transações do mês, sem abastecimento de gerador)
  const ehGerador = t => /gerador/.test(sem([t.modelo, t.familia, t.placa, t.servico].join(" ")));
  const ehPedagio = t => /pedag/.test(sem(t.servico));
  // km válido (mesma regra dos alertas): maior que 0, até 1.500 km e no máximo 25 km/L
  const kmValido = t => { const km = n0(numBR(t.km_rodados)), l = n0(numBR(t.litros)); return km > 0 && km <= 1500 && l > 0 && km / l <= 25; };
  const maisFrequente = l => { const c = {}; l.filter(Boolean).forEach(x => { c[x] = (c[x] || 0) + 1; }); return Object.keys(c).sort((a, b) => c[b] - c[a])[0] || ""; };
  function rankingTecnicos(rows, comGerador) {
    const veic = new Map(TODOS.veic.map(v => [v.placa, v]));
    const por = new Map();
    rows.forEach(t => {
      const nome = (t.motorista || t.responsavel || "").replace(/\s+/g, " ").trim().toUpperCase();
      if (!nome) return;
      if (!por.has(nome)) por.set(nome, { nome, ts: [], ped: 0 });
      const o = por.get(nome);
      if (ehPedagio(t)) { o.ped += n0(numBR(t.valor)); return; }
      if (ehGerador(t) && !comGerador) return;
      o.ts.push(t);
    });
    const out = [];
    por.forEach(o => {
      if (!o.ts.length) return;
      const vs = o.ts.map(t => veic.get(t.placa) || (d => d.gestor ? { gestor_operacao: d.gestor, operacao: d.contrato } : {})(donoTrans(t)));
      const val = o.ts.filter(kmValido);
      const litros = o.ts.reduce((a, t) => a + n0(numBR(t.litros)), 0), gasto = o.ts.reduce((a, t) => a + n0(numBR(t.valor)), 0);
      const kmV = val.reduce((a, t) => a + n0(numBR(t.km_rodados)), 0), litV = val.reduce((a, t) => a + n0(numBR(t.litros)), 0), gastoV = val.reduce((a, t) => a + n0(numBR(t.valor)), 0);
      // combustíveis que o técnico usou, com a parte de cada um nos litros (ex.: Etanol 80% · Gasolina 20%)
      const porC = {}; o.ts.forEach(t => { const c = String(t.combustivel || "").trim().toUpperCase() || "NÃO INFORMADO"; porC[c] = (porC[c] || 0) + n0(numBR(t.litros)); });
      const combs = Object.keys(porC).sort((a, b) => porC[b] - porC[a]);
      const combTxt = combs.map(c => combs.length > 1 && litros ? `${c.charAt(0) + c.slice(1).toLowerCase()} ${Math.round(porC[c] / litros * 100)}%` : c.charAt(0) + c.slice(1).toLowerCase()).join(" · ");
      out.push({ nome: o.nome, combs, combTxt, contrato: maisFrequente(vs.map(v => v.operacao)), gestor: maisFrequente(vs.map(v => v.gestor_operacao)),
        classif: maisFrequente(o.ts.map(t => t.tipo_frota).concat(vs.map(v => v.tipo_frota))), abast: o.ts.length, litros, gasto, kmV,
        kml: litV ? kmV / litV : null, gastoV, rskm: kmV ? gastoV / kmV : null, inval: o.ts.length - val.length, pctInval: (o.ts.length - val.length) / o.ts.length * 100, ped: o.ped });
    });
    return out.sort((a, b) => b.gasto - a.gasto);
  }
  // tela própria: filtros, ordem pela coluna clicada (sempre do maior para o menor)
  const frk = { mes: "", gestor: "", contrato: "", busca: "", ord: "gasto", comb: "" };
  function desenharRanking() {
    const hoje = new Date(), atual = hoje.getFullYear() + "-" + String(hoje.getMonth() + 1).padStart(2, "0");
    const meses = Array.from(new Set(TRANS_MESES.concat([atual]))).sort().reverse();
    if (!frk.mes) frk.mes = meses.find(m => TRANS_C[m] && TRANS_C[m].length) || meses[0];
    $("#rk-mes").innerHTML = meses.map(m => `<option value="${m}">${nomeMes(m)}</option>`).join(""); $("#rk-mes").value = frk.mes;
    if (!TRANS_C[frk.mes]) { $("#rk-msg").textContent = "Carregando as transações…"; garantirTrans(frk.mes, desenharRanking); return; }
    $("#rk-msg").textContent = TRANS_C[frk.mes].length ? "" : "Sem transações nesse mês. Suba a base de transações em Consumo de combustível.";
    frk.gestor = opSel("#rk-gestor", gestoresDo($("#rk-contrato").value), "Todos os gestores");
    frk.contrato = opSel("#rk-contrato", contratosDo(frk.gestor), "Todos os contratos");
    const todasRows = transFiltradas(frk.mes, { g: frk.gestor, ct: frk.contrato });
    // lista todos os combustíveis que existem na base do mês (mesmo os que só aparecem em outro contrato, equipe ou gerador)
    frk.comb = opSel("#rk-comb", Array.from(new Set((TRANS_C[frk.mes] || []).filter(t => !ehPedagio(t)).map(t => String(t.combustivel || "").trim().toUpperCase()).filter(Boolean))).sort(), "Todos os combustíveis");
    // NOVO (24/11/2026): filtro pela coluna NUMERO FROTA da Ticket Log (ex.: GMG); vazio aparece como "Vazio"
    const nrF = t => nrFrotaDe(t);
    const doMes = (TRANS_C[frk.mes] || []).filter(t => !ehPedagio(t));
    const temNr = doMes.some(t => nrF(t));
    const opsF = Array.from(new Set(doMes.map(nrF).filter(Boolean))).sort();
    const selF = $("#rk-frota"), vF = selF.value;
    // enquanto a coluna NUMERO FROTA não foi gravada, separa o GMG pelo modelo/família/serviço (como nos alertas)
    const temGmg = doMes.some(ehGMG);
    selF.innerHTML = '<option value="">Número frota: todos</option>' + (temNr
      ? opsF.map(x => `<option value="${esc(x)}">${esc(x)}</option>`).join("") + (doMes.some(t => !nrF(t)) ? '<option value="__vazio__">Vazio</option>' : "")
      : (temGmg ? '<option value="__gmg__">GMG (gerador)</option><option value="__naogmg__">Sem GMG</option>' : ""));
    selF.value = Array.from(selF.options).some(o => o.value === vF) ? vF : "";
    selF.title = temNr ? "" : "As transações gravadas ainda não têm a coluna NUMERO FROTA: suba de novo a base de transações do mês.";
    frk.frota = selF.value;
    const bateF = t => !frk.frota || ehPedagio(t) || (frk.frota === "__vazio__" ? !nrF(t) : frk.frota === "__gmg__" ? ehGMG(t) : frk.frota === "__naogmg__" ? !ehGMG(t) : nrF(t) === frk.frota);
    // NOVO (25/11/2026): filtro MODELO VEICULO (ex.: CARTAO CORINGA, CARTAO OPERAÇÃO, STRADA…) — coringa primeiro
    const modF = t => String(t.modelo || "").replace(/\s+/g, " ").trim().toUpperCase();
    const ehCartaoX = m => /^CART(AO|ÃO)/.test(m);
    // tipos de cartão: CARTAO CORINGA, CARTAO OPERAÇÃO… e o resto (carros com placa) = "Cartão comum"
    const tipoCartao = t => { const m = modF(t); return ehCartaoX(m) ? m : "__comum__"; };
    const opsM = Array.from(new Set(doMes.map(modF).filter(ehCartaoX))).sort((x, y) => (/CORINGA/.test(y) - /CORINGA/.test(x)) || x.localeCompare(y));
    const selM = $("#rk-modelo"), vM = selM.value;
    selM.innerHTML = '<option value="">Tipo de cartão: todos</option>' + opsM.map(x => `<option value="${esc(x)}">${esc(x)}</option>`).join("") + (doMes.some(t => tipoCartao(t) === "__comum__") ? '<option value="__comum__">Cartão comum</option>' : "");
    selM.value = Array.from(selM.options).some(o => o.value === vM) ? vM : "";
    frk.modelo = selM.value;
    const bateM = t => !frk.modelo || ehPedagio(t) || tipoCartao(t) === frk.modelo;
    const rows = todasRows.filter(t => bateF(t) && bateM(t) && (!frk.comb || ehPedagio(t) || String(t.combustivel || "").trim().toUpperCase() === frk.comb));
    let r = rankingTecnicos(rows, !!(frk.frota || frk.modelo));
    if (frk.busca) r = r.filter(x => sem([x.nome, x.contrato, x.gestor].join(" ")).includes(frk.busca));
    const val = (x, k) => x[k] == null ? -Infinity : x[k];
    r.sort((a, b) => typeof a[frk.ord] === "string" ? sem(b[frk.ord]).localeCompare(sem(a[frk.ord])) : val(b, frk.ord) - val(a, frk.ord));
    const maxG = Math.max(0, ...r.map(x => x.gasto)), maxK = Math.max(0, ...r.map(x => x.kmV));
    const rs = r.map(x => x.rskm).filter(x => x != null).sort((a, b) => a - b), med = rs.length ? rs[Math.floor(rs.length / 2)] : 0;
    const corRsKm = v => v == null ? "" : v <= med * 0.85 ? "#bbf7d0" : v <= med * 1.15 ? "#fef08a" : v <= med * 1.4 ? "#fdba74" : "#fca5a5";
    const ped = r.some(x => x.ped);
    $("#cb-rank-info").textContent = `${r.length} técnicos · R$ ${fmtN(r.reduce((s, x) => s + x.gasto, 0), 2)} · R$/km mediano ${fmtN(med, 2)}`;
    const th = (k, t) => `<th class="num${frk.ord === k ? " ord" : ""}" data-k="${k}">${t}</th>`;
    $("#cb-rank-cab").innerHTML = `<tr><th class="num">#</th><th>Técnico</th><th>Contrato principal</th><th>Gestor</th><th>Classif.</th><th>Combustível</th>${th("abast", "Abast.")}${th("litros", "Litros")}${th("gasto", "Valor gasto")}
      ${th("kmV", "Km rodado (válido)")}${th("kml", "Km/L")}${th("gastoV", "Valor c/ km válido")}${th("rskm", "R$ por km")}${th("inval", "Abast. c/ km inválido")}${th("pctInval", "% abast. c/ km inválido")}${ped ? th("ped", "Pedágio") : ""}</tr>`;
    $$("#cb-rank-cab th[data-k]").forEach(h => h.addEventListener("click", () => { frk.ord = h.dataset.k; desenharRanking(); }));
    $("#tb-rank").innerHTML = r.map((x, i) => `<tr><td class="num">${i + 1}</td><td><b>${esc(x.nome)}</b></td><td>${esc(x.contrato)}</td><td>${esc(x.gestor)}</td><td>${esc(x.classif)}</td><td class="comb-txt">${esc(x.combTxt)}</td>
      <td class="num">${x.abast}</td><td class="num">${fmtN(x.litros)}</td>
      <td class="num barra-fundo" style="--p:${maxG ? x.gasto / maxG * 100 : 0}%;--c:#fdba74">R$ ${fmtN(x.gasto, 2)}</td>
      <td class="num barra-fundo" style="--p:${maxK ? x.kmV / maxK * 100 : 0}%;--c:#93c5fd">${fmtN(x.kmV)}</td>
      <td class="num">${fmtN(x.kml, 2)}</td><td class="num">R$ ${fmtN(x.gastoV, 2)}</td>
      <td class="num" style="background:${corRsKm(x.rskm)};font-weight:700">${x.rskm == null ? "-" : "R$ " + fmtN(x.rskm, 2)}</td>
      <td class="num">${x.inval || "-"}</td><td class="num">${x.inval ? fmtN(x.pctInval, 1) + "%" : "-"}</td>${ped ? `<td class="num">${x.ped ? "R$ " + fmtN(x.ped, 2) : "-"}</td>` : ""}</tr>`).join("")
      || `<tr><td colspan="16" class="info">Nenhum técnico com esse filtro.</td></tr>`;
    RANK_ATUAL = { mes: frk.mes, r, ped };
    desenharPorGestor(frk.mes, todasRows);
  }
  // por gestor: mês escolhido x mesmo período do mês anterior (do dia 1 até o último dia com transação no mês)
  const mesAnterior = m => { const [a, b] = m.split("-").map(Number); return b === 1 ? `${a - 1}-12` : `${a}-${String(b - 1).padStart(2, "0")}`; };
  function desenharPorGestor(mes, rows) {
    const ant = mesAnterior(mes);
    if (!TRANS_C[ant]) garantirTrans(ant, () => { if (!$("#aba-rank").hidden) desenharRanking(); });
    const sem_ped = l => l.filter(t => !ehPedagio(t));
    const atual = sem_ped(rows);
    const ultDia = atual.reduce((m, t) => Math.max(m, Number(String(t.data).slice(8, 10)) || 0), 0);
    const anterior = sem_ped(transFiltradas(ant, { g: frk.gestor, ct: frk.contrato })).filter(t => (Number(String(t.data).slice(8, 10)) || 0) <= ultDia);
    const alertaSet = new Set(analisarTrans(rows).filter(a => a.alerta).map(a => a.t));
    const gestorDe = t => donoTrans(t).gestor || "Sem gestor (placa e motorista fora da base)";
    const g = {};
    const soma = (l, campo) => { l.forEach(t => { const k = gestorDe(t); g[k] = g[k] || { ab: 0, l: 0, v: 0, va: 0, al: 0 }; if (campo === "v") { g[k].ab++; g[k].l += n0(numBR(t.litros)); g[k].v += n0(numBR(t.valor)); if (alertaSet.has(t)) g[k].al++; } else g[k].va += n0(numBR(t.valor)); }); };
    soma(atual, "v"); soma(anterior, "va");
    const tot = Object.values(g).reduce((s, x) => s + x.v, 0);
    const lista = Object.keys(g).filter(k => g[k].v || g[k].va).sort((a, b) => g[b].v - g[a].v);
    const maxV = Math.max(0, ...lista.map(k => g[k].v));
    const dd = String(ultDia || 0).padStart(2, "0"), per = `(01–${dd})`;
    const nm = m => nomeMes(m).split("/")[0].replace(/^./, c => c.toUpperCase());
    const temAnt = (TRANS_C[ant] || []).length > 0;
    $("#rk-ges-info").textContent = `${lista.length} gestores · ${nm(mes)} ${per}` + (temAnt ? ` x ${nm(ant)} ${per}` : "");
    $("#rk-ges-cab").innerHTML = `<tr><th class="num">#</th><th>Gestor</th><th class="num">Abast.</th><th class="num">Litros</th><th class="num">Valor ${nm(mes)} ${per}</th><th class="num">% do total</th>
      <th class="num">Valor ${nm(ant)} ${per}</th><th class="num">Variação R$</th><th class="num">Variação %</th><th class="num">Abast. c/ alerta</th></tr>`;
    $("#tb-rk-ges").innerHTML = lista.map((k, i) => { const x = g[k], dv = x.v - x.va, dp = x.va ? dv / x.va * 100 : null;
      return `<tr><td class="num">${i + 1}</td><td><b>${esc(k)}</b></td><td class="num">${x.ab.toLocaleString("pt-BR")}</td><td class="num">${fmtN(x.l)}</td>
      <td class="num barra-fundo" style="--p:${maxV ? x.v / maxV * 100 : 0}%;--c:#fdba74">R$ ${fmtN(x.v, 2)}</td><td class="num">${tot ? fmtN(x.v / tot * 100, 1) + "%" : "-"}</td>
      <td class="num ant">${temAnt ? "R$ " + fmtN(x.va, 2) : "-"}</td>
      <td class="num ${dv > 0 ? "var-mais" : "var-menos"}">${temAnt ? (dv < 0 ? "-" : "") + "R$ " + fmtN(Math.abs(dv), 2) : "-"}</td>
      <td class="num ${dv > 0 ? "var-mais" : "var-menos"}">${temAnt && dp != null ? fmtN(dp, 1) + "%" : "-"}</td><td class="num">${x.al || "-"}</td></tr>`; }).join("") || '<tr><td colspan="10" class="info">Sem transações.</td></tr>';
    // NOVO (24/11/2026): "Sem gestor" aberto por INFORMACAO ADICIONAL 1 (placas que não estão na base de veículos)
    const semG = atual.filter(t => !donoTrans(t).gestor);
    const box = $("#rk-semg");
    if (box) {
      if (!semG.length) box.innerHTML = "";
      else {
        const temInfo = semG.some(t => String(t.info1 || "").trim());
        const gi = {};
        semG.forEach(t => { const k = String(t.info1 || "").trim() || "(sem informação adicional 1)"; const o = gi[k] = gi[k] || { pl: new Set(), ab: 0, v: 0 }; o.pl.add(t.placa); o.ab++; o.v += n0(numBR(t.valor)); });
        const ks = Object.keys(gi).sort((x, y) => gi[y].v - gi[x].v), tv = semG.reduce((s, t) => s + n0(numBR(t.valor)), 0);
        box.innerHTML = `<h3 style="margin:18px 0 4px">Sem gestor — por Informação adicional 1 <small class="info">${semG.length} abastecimentos de ${new Set(semG.map(t => t.placa)).size} placas fora da base de veículos e com motorista que não é condutor de nenhum carro nem está na hierarquia · passe o mouse em "placas" para ver quais são</small></h3>` +
          (temInfo ? "" : `<p class="nota">As transações gravadas ainda não têm a coluna INFORMACAO ADICIONAL 1. Suba de novo a base de transações deste mês (botão "Subir transações"): ela preenche essa coluna sem duplicar nada.</p>`) +
          `<div style="overflow:auto"><table class="tab-rank"><thead><tr><th>Informação adicional 1</th><th class="num">Placas</th><th class="num">Abast.</th><th class="num">Valor</th><th class="num">% do sem gestor</th></tr></thead><tbody>` +
          ks.map(k => { const o = gi[k], pls = Array.from(o.pl).sort();
            return `<tr><td><b>${esc(k)}</b></td><td class="num" title="${esc(pls.join(", "))}" style="cursor:help;text-decoration:underline dotted">${pls.length}</td><td class="num">${o.ab}</td><td class="num">R$ ${fmtN(o.v, 2)}</td><td class="num">${tv ? fmtN(o.v / tv * 100, 1) + "%" : "-"}</td></tr>`; }).join("") +
          `</tbody></table></div>`;
      }
    }
    $("#rk-ges-nota").textContent = temAnt ? `Comparação com o mesmo período de ${nm(ant)} (01 a ${dd}). Variação positiva = gastou mais que em ${nm(ant)}.`
      : `Para comparar com ${nm(ant)}, suba a base de transações de ${nomeMes(ant)} no botão "Subir transações (qualquer mês)".`;
  }
  $("#rk-mes").addEventListener("change", e => { frk.mes = e.target.value; desenharRanking(); });
  $("#arq-trans-rk").addEventListener("change", async e => {
    const arqs = Array.from(e.target.files || []); e.target.value = "";
    const caixa = $("#rk-msgs"); caixa.innerHTML = "";
    for (const arq of arqs) { const el = document.createElement("p"); el.className = "msg"; caixa.appendChild(el); await subirArquivoComb(arq, el); }
    desenharRanking();
  });
  $("#rk-gestor").addEventListener("change", () => desenharRanking());
  $("#rk-contrato").addEventListener("change", () => desenharRanking());
  $("#rk-comb").addEventListener("change", () => desenharRanking());
  $("#rk-frota").addEventListener("change", () => desenharRanking());
  $("#rk-modelo").addEventListener("change", () => desenharRanking());
  $("#rk-busca").addEventListener("input", e => { frk.busca = sem(e.target.value); desenharRanking(); });
  let RANK_ATUAL = null;
  // ---------- aba Alertas de combustível (por condutor)
  const fal = { mes: "", busca: "", ord: "gasto", so: true, aberto: "" };
  let AL_ATUAL = null;
  function desenharAlertasComb() {
    const hoje = new Date(), atual = hoje.getFullYear() + "-" + String(hoje.getMonth() + 1).padStart(2, "0");
    const meses = Array.from(new Set(TRANS_MESES.concat([atual]))).sort().reverse();
    if (!fal.mes) fal.mes = meses.find(m => TRANS_C[m] && TRANS_C[m].length) || meses[0];
    $("#al-mes").innerHTML = meses.map(m => `<option value="${m}">${nomeMes(m)}</option>`).join(""); $("#al-mes").value = fal.mes;
    if (!TRANS_C[fal.mes]) { $("#al-msg").textContent = "Carregando as transações…"; garantirTrans(fal.mes, desenharAlertasComb); return; }
    $("#al-msg").textContent = TRANS_C[fal.mes].length ? "" : "Sem transações nesse mês. Suba a base de transações.";
    const g = opSel("#al-gestor", gestoresDo($("#al-contrato").value), "Todos os gestores");
    const ct = opSel("#al-contrato", contratosDo(g), "Todos os contratos");
    const an = analisarTrans(transFiltradas(fal.mes, { g, ct }));
    let l = alertasPorCondutor(an);
    const tot = { consumo: 0, kmInv: 0, dia2: 0, l80: 0, preco: 0, aditivado: 0, oficina: 0 };
    l.forEach(x => Object.keys(tot).forEach(k => { tot[k] += x[k]; }));
    const totAl = l.reduce((s, x) => s + x.comAlerta, 0), totAb = l.reduce((s, x) => s + x.abast, 0);
    $("#al-chips").innerHTML = `<span class="chip vermelho">${totAl} de ${totAb} abastecimentos com alerta</span>` + Object.keys(tot).filter(k => tot[k]).map(k => `<span class="chip laranja">${esc(NOMES_ALERTA[k])}: ${tot[k]}</span>`).join("")
      + `<span class="chip">R$ ${fmtN(l.reduce((s, x) => s + x.gastoAlerta, 0), 2)} em abastecimentos com alerta</span>`;
    if (fal.so) l = l.filter(x => x.comAlerta);
    if (fal.busca) l = l.filter(x => sem([x.nome, x.gestor, x.contrato].join(" ")).includes(fal.busca));
    const vv = (x, k) => x[k] == null ? -Infinity : x[k];
    l.sort((a, b) => vv(b, fal.ord) - vv(a, fal.ord));
    const maxP = Math.max(1, ...l.map(x => x.pontos));
    const th = (k, t) => `<th class="num${fal.ord === k ? " ord" : ""}" data-k="${k}">${t}</th>`;
    $("#al-cab").innerHTML = `<tr><th class="num">#</th><th>Condutor</th><th>Gestor</th><th>Modelo</th><th>Combustível</th>${th("abast", "Abast.")}${th("comAlerta", "Abast. c/ alerta")}${th("consumo", "Consumo &lt; 5 km/l")}${th("kmInv", "Km inválido")}
      ${th("dia2", "2+ no dia")}${th("l80", "&gt; 80 L")}${th("preco", "Preço alto")}${th("aditivado", "Aditivado")}${th("pontos", "Pontos")}${th("kml", "Km/l")}${th("gasto", "Valor gasto")}${th("gastoAlerta", "Valor c/ alerta")}<th>O que olhar</th></tr>`;
    $$("#al-cab th[data-k]").forEach(h => h.addEventListener("click", () => { fal.ord = h.dataset.k; desenharAlertasComb(); }));
    const n0s = v => v ? `<b style="color:#b91c1c">${v}</b>` : "-";
    $("#al-info").textContent = `${nomeMes(fal.mes)} · ${l.length} condutores`;
    $("#tb-al").innerHTML = l.map((x, i) => `<tr class="clica${x.comAlerta ? " tem-alerta" : ""}" data-nome="${esc(x.nome)}"><td class="num">${i + 1}</td><td><b>${esc(x.nome)}</b></td><td>${esc(x.gestor)}</td><td class="modelo">${esc(x.modelo)}</td><td class="comb-txt">${esc(x.combTxt)}</td>
      <td class="num">${x.abast}</td><td class="num">${n0s(x.comAlerta)}</td><td class="num">${n0s(x.consumo)}</td><td class="num">${n0s(x.kmInv)}</td><td class="num">${n0s(x.dia2)}</td>
      <td class="num">${n0s(x.l80)}</td><td class="num">${n0s(x.preco)}</td><td class="num">${n0s(x.aditivado)}</td>
      <td class="num pontos barra-fundo" style="--p:${x.pontos / maxP * 100}%;--c:#fdba74">${fmtN(x.pontos, 1)}</td><td class="num">${fmtN(x.kml, 2)}</td>
      <td class="num">R$ ${fmtN(x.gasto, 2)}</td><td class="num">R$ ${fmtN(x.gastoAlerta, 2)}</td><td class="olhar">${esc(x.olhar)}</td></tr>`
      + (fal.aberto === x.nome ? `<tr class="det-al"><td></td><td colspan="17">${x.itens.filter(a => a.alerta || a.hodo).sort((a, b) => String(a.t.data).localeCompare(String(b.t.data))).map(a =>
        `${esc(String(a.t.data).slice(8, 10) + "/" + String(a.t.data).slice(5, 7) + " " + String(a.t.data).slice(11, 16))} · <b>${esc(a.t.placa)}</b> ${esc(a.t.modelo || "")} · ${esc(a.t.combustivel || "")} · ${fmtN(a.l, 2)} L · ${fmtN(a.km)} km (${fmtN(a.kml, 2)} km/l) · R$ ${fmtN(a.preco, 2)}/L · ${esc(a.t.estabelecimento || "")} ${esc(a.t.uf || "")} → <b style="color:#b91c1c">${Object.keys(a.f).filter(k => a.f[k]).map(k => NOMES_ALERTA[k]).concat(a.hodo ? ["hodômetro fixo"] : []).join(", ")}</b>`).join("<br>")}</td></tr>` : "")).join("")
      || '<tr><td colspan="18" class="info">Nenhum condutor com alerta nesse filtro. 👍</td></tr>';
    AL_ATUAL = { mes: fal.mes, l };
  }
  $("#tb-al").addEventListener("click", e => { const tr = e.target.closest("tr[data-nome]"); if (!tr) return; fal.aberto = fal.aberto === tr.dataset.nome ? "" : tr.dataset.nome; desenharAlertasComb(); });
  $("#al-mes").addEventListener("change", e => { fal.mes = e.target.value; desenharAlertasComb(); });
  $("#al-gestor").addEventListener("change", () => desenharAlertasComb());
  $("#al-contrato").addEventListener("change", () => desenharAlertasComb());
  $("#al-busca").addEventListener("input", e => { fal.busca = sem(e.target.value); desenharAlertasComb(); });
  $("#al-so").addEventListener("change", e => { fal.so = e.target.checked; desenharAlertasComb(); });
  $("#bt-al-exp").addEventListener("click", () => {
    if (!AL_ATUAL) return;
    const r2 = v => v == null ? null : Math.round(v * 100) / 100;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AL_ATUAL.l.map((x, i) => ({ "#": i + 1, "Condutor": x.nome, "Gestor": x.gestor, "Modelo": x.modelo, "Combustível": x.combTxt, "Abast.": x.abast, "Abast. c/ alerta": x.comAlerta,
      "Consumo < 5 km/l": x.consumo, "Km inválido": x.kmInv, "2+ no dia": x.dia2, "> 80 L": x.l80, "Preço alto": x.preco, "Aditivado": x.aditivado, "Pontos": r2(x.pontos), "Km/l": r2(x.kml),
      "Valor gasto": r2(x.gasto), "Valor c/ alerta": r2(x.gastoAlerta), "O que olhar": x.olhar }))), "Por condutor");
    const det = [];
    AL_ATUAL.l.forEach(x => x.itens.filter(a => a.alerta || a.hodo).forEach(a => det.push({ "Data": a.t.data, "Condutor": x.nome, "Placa": a.t.placa, "Modelo": a.t.modelo, "Combustível": a.t.combustivel, "Litros": r2(a.l),
      "Km rodado": a.km, "Km/l": r2(a.kml), "R$/litro": r2(a.preco), "Mediana R$/litro (combustível + UF)": r2(a.medPreco), "Valor": r2(n0(numBR(a.t.valor))), "Posto": a.t.estabelecimento, "UF": a.t.uf,
      "Alertas": Object.keys(a.f).filter(k => a.f[k]).map(k => NOMES_ALERTA[k]).concat(a.hodo ? ["hodômetro fixo"] : []).join(", ") })));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(det.length ? det : [{ "Data": "" }]), "Abastecimentos com alerta");
    XLSX.writeFile(wb, `Alertas_combustivel_${AL_ATUAL.mes}.xlsx`);
  });
  $("#bt-rank-exp").addEventListener("click", () => {
    if (!RANK_ATUAL) return;
    const r2 = v => v == null ? null : Math.round(v * 100) / 100;
    const ws = XLSX.utils.json_to_sheet(RANK_ATUAL.r.map((x, i) => Object.assign({ "#": i + 1, "Técnico": x.nome, "Contrato principal": x.contrato, "Gestor": x.gestor, "Classif.": x.classif, "Combustível": x.combTxt,
      "Abast.": x.abast, "Litros": r2(x.litros), "Valor gasto": r2(x.gasto), "Km rodado (válido)": x.kmV, "Km/L": r2(x.kml), "Valor c/ km válido": r2(x.gastoV), "R$ por km": r2(x.rskm),
      "Abast. c/ km inválido": x.inval, "% abast. c/ km inválido": r2(x.pctInval) }, RANK_ATUAL.ped ? { "Pedágio": r2(x.ped) } : {})));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Gasto x Km por técnico"); XLSX.writeFile(wb, `Gasto_x_Km_tecnico_${RANK_ATUAL.mes}.xlsx`);
  });

  // mais frequente motorista de cada placa no mês (para a coluna "Quem abastece")
  function motoristaDoMes(mes) {
    const m = new Map();
    (TRANS_C[mes] || []).forEach(t => { const n = t.motorista || ""; if (!n) return; const c = m.get(t.placa) || {}; c[n] = (c[n] || 0) + 1; m.set(t.placa, c); });
    const out = new Map();
    m.forEach((c, p) => out.set(p, Object.keys(c).sort((a, b) => c[b] - c[a])));
    return out;
  }

  async function subirTransacoes(m, el) {
    const limpa = h => sem(h).replace(/[^a-z0-9/ ]/g, "").replace(/\s+/g, " ").trim();
    const r0 = m.findIndex(l => l.some(c => limpa(c) === "codigo transacao"));
    const cab = m[r0].map(limpa), col = n => cab.indexOf(n);
    const ix = { codigo: col("codigo transacao"), data: col("data transacao"), placa: col("placa"), motorista: col("nome motorista"), matricula: col("matricula"),
      responsavel: col("responsavel"), cartao: col("numero cartao"), modelo: col("modelo veiculo"), combustivel: col("tipo combustivel"), litros: col("litros"),
      vl_litro: col("vl/litro"), valor: col("valor emissao"), hodometro: col("hodometro ou horimetro"), km_rodados: col("km rodados ou horas trabalhadas"),
      km_litro: col("km/litro ou litros/hora"), estabelecimento: col("nome estabelecimento"), cidade: col("cidade"), uf: col("uf"),
      servico: col("servico"), tipo_frota: col("tipo frota"), familia: col("familia veiculo"),
      nr_frota: cab.findIndex(h => /^(numero|nr|n|no|num)\.? ?(da |de )?frota/.test(h)),
      // INFORMACAO ADICIONAL 1 (contrato/centro de custo do cartão na Ticket Log) — aceita "adidional" (como vem no arquivo)
      info1: (() => { const l = cab.map((h, i) => [h, i]).filter(([h]) => /^informac?ao adi/.test(h)); const um = l.find(([h]) => /1$/.test(h)); return um ? um[1] : (l[0] ? l[0][1] : -1); })() };
    if (ix.data < 0 || ix.placa < 0 || ix.litros < 0) return msg(el, "Base de transações sem as colunas DATA TRANSACAO, PLACA ou LITROS.", "erro");
    const dataIso = v => {
      if (typeof v === "number" && v > 20000) { const d = new Date(Math.round((v - 25569) * 86400000)); return d.toISOString().slice(0, 19).replace("T", " "); }
      const t = String(v || "").trim(), x = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\D+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
      if (x) return `${x[3].length === 2 ? "20" + x[3] : x[3]}-${x[2].padStart(2, "0")}-${x[1].padStart(2, "0")} ${(x[4] || "00").padStart(2, "0")}:${x[5] || "00"}:${x[6] || "00"}`;
      const iso = t.match(/^(\d{4}-\d{2}-\d{2})[ T]?(\d{2}:\d{2}(:\d{2})?)?/); return iso ? `${iso[1]} ${iso[2] ? (iso[2].length === 5 ? iso[2] + ":00" : iso[2]) : "00:00:00"}` : "";
    };
    const regs = [];
    m.slice(r0 + 1).forEach(l => {
      const txt = k => ix[k] >= 0 ? String(l[ix[k]] == null ? "" : l[ix[k]]).trim() : "";
      const num = k => { const n = ix[k] >= 0 ? numBR(l[ix[k]]) : null; return n == null ? "" : n; };
      const placa = txt("placa").toUpperCase().replace(/[^A-Z0-9]/g, "");
      const data = dataIso(ix.data >= 0 ? l[ix.data] : "");
      if (placa.length < 6 || !data) return;
      regs.push({ codigo: txt("codigo"), data, mes: data.slice(0, 7), placa, motorista: txt("motorista"), matricula: txt("matricula"), responsavel: txt("responsavel"),
        cartao: txt("cartao"), modelo: txt("modelo"), combustivel: txt("combustivel"), litros: num("litros"), vl_litro: num("vl_litro"), valor: num("valor"),
        hodometro: num("hodometro"), km_rodados: num("km_rodados"), km_litro: num("km_litro"), estabelecimento: txt("estabelecimento"), cidade: txt("cidade"), uf: txt("uf"),
        servico: txt("servico"), tipo_frota: txt("tipo_frota"), familia: txt("familia"), info1: txt("info1"), nr_frota: txt("nr_frota") });
    });
    if (!regs.length) return msg(el, "Nenhuma transação com placa e data na planilha.", "erro");
    let novos = 0, rep = 0, inf = 0;
    for (let i = 0; i < regs.length; i += 500) {
      msg(el, `Enviando transações ${i + 1} a ${Math.min(i + 500, regs.length)} de ${regs.length}…`);
      const r = await api("subir_transacoes", { regs: regs.slice(i, i + 500) });
      novos += (r.r && r.r.novos) || 0; rep += (r.r && r.r.repetidas) || 0; inf += (r.r && r.r.info) || 0;
    }
    const meses = Array.from(new Set(regs.map(r => r.mes)));
    meses.forEach(mm => { delete TRANS_C[mm]; });
    const al = alertasTrans(regs);
    msg(el, `✓ Transações: ${regs.length} abastecimentos (${novos} novos${rep ? ", " + rep + " já estavam gravados" : ""}${inf ? ", " + inf + " antigas completadas (Número frota / Informação adicional 1)" : ""}) de ${meses.map(nomeMes).join(", ")} · ${al.filter(a => a.dia2).length} ${al.filter(a => a.dia2).length === 1 ? "dia" : "dias"} com 2+ abastecimentos · ${al.reduce((s, a) => s + a.acima, 0)} acima de 80 L.`, "ok");
    if (!fcb.mes || meses.indexOf(fcb.mes) < 0) fcb.mes = fcb.mes || meses.sort().pop();
    desenharTrans(); desenharInicio();
  }

  // ---------- extrato simplificado
  async function subirExtrato(m, el) {
    const limpa = h => sem(h).replace(/[^a-z0-9/ ]/g, "").replace(/\s+/g, " ").trim();
    const dinheiro = v => numBR(String(v == null ? "" : v).split("=")[0]);
    const rh = m.findIndex(l => l.some(c => limpa(c) === "cartao") && l.some(c => limpa(c) === "placa"));
    const res = {}; let secao = "";
    m.slice(0, rh < 0 ? 80 : rh).forEach(l => {
      const rot = limpa(l[0]); if (!rot) return;
      const val = l.slice(1).map(c => String(c == null ? "" : c).trim()).find(Boolean) || "";
      if (rot.startsWith("limite dos cartoes periodo atual")) { secao = "atual"; return; }
      if (rot.startsWith("limite dos cartoes proxima")) { secao = "prox"; return; }
      if (rot.startsWith("filtros")) { secao = "fim"; return; }
      if (!val) return;
      if (rot === "limite contratual") res.limite_contratual = dinheiro(val);
      else if (rot.startsWith("valor contratual consumido")) { res.consumido_contratual = dinheiro(val); const p = val.match(/=\s*([\d.,]+)\s*%/); if (p) res.pct_consumido = numBR(p[1]); }
      else if (rot.startsWith("saldo contratual")) res.saldo_contratual = dinheiro(val);
      else if (secao === "atual" && rot === "limite total") res.limite_total = dinheiro(val);
      else if (secao === "atual" && rot.startsWith("limite distribuido")) res.limite_distribuido = dinheiro(val);
      else if (secao === "atual" && rot.startsWith("valor disponivel")) res.disponivel_alocacao = dinheiro(val);
      else if (secao === "atual" && rot.startsWith("valor consumido")) res.consumido_compras = dinheiro(val);
      else if (secao === "prox" && rot.startsWith("data da proxima")) res.proxima_renovacao = val;
      else if (secao === "prox" && rot.startsWith("limite distribuido")) res.limite_distribuido_prox = dinheiro(val);
    });
    const dt = m.slice(0, 15).map(l => l.join(" ")).join(" ").match(/\d{2}\/\d{2}\/\d{4}/);
    res.data_ref = dt && !/renova/i.test(m.slice(0, 15).map(l => l.join(" ")).join(" ").split(dt[0])[0].slice(-60)) ? dt[0] : new Date().toLocaleDateString("pt-BR");
    const cartoes = [];
    if (rh >= 0) {
      const cab = m[rh].map(limpa), cab2 = (m[rh + 1] || []).map(limpa);
      const c1 = t => cab.findIndex(t), c2 = t => cab2.findIndex(t);
      const ix = { cartao: c1(h => h === "cartao"), placa: c1(h => h === "placa"), nr_frota: c1(h => h.startsWith("nr")), familia: c1(h => h.startsWith("famil")),
        tipo_frota: c1(h => h.startsWith("tipo frota")), modelo: c1(h => h === "modelo"), fabricante: c1(h => h.startsWith("fabric")), cidade: c1(h => h.startsWith("cidade")),
        responsavel: c1(h => h.startsWith("nome resp")), limite_proximo: c1(h => h.startsWith("limite proximo")),
        valor_reservado: c2(h => h.startsWith("valor reserv")), limite_atual: c2(h => h.startsWith("limite atual")), utilizado: c2(h => h.startsWith("compras")), saldo: c2(h => h === "saldo") };
      const corpo = m.slice(rh + 2);
      const so1 = l => l.filter(c => String(c == null ? "" : c).trim() !== "").length === 1 && String(l[0] == null ? "" : l[0]).trim() !== "";
      const ehCartao = l => /\*|\d{4}\s?\d{2}/.test(String(l[ix.cartao] == null ? "" : l[ix.cartao])) && String(l[ix.placa] || l[ix.cartao] || "").trim() !== "";
      let cc = "", grupo = "";
      corpo.forEach((l, i) => {
        if (so1(l)) {
          const prox = corpo.slice(i + 1).find(x => x.some(c => String(c == null ? "" : c).trim() !== ""));
          if (prox && so1(prox)) { cc = String(l[0]).trim(); grupo = ""; } else grupo = String(l[0]).trim();
          return;
        }
        if (!ehCartao(l)) return;
        const txt = k => ix[k] >= 0 ? String(l[ix[k]] == null ? "" : l[ix[k]]).trim() : "";
        const num = k => ix[k] >= 0 ? numBR(l[ix[k]]) : null;
        const r = { cartao: txt("cartao"), placa: txt("placa").toUpperCase().replace(/[^A-Z0-9]/g, ""), centro_custo: cc, grupo, nr_frota: txt("nr_frota"), familia: txt("familia"),
          tipo_frota: txt("tipo_frota"), modelo: txt("modelo"), fabricante: txt("fabricante"), cidade: txt("cidade"), responsavel: txt("responsavel"),
          valor_reservado: num("valor_reservado"), limite_atual: num("limite_atual"), utilizado: num("utilizado"), saldo: num("saldo"), limite_proximo: num("limite_proximo") };
        Object.keys(r).forEach(k => { if (r[k] === null) r[k] = ""; });
        cartoes.push(r);
      });
    }
    if (res.saldo_contratual == null && !cartoes.length) return msg(el, "Não consegui ler o Extrato Simplificado.", "erro");
    msg(el, `Enviando extrato (${cartoes.length} cartões)…`);
    await api("subir_extrato", { dados: res, regs: cartoes });
    montarFiltrosComb(); desenharComb(); desenharInicioComb();
    const soma = cartoes.reduce((a, c) => a + n0(c.utilizado), 0);
    msg(el, `✓ Extrato de ${res.data_ref} gravado: saldo contratual R$ ${fmtN(res.saldo_contratual, 2)}, consumido no período R$ ${fmtN(res.consumido_compras, 2)}, ${cartoes.length} cartões (compras somadas R$ ${fmtN(soma, 2)}).`, "ok");
  }



  // início: cartões "Limite disponível" e "Gasto no mês" (mês atual; filtra por centro de custo)
  function desenharInicioComb() {
    const ext = EXTRATO[0];
    $("#ini-comb").hidden = !COMB.length && !CARTOES.length && !ext;
    $("#ic-cc-txt").hidden = $("#ini-comb").hidden;
    if ($("#ini-comb").hidden) { $("#ini-graf-caixa").hidden = true; return; }
    // só o mês atual
    const hoje = new Date();
    const mes = hoje.getFullYear() + "-" + String(hoje.getMonth() + 1).padStart(2, "0");
    const doMes = COMB.filter(x => x.mes === mes);
    atualizarListaCC();
    const ccs = Array.from(new Set(doMes.map(x => x.centro_custo).concat(CARTOES.map(x => x.centro_custo)).filter(Boolean))).sort();
    const txtCc = ($("#ic-cc-txt").value || "").trim();
    const cc = ccDoCampo(txtCc);
    const g = $("#ic-gestor").value, ct = $("#ic-contrato").value;
    const filtroIni = x => bateCC(x.centro_custo, cc) && bateDono(x.placa, x, g, ct);
    const rs = doMes.filter(filtroIni);
    const litros = rs.reduce((a, x) => a + n0(x.litros), 0), kmIni = rs.reduce((a, x) => a + n0(x.km_rodados), 0);
    const ca = CARTOES.filter(filtroIni);
    const per = (doMes.find(x => x.periodo) || {}).periodo;
    let saldo, sub, gasto, periodo, det;
    if (ext && !ESCOPO && !cc && !g && !ct && !equipeAtiva()) {
      // contrato inteiro: números do Extrato Simplificado (iguais ao portal)
      saldo = n0(ext.saldo_contratual);
      sub = `${fmtN(100 - n0(ext.pct_consumido), 2)}% disponível${ext.proxima_renovacao ? " · renova em " + ext.proxima_renovacao : ""}`;
      gasto = n0(ext.consumido_compras);
      periodo = `período atual${ext.proxima_renovacao ? " · renova em " + ext.proxima_renovacao : ""} · extrato de ${ext.data_ref}`;
      det = `${ca.length} cartões` + (litros ? ` · ${fmtN(litros)} litros · ${fmtN(kmIni)} km em ${nomeMes(mes)}` : "") + ` · ${ccs.length} centros de custo`;
    } else if (ca.length) {
      // um centro de custo (ou login de equipe): soma dos cartões do extrato
      saldo = ca.reduce((a, x) => a + n0(x.saldo), 0);
      const lim = ca.reduce((a, x) => a + n0(x.limite_atual), 0);
      sub = `de R$ ${fmtN(lim, 2)} de limite · ${ca.length} ${ca.length === 1 ? "cartão" : "cartões"}`;
      gasto = ca.reduce((a, x) => a + n0(x.utilizado), 0);
      periodo = `período atual · extrato de ${ca[0].data_ref || "-"}`;
      det = (litros ? `${fmtN(litros)} litros · ${fmtN(kmIni)} km · ` : "") + `${new Set(ca.map(x => x.placa)).size} carro(s)`;
    } else {
      // sem extrato: usa o relatório de consumo do mês
      const porPlaca = new Map();
      rs.forEach(x => { if (!porPlaca.has(x.placa) && x.saldo !== "") porPlaca.set(x.placa, x); });
      saldo = Array.from(porPlaca.values()).reduce((a, x) => a + n0(x.saldo), 0);
      const lim = Array.from(porPlaca.values()).reduce((a, x) => a + n0(x.limite), 0);
      sub = lim ? `de R$ ${fmtN(lim, 2)} de limite · ${porPlaca.size} ${porPlaca.size === 1 ? "cartão" : "cartões"} (suba o Extrato Simplificado para o saldo do contrato)` : "";
      gasto = rs.reduce((a, x) => a + n0(x.total), 0);
      periodo = doMes.length ? (per ? per + " · " + nomeMes(mes) : nomeMes(mes)) : "Ainda não subiu o consumo de " + nomeMes(mes);
      det = `${fmtN(litros)} litros · ${fmtN(kmIni)} km · ${new Set(rs.map(x => x.placa)).size} carro(s)`;
    }
    // gasto do mês de calendário (igual ao portal): soma das transações de 01/mm até a última transação
    if (!TRANS_C[mes]) garantirTrans(mes, () => desenharInicioComb());
    const trMes = TRANS_C[mes] || [];
    if (trMes.length) {
      const ccPl = new Map(TODOS.cart.map(c => [String(c.placa || "").toUpperCase(), c.centro_custo]));
      const vis = new Set(VEICULOS.map(v => v.placa));
      const trs = trMes.filter(t => noEscopo(t.placa) && (!equipeAtiva() || vis.has(t.placa)) && bateCC(ccPl.get(t.placa) || "", cc) && bateDono(t.placa, null, g, ct));
      const gt = trs.reduce((a, t) => a + n0(numBR(t.valor)), 0), lt = trs.reduce((a, t) => a + n0(numBR(t.litros)), 0), kt = trs.reduce((a, t) => a + n0(numBR(t.km_rodados)), 0);
      const ult = trs.reduce((m, t) => String(t.data) > m ? String(t.data) : m, "");
      gasto = gt;
      periodo = `de 01/${mes.slice(5, 7)} até ${ult ? ult.slice(8, 10) + "/" + ult.slice(5, 7) : "hoje"} · transações` + (ext && !cc && !g && !ct && !equipeAtiva() && ext.proxima_renovacao ? ` · cartões renovam em ${ext.proxima_renovacao}` : "");
      det = `${trs.length.toLocaleString("pt-BR")} abastecimentos · ${fmtN(lt)} litros · ${fmtN(kt)} km` + (kt ? ` · R$/km ${fmtN(gt / kt, 2)}` : "");
    }
    $("#ic-saldo").textContent = "R$ " + fmtN(saldo, 2);
    $("#ic-limite").textContent = sub;
    $("#ic-periodo").textContent = periodo;
    $("#ic-gasto").textContent = "R$ " + fmtN(gasto, 2);
    $("#ic-det").textContent = det;
    if (cc.startsWith("texto:")) $("#ic-det").textContent += ` · ${new Set(CARTOES.concat(doMes).map(x => x.centro_custo).filter(c => c && bateCC(c, cc))).size} centros com “${txtCc}”`;
    desenharGrafComb(filtroIni, [cc ? nomeCC(cc) : "", g, ct].filter(Boolean).join(" · "));
  }
  // gráfico: barras por combustível em cada mês + linha do total (como no portal da Ticket Log)
  const GRUPOS_COMB = [["Diesel", "#3a3a3a", t => t.includes("diesel")], ["Etanol", "#0f9d7a", t => t.includes("etanol") || t.includes("alcool")],
    ["Gás natural", "#11b5e4", t => t.includes("gnv") || t.includes("gas natural")], ["Gasolina", "#f0291f", t => t.includes("gasolina")], ["Não especificado", "#111111", () => true]];
  let GRAF = null;
  function desenharGrafComb(filtro, rotulo) {
    const caixa = $("#ini-graf-caixa");
    if (!window.Chart || !COMB.length) { caixa.hidden = true; return; }
    caixa.hidden = false;
    const meses = Array.from(new Set(COMB.map(x => x.mes))).sort().slice(-12);
    const soma = {};
    COMB.filter(filtro).forEach(x => {
      const t = sem(x.combustivel); const g = GRUPOS_COMB.find(gr => gr[2](t))[0];
      soma[x.mes + "|" + g] = (soma[x.mes + "|" + g] || 0) + n0(x.total);
    });
    const rot = meses.map(m => nomeMes(m).split("/")[0].replace(/^./, c => c.toUpperCase()) + (new Set(meses.map(x => x.slice(0, 4))).size > 1 ? "/" + m.slice(2, 4) : ""));
    const usados = GRUPOS_COMB.filter(g => meses.some(m => soma[m + "|" + g[0]]));
    const barras = usados.map(g => ({ type: "bar", label: g[0], data: meses.map(m => Math.round((soma[m + "|" + g[0]] || 0) * 100) / 100), backgroundColor: g[1],
      borderRadius: 12, borderSkipped: false, barPercentage: 0.55, categoryPercentage: 0.7, maxBarThickness: 22, order: 2 }));
    const total = { type: "line", label: "Total", data: meses.map((m, i) => barras.reduce((a, b) => a + b.data[i], 0)), borderColor: "#111", backgroundColor: "#fff",
      borderWidth: 2.5, pointRadius: 5, pointHoverRadius: 7, pointBorderWidth: 2.5, pointBackgroundColor: "#fff", pointBorderColor: "#111", tension: 0, order: 1 };
    $("#ini-graf-cc").textContent = rotulo ? "· " + rotulo : "";
    const real = v => "R$ " + Number(v).toLocaleString("pt-BR", { maximumFractionDigits: 0 });
    const cfg = { data: { labels: rot, datasets: [total].concat(barras) },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: "index", intersect: false },
        plugins: { legend: { position: "bottom", labels: { usePointStyle: true, pointStyle: "circle", color: "#334155", font: { size: 13, weight: "600" }, padding: 16 } },
          tooltip: { callbacks: { label: c => ` ${c.dataset.label}: R$ ${Number(c.raw).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` } } },
        scales: { x: { grid: { display: false }, ticks: { color: "#334155", font: { size: 13, weight: "600" } } },
          y: { beginAtZero: true, ticks: { color: "#64748b", callback: real, maxTicksLimit: 6 }, grid: { color: "#e2e8f0" }, border: { display: false } } } } };
    if (GRAF) GRAF.destroy();
    GRAF = new Chart($("#ini-graf"), cfg);
  }
  $("#ini-graf-caixa").addEventListener("click", e => e.stopPropagation());

  ["#ic-gestor", "#ic-contrato"].forEach(id => $(id).addEventListener("change", () => desenharInicio()));
  $("#ic-cc-txt").addEventListener("click", e => e.stopPropagation());
  $("#ic-cc-txt").addEventListener("input", () => desenharInicioComb());
  $$("#ini-comb [data-ir]").forEach(el => el.addEventListener("click", () => {
    const hoje = new Date(); fcb.mes = hoje.getFullYear() + "-" + String(hoje.getMonth() + 1).padStart(2, "0");
    $("#cb-gestor").value = $("#ic-gestor").value; $("#cb-contrato").value = $("#ic-contrato").value; $("#cb-cc-txt").value = $("#ic-cc-txt").value;
    montarFiltrosComb(); desenharComb(); irPara("combustivel");
  }));

  // ------------------------------------------------------------------ hierarquia (De-Para)
  let HIER = [];
  const fhi = { gerente: "", coord: "", sup: "", busca: "" };
  const podeHier = h => EU && (EU.perfil === "admin" || (EU.perfil === "gerente" && sem(h.gerente) === sem(EU.gestor_operacao)));
  function desenharHier() {
    const uniq = k => Array.from(new Set(HIER.map(h => h[k]).filter(Boolean))).sort((a, b) => sem(a).localeCompare(sem(b)));
    fhi.gerente = opSel("#hi-gerente", uniq("gerente"), "Todos os gerentes");
    fhi.coord = opSel("#hi-coord", uniq("coordenador"), "Todos os coordenadores");
    fhi.sup = opSel("#hi-sup", uniq("supervisor"), "Todos os supervisores");
    $("#dl-hi-sup").innerHTML = uniq("supervisor").map(x => `<option value="${esc(x)}">`).join("");
    $("#dl-hi-coord").innerHTML = uniq("coordenador").map(x => `<option value="${esc(x)}">`).join("");
    $("#dl-hi-ger").innerHTML = Array.from(new Set(uniq("gerente").concat(VEICULOS.map(v => v.gestor_operacao).filter(Boolean)))).sort().map(x => `<option value="${esc(x)}">`).join("");
    const carro = {}; VEICULOS.forEach(v => { if (v.condutor) (carro[sem(v.condutor)] = carro[sem(v.condutor)] || []).push(v.placa); });
    const l = HIER.filter(h => (!fhi.gerente || h.gerente === fhi.gerente) && (!fhi.coord || h.coordenador === fhi.coord) && (!fhi.sup || h.supervisor === fhi.sup)
      && (!fhi.busca || sem([h.colaborador, h.matricula, h.supervisor, h.coordenador, h.gerente, (carro[sem(h.colaborador)] || []).join(" ")].join(" ")).includes(fhi.busca)));
    const comCarro = l.filter(h => carro[sem(h.colaborador)]).length;
    $("#hi-info").textContent = `${l.length} colaboradores · ${comCarro} com carro`;
    $("#tb-hier").innerHTML = l.slice(0, 1500).map(h => `<tr><td><b>${esc(h.colaborador)}</b></td><td>${esc(h.matricula)}</td><td>${esc(h.supervisor)}</td><td>${esc(h.coordenador)}</td><td>${esc(h.gerente)}</td>
      <td>${(carro[sem(h.colaborador)] || []).map(p => `<span class="tag">${esc(p)}</span>`).join(" ") || '<span class="info">—</span>'}</td>
      <td class="info">${esc(h.atualizado_por || "")}${h.atualizado_em ? " · " + dataHora(h.atualizado_em) : ""}</td>
      <td>${podeHier(h) ? `<button class="link" data-hi-ed="${esc(h.colaborador)}">Editar</button>` : ""}</td></tr>`).join("")
      || `<tr><td colspan="8" class="info">${HIER.length ? "Ninguém com esse filtro." : "Nenhuma hierarquia ainda. Suba o De-Para ou adicione os colaboradores."}</td></tr>`;
  }
  [["#hi-gerente", "gerente"], ["#hi-coord", "coord"], ["#hi-sup", "sup"]].forEach(([id, k]) => $(id).addEventListener("change", e => { fhi[k] = e.target.value; desenharHier(); }));
  $("#hi-busca").addEventListener("input", e => { fhi.busca = sem(e.target.value); desenharHier(); });
  async function salvarHier(regs, el) {
    const r = await api("salvar_hierarquia", { regs });
    const x = r.r || {};
    msg(el, `✓ ${x.gravados || 0} colaboradores gravados${x.apagados ? `, ${x.apagados} excluídos` : ""} · ${x.carros || 0} carros com supervisor/coordenador atualizados.`, "ok");
  }
  function editarHier(h) {
    const novo = !h; h = h || { colaborador: "", matricula: "", supervisor: "", coordenador: "", gerente: EU.perfil === "gerente" ? EU.gestor_operacao : "" };
    abrirModal(`<h2>${novo ? "Adicionar colaborador" : esc(h.colaborador)}</h2>
      <div class="campos">
        <label class="larga">Colaborador (nome igual ao do condutor do carro)<input id="hx-col" value="${esc(h.colaborador)}" list="dl-pessoas" ${novo ? "" : "readonly"}></label>
        <label>Matrícula<input id="hx-mat" value="${esc(h.matricula)}"></label>
        <label>Supervisor<input id="hx-sup" value="${esc(h.supervisor)}" list="dl-hi-sup"></label>
        <label>Coordenador<input id="hx-coord" value="${esc(h.coordenador)}" list="dl-hi-coord"></label>
        <label>Gerente<input id="hx-ger" value="${esc(h.gerente)}" list="dl-hi-ger" ${EU.perfil === "gerente" ? "readonly" : ""}></label>
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="hx-ok">Gravar</button>${novo ? "" : '<button class="bt bt-sec" id="hx-del">Excluir</button>'}<span class="msg" id="hx-msg"></span></div>`);
    $("#hx-ok").addEventListener("click", async () => {
      const reg = { colaborador: $("#hx-col").value.trim(), matricula: $("#hx-mat").value.trim(), supervisor: $("#hx-sup").value.trim(), coordenador: $("#hx-coord").value.trim(), gerente: $("#hx-ger").value.trim() };
      if (!reg.colaborador) return msg($("#hx-msg"), "Digite o nome do colaborador.", "erro");
      try { await salvarHier([reg], $("#hi-msg")); fecharModal(); } catch (e) { msg($("#hx-msg"), e.message, "erro"); }
    });
    if (!novo) $("#hx-del").addEventListener("click", async () => {
      if (!window.confirm("Excluir " + h.colaborador + " da hierarquia?")) return;
      try { await salvarHier([{ colaborador: h.colaborador, apagar: true }], $("#hi-msg")); fecharModal(); } catch (e) { msg($("#hx-msg"), e.message, "erro"); }
    });
  }
  $("#bt-hi-novo").addEventListener("click", () => editarHier(null));
  $("#tb-hier").addEventListener("click", e => { const b = e.target.closest("[data-hi-ed]"); if (b) editarHier(HIER.find(h => h.colaborador === b.dataset.hiEd)); });
  // De-Para: colunas Colaborador/Nome, Supervisor, Coordenador, Gerente (Matrícula opcional), cabeçalho em qualquer linha
  $("#arq-hier").addEventListener("change", async e => {
    const arq = e.target.files[0]; e.target.value = "";
    if (!arq) return;
    const el = $("#hi-msg");
    try {
      msg(el, "Lendo " + arq.name + "…");
      const wb = XLSX.read(await arq.arrayBuffer(), { type: "array" });
      const regs = [];
      const nomeCol = h => ["colaborador", "nome", "funcionario", "tecnico", "nome do colaborador", "nome colaborador", "condutor", "colaboradores"].includes(h);
      wb.SheetNames.forEach(nome => {
        const m = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: false });
        const r0 = m.slice(0, 30).findIndex(l => l.some(c => nomeCol(sem(c))) && l.some(c => /^(supervisor|coordenador|gerente|gestor)/.test(sem(c))));
        if (r0 < 0) return;
        const cab = m[r0].map(sem), col = t => cab.findIndex(t);
        const ix = { colaborador: col(nomeCol), matricula: col(h => h.startsWith("matricula") || h === "re"), supervisor: col(h => h.startsWith("supervisor")),
          coordenador: col(h => h.startsWith("coordenador")), gerente: col(h => h.startsWith("gerente") || h.startsWith("gestor")) };
        m.slice(r0 + 1).forEach(l => {
          const t = k => ix[k] >= 0 ? String(l[ix[k]] == null ? "" : l[ix[k]]).replace(/\s+/g, " ").trim() : "";
          if (!t("colaborador")) return;
          const reg = { colaborador: t("colaborador") };
          ["matricula", "supervisor", "coordenador", "gerente"].forEach(k => { if (ix[k] >= 0) reg[k] = t(k); });
          regs.push(reg);
        });
      });
      if (!regs.length) return msg(el, "Não achei as colunas Colaborador (ou Nome) e Supervisor/Coordenador/Gerente na planilha.", "erro");
      const gerentes = Array.from(new Set(regs.map(r => r.gerente).filter(Boolean)));
      if (!window.confirm(`${regs.length} colaboradores na planilha${gerentes.length ? " (gerente: " + gerentes.slice(0, 3).join(", ") + (gerentes.length > 3 ? "…" : "") + ")" : ""}. Gravar na hierarquia?`)) return msg(el, "");
      const tot = { gravados: 0, carros: 0 };
      for (let i = 0; i < regs.length; i += 300) {
        msg(el, `Gravando ${i + 1} a ${Math.min(i + 300, regs.length)} de ${regs.length}…`);
        const r = await api("salvar_hierarquia", { regs: regs.slice(i, i + 300) });
        tot.gravados += (r.r && r.r.gravados) || 0; tot.carros += (r.r && r.r.carros) || 0;
      }
      msg(el, `✓ ${tot.gravados} colaboradores gravados · ${tot.carros} carros com supervisor/coordenador atualizados.`, "ok");
    } catch (err) { msg(el, "Erro ao subir: " + (err.message || err), "erro"); }
  });
  $("#bt-hi-exp").addEventListener("click", () => {
    const carro = {}; VEICULOS.forEach(v => { if (v.condutor) (carro[sem(v.condutor)] = carro[sem(v.condutor)] || []).push(v.placa); });
    const ws = XLSX.utils.json_to_sheet(HIER.map(h => ({ "Colaborador": h.colaborador, "Matrícula": h.matricula, "Supervisor": h.supervisor, "Coordenador": h.coordenador,
      "Gerente": h.gerente, "Carro": (carro[sem(h.colaborador)] || []).join(", "), "Atualizado por": h.atualizado_por })));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Hierarquia"); XLSX.writeFile(wb, "Hierarquia_frota.xlsx");
  });

  // ------------------------------------------------------------------ preencher supervisor e coordenador pelo Organograma
  // fonte 1: o Organograma que o Painel SLA já publica (mesmo endereço do site); fonte 2: a planilha do Organograma
  async function colaboradoresDoPainel() {
    for (const url of ["../Painel-SLA/dados.json", "/Painel-SLA/dados.json"]) {
      try {
        const r = await fetch(url + "?t=" + Date.now(), { cache: "no-store" });
        if (!r.ok) continue;
        const d = await r.json();
        if (d && Array.isArray(d.colaboradores) && d.colaboradores.length) return d.colaboradores;
      } catch (e) { /* tenta a próxima */ }
    }
    return null;
  }
  function colaboradoresDaPlanilha(wb) {
    const out = [];
    const abas = wb.SheetNames.filter(n => /organograma|pj/i.test(n));
    (abas.length ? abas : wb.SheetNames).forEach(nome => {
      const m = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: true });
      const r0 = m.slice(0, 30).findIndex(l => l.some(c => sem(c) === "nome") && l.some(c => sem(c).startsWith("supervisor")));
      if (r0 < 0) return;
      const cab = m[r0].map(sem), col = t => cab.findIndex(t);
      const ix = { nome: col(h => h === "nome"), matricula: col(h => h.startsWith("matricula")), placa: col(h => h.startsWith("placa")),
        supervisor: col(h => h.startsWith("supervisor")), coordenador: col(h => h.startsWith("coordenador")), demissao: col(h => h.includes("demiss")), status: col(h => h === "status" || h.startsWith("situacao")) };
      m.slice(r0 + 1).forEach(l => {
        const t = k => ix[k] >= 0 ? String(l[ix[k]] == null ? "" : l[ix[k]]).trim() : "";
        if (!t("nome")) return;
        out.push({ nome: t("nome"), matricula: t("matricula"), placa: t("placa"), supervisor: t("supervisor"), coordenador: t("coordenador"),
          ativo: !t("demissao") && !/deslig|demit|inativ/i.test(t("status")) });
      });
    });
    return out;
  }
  function casarOrganograma(colabs) {
    const limpaPl = p => String(p || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    const limpaMat = m => String(m || "").replace(/\D/g, "").replace(/^0+/, "");
    const ord = colabs.slice().sort((a, b) => (b.ativo === false ? 0 : 1) - (a.ativo === false ? 0 : 1));   // ativos primeiro
    const porPlaca = new Map(), porMat = new Map(), porNome = new Map();
    ord.forEach(c => {
      if (!c.supervisor && !c.coordenador) return;
      const pl = limpaPl(c.placa), mt = limpaMat(c.matricula), nm = sem(c.nome);
      if (pl.length >= 6 && !porPlaca.has(pl)) porPlaca.set(pl, c);
      if (mt && !porMat.has(mt)) porMat.set(mt, c);
      if (nm && !porNome.has(nm)) porNome.set(nm, c);
    });
    return VEICULOS.map(v => {
      // liga só pelo nome do condutor ou pela matrícula (a placa do Organograma não é usada)
      const c = porNome.get(sem(v.condutor)) || porMat.get(limpaMat(v.matricula_sra));
      if (!c) return null;
      const como = porNome.get(sem(v.condutor)) === c ? "nome do condutor" : "matrícula";
      return { placa: v.placa, condutor: v.condutor, colab: c.nome, como, supAtual: v.supervisor || "", coordAtual: v.coordenador || "",
        supNovo: String(c.supervisor || "").replace(/\s+/g, " ").trim(), coordNovo: String(c.coordenador || "").replace(/\s+/g, " ").trim() };
    }).filter(Boolean);
  }
  function mostrarOrganograma(colabs, fonte) {
    const casados = casarOrganograma(colabs);
    const muda = (x, sub) => sub ? ((x.supNovo && sem(x.supNovo) !== sem(x.supAtual)) || (x.coordNovo && sem(x.coordNovo) !== sem(x.coordAtual)))
      : ((!x.supAtual && x.supNovo) || (!x.coordAtual && x.coordNovo));
    const desenha = () => {
      const sub = $("#org-sub").checked;
      const l = casados.filter(x => muda(x, sub));
      $("#org-res").textContent = `${casados.length} de ${VEICULOS.length} carros encontrados no Organograma · ${l.length} vão mudar`;
      $("#org-tb").innerHTML = l.slice(0, 400).map(x => `<tr><td><b>${esc(x.placa)}</b></td><td>${esc(x.condutor)}<small class="resp">achado por ${x.como}</small></td>
        <td>${esc(x.supAtual || "—")} → <b>${esc(sub || !x.supAtual ? x.supNovo || x.supAtual : x.supAtual)}</b></td>
        <td>${esc(x.coordAtual || "—")} → <b>${esc(sub || !x.coordAtual ? x.coordNovo || x.coordAtual : x.coordAtual)}</b></td></tr>`).join("") || '<tr><td colspan="4" class="info">Nada para mudar.</td></tr>';
      $("#org-gravar").disabled = !l.length;
      return l;
    };
    const semAchar = VEICULOS.length - casados.length;
    abrirModal(`<h2>Preencher pelo Organograma</h2>
      <p class="nota">Fonte: ${esc(fonte)} (${colabs.length} colaboradores). Cada carro é ligado ao colaborador pelo <b>nome do condutor</b> ou pela <b>matrícula</b> e recebe o supervisor e o coordenador dele.${semAchar ? ` ${semAchar} carros não foram encontrados (condutor coringa, uso coletivo ou nome diferente).` : ""}</p>
      <label class="info"><input type="checkbox" id="org-sub"> também trocar quem já tem supervisor/coordenador diferente do Organograma</label>
      <p class="info" id="org-res"></p>
      <div class="tabela-caixa" style="max-height:50vh"><table class="tabela"><thead><tr><th>Placa</th><th>Condutor</th><th>Supervisor</th><th>Coordenador</th></tr></thead><tbody id="org-tb"></tbody></table></div>
      <div class="linha-bt"><button class="bt bt-pri" id="org-gravar">Gravar</button><span class="msg" id="org-msg"></span></div>`);
    desenha();
    $("#org-sub").addEventListener("change", desenha);
    $("#org-gravar").addEventListener("click", async () => {
      const sub = $("#org-sub").checked;
      const regs = desenha().map(x => ({ placa: x.placa, supervisor: sub || !x.supAtual ? x.supNovo : "", coordenador: sub || !x.coordAtual ? x.coordNovo : "" }));
      $("#org-gravar").disabled = true; msg($("#org-msg"), "Gravando " + regs.length + " carros…");
      try { const r = await api("salvar_equipes", { regs }); msg($("#org-msg"), `✓ ${(r.r && r.r.alterados) || 0} carros atualizados.`, "ok"); desenha(); }
      catch (e) { msg($("#org-msg"), e.message, "erro"); $("#org-gravar").disabled = false; }
    });
  }
  $("#bt-org").addEventListener("click", async () => {
    const bt = $("#bt-org"); bt.disabled = true;
    const colabs = await colaboradoresDoPainel();
    bt.disabled = false;
    if (colabs) return mostrarOrganograma(colabs, "Organograma do Painel SLA");
    if (window.confirm("Não consegui ler o Organograma do Painel SLA. Quer escolher a planilha do Organograma?")) $("#arq-org").click();
  });
  $("#arq-org").addEventListener("change", async e => {
    const arq = e.target.files[0]; e.target.value = "";
    if (!arq) return;
    const wb = XLSX.read(await arq.arrayBuffer(), { type: "array" });
    const colabs = colaboradoresDaPlanilha(wb);
    if (!colabs.length) return window.alert("Não achei as colunas Nome e Supervisor nessa planilha.");
    mostrarOrganograma(colabs, arq.name);
  });

  // ------------------------------------------------------------------ exportar Excel
  $("#bt-exportar").addEventListener("click", async () => {
    const bt = $("#bt-exportar"); bt.disabled = true; const t0 = bt.textContent; bt.textContent = "Gerando…";
    try {
      const hist = { data: HIST.slice().sort((a, b) => String(a.placa).localeCompare(String(b.placa)) || String(a.inicio).localeCompare(String(b.inicio))) };
      const oc = { data: OCORRENCIAS };
      const d = x => x ? dataHora(x) : null;   // texto dd/mm/aaaa hh:mm (não muda com o fuso do Excel)
      const veic = filtrados(false).map(v => {
        const o = {};
        CAMPOS.forEach(f => { const x = v[f.c]; o[f.r] = x == null || x === "" ? null : f.t === "date" ? dataBr(x) : x; });
        Object.assign(o, { "Coordenador": v.coordenador, "Supervisor": v.supervisor, "Condutor desde": d(v.condutor_desde),
          "Garagem": v.garagem_tipo, "Endereço da garagem": v.garagem_endereco, "Garagem alterada por": v.garagem_por, "Garagem alterada em": d(v.garagem_em) });
        return o;
      });
      const placas = new Set(filtrados(false).map(v => v.placa));
      const tempo = (hist.data || []).filter(h => placas.has(h.placa)).map(h => ({
        "Placa": h.placa, "Condutor": h.condutor, "Início": d(h.inicio), "Fim": h.fim ? d(h.fim) : "com o carro agora", "Registrado por": h.registrado_por }));
      const ocor = (oc.data || []).filter(o => placas.has(o.placa)).map(o => ({
        "Data": d(o.data_hora), "Placa": o.placa, "Tipo": o.tipo, "Condutor no dia": o.condutor, "Local": o.local, "Descrição": o.descricao,
        "Status": o.status, "Tratativa": o.tratativa, "Registrado por": o.registrado_por, "Tratado por": o.tratado_por }));
      const wb = XLSX.utils.book_new();
      const aba = (linhas, nome, cab) => {
        const ws = XLSX.utils.json_to_sheet(linhas.length ? linhas : [Object.fromEntries(cab.map(c => [c, ""]))], { cellDates: true, dateNF: "dd/mm/yyyy hh:mm" });
        ws["!cols"] = (linhas.length ? Object.keys(linhas[0]) : cab).map(k => ({ wch: Math.max(12, Math.min(45, k.length + 4, ...linhas.slice(0, 300).map(l => String(l[k] == null ? "" : l[k]).length + 2))) }));
        XLSX.utils.book_append_sheet(wb, ws, nome);
      };
      aba(veic, "Veículos", ["Placa"]);
      aba(tempo, "Linha do tempo", ["Placa", "Condutor", "Início", "Fim", "Registrado por"]);
      aba(ocor, "Ocorrências", ["Data", "Placa", "Tipo", "Condutor no dia", "Descrição", "Status"]);
      const hoje = new Date(); const dd = n => String(n).padStart(2, "0");
      XLSX.writeFile(wb, `Frota_${hoje.getFullYear()}-${dd(hoje.getMonth() + 1)}-${dd(hoje.getDate())}.xlsx`);
    } catch (e) { msg($("#imp-msg"), "Erro ao exportar: " + (e.message || e), "erro"); }
    bt.disabled = false; bt.textContent = t0;
  });

  // ------------------------------------------------------------------ ocorrências
  $("#o-tipo").innerHTML = '<option value="">Todos os tipos</option>' + TIPOS_OCOR.map(t => `<option>${t}</option>`).join("");
  [["#o-status", "status"], ["#o-tipo", "tipo"]].forEach(([id, k]) => $(id).addEventListener("change", e => { fofo[k] = e.target.value; desenharOcorrencias(); }));
  $("#o-busca").addEventListener("input", e => { fofo.busca = sem(e.target.value); desenharOcorrencias(); });
  const COR_OCOR = { "Aberta": "#b91c1c", "Em tratativa": "#b45309", "Encerrada": "#15803d" };
  function desenharOcorrencias() {
    const l = OCORRENCIAS.filter(o => (!fofo.placas || fofo.placas.has(o.placa)) && (!fofo.status || o.status === fofo.status) && (!fofo.tipo || o.tipo === fofo.tipo) &&
      (!fofo.busca || sem([o.placa, o.condutor, o.descricao, o.local, o.registrado_por].join(" ")).includes(fofo.busca)));
    $("#o-info").innerHTML = `${l.length} de ${OCORRENCIAS.length} ocorrências` + (fofo.placas ? ` · só ${esc(fofo.rotulo)} <button class="link" id="o-tirar">tirar filtro</button>` : "");
    if (fofo.placas) $("#o-tirar").addEventListener("click", () => { fofo.placas = null; fofo.rotulo = ""; desenharOcorrencias(); });
    $("#tb-ocorrencias").innerHTML = l.length ? l.map(o => `<tr class="clica" data-id="${o.id}">
      <td>${dataHora(o.data_hora)}</td><td><b>${esc(o.placa)}</b></td><td>${esc(o.tipo)}</td><td>${esc(o.condutor) || "não identificado"}</td>
      <td>${esc(o.local) || "—"}</td><td style="white-space:normal;min-width:240px">${esc(o.descricao)}</td>
      <td><span class="st" style="background:${COR_OCOR[o.status] || "#475569"}">${esc(o.status)}</span></td><td>${esc(o.registrado_por) || "—"}</td></tr>`).join("")
      : '<tr><td colspan="8" style="text-align:center;color:#475569;padding:24px">Nenhuma ocorrência.</td></tr>';
  }
  $("#tb-ocorrencias").addEventListener("click", e => { const tr = e.target.closest("tr[data-id]"); if (tr) abrirOcorrencia(Number(tr.dataset.id)); });
  $("#bt-nova-ocor").addEventListener("click", () => novaOcorrencia(""));
  $("#bt-quem").addEventListener("click", async () => {
    const placa = $("#q-placa").value, q = $("#q-data").value;
    if (!placa || !q) return;
    const data = quemLocal(placa, new Date(q));
    $("#q-resp").innerHTML = data ? `👤 <b>${esc(data)}</b> estava com o ${esc(placa)} em ${dataHora(q)}` : `Nenhum condutor registrado para o ${esc(placa)} nessa data.`;
  });

  function novaOcorrencia(placa) {
    abrirModal(`<h2>Registrar ocorrência</h2>
      <div class="campos">
        <label>Placa<select id="n-placa">${VEICULOS.map(v => `<option ${v.placa === placa ? "selected" : ""}>${esc(v.placa)}</option>`).join("")}</select></label>
        <label>Data e hora<input id="n-data" type="datetime-local" value="${paraInput(new Date())}"></label>
        <label>Tipo<select id="n-tipo">${TIPOS_OCOR.map(t => `<option>${t}</option>`).join("")}</select></label>
        <label>Local<input id="n-local" placeholder="Endereço ou referência"></label>
        <label class="larga">Condutor no dia<input id="n-condutor" placeholder="vem sozinho pela linha do tempo"></label>
        <textarea id="n-desc" placeholder="O que aconteceu? (obrigatório)"></textarea>
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="n-salva">Registrar</button><span id="n-msg" class="msg"></span></div>`);
    const achar = async () => {
      const data = quemLocal($("#n-placa").value, doCampo($("#n-data").value));
      $("#n-condutor").value = data || "";
      $("#n-condutor").placeholder = data ? "" : "nenhum condutor registrado nessa data — digite se souber";
    };
    $("#n-placa").addEventListener("change", achar); $("#n-data").addEventListener("change", achar); achar();
    $("#n-salva").addEventListener("click", async () => {
      const desc = $("#n-desc").value.trim();
      if (!desc) return msg($("#n-msg"), "Descreva o que aconteceu.", "erro");
      msg($("#n-msg"), "Gravando…");
      try {
        await api("nova_ocorrencia", { dados: {
          placa: $("#n-placa").value, data_hora: doCampo($("#n-data").value).toISOString(), tipo: $("#n-tipo").value,
          local: $("#n-local").value.trim(), condutor: $("#n-condutor").value.trim(), descricao: desc } });
      } catch (err) { return msg($("#n-msg"), err.message, "erro"); }
      fecharModal();
      irPara("ocorrencias");
    });
  }

  function abrirOcorrencia(id) {
    const o = OCORRENCIAS.find(x => x.id === id); if (!o) return;
    abrirModal(`<h2>${esc(o.tipo)} · ${esc(o.placa)}</h2>
      <div class="dados">
        <div><small>Data</small>${dataHora(o.data_hora)}</div><div><small>Condutor no dia</small>${esc(o.condutor) || "não identificado"}</div>
        <div><small>Local</small>${esc(o.local) || "—"}</div><div><small>Registrado por</small>${esc(o.registrado_por) || "—"} · ${dataHora(o.registrado_em)}</div>
      </div>
      <p style="white-space:pre-wrap">${esc(o.descricao)}</p>
      <h3>Tratativa</h3>
      <div class="campos">
        <label>Status<select id="t-status">${["Aberta", "Em tratativa", "Encerrada"].map(s => `<option ${o.status === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>
        <textarea id="t-texto" placeholder="O que foi feito (desconto, advertência, conserto…)">${esc(o.tratativa || "")}</textarea>
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="t-salva">Salvar tratativa</button><span id="t-msg" class="msg"></span>
        ${o.tratado_por ? `<span class="info">última tratativa: ${esc(o.tratado_por)} · ${dataHora(o.tratado_em)}</span>` : ""}</div>`);
    $("#t-salva").addEventListener("click", async () => {
      msg($("#t-msg"), "Gravando…");
      try { await api("tratar_ocorrencia", { id: o.id, status: $("#t-status").value, tratativa: $("#t-texto").value.trim() }); fecharModal(); }
      catch (err) { msg($("#t-msg"), err.message, "erro"); }
    });
  }

  // ------------------------------------------------------------------ importar planilha (admin)
  // Lê a aba "Base ativa" (ou a primeira aba que tiver a coluna "Placa"), cabeçalho em qualquer linha.
  const MAPA = CAMPOS.map(f => [f.c, f.teste, f.t]).concat(EXTRA_IMPORT.map(f => [f.c, f.teste]));
  // converte a célula da planilha conforme o tipo da coluna
  const converter = (tipo, v) => {
    if (v == null || v === "") return null;
    const txt = String(v).trim();
    if (!txt || txt === "-" || txt === "—") return null;
    if (tipo === "num") return numero(v);
    if (tipo === "int") { const n = parseInt(txt, 10); return isFinite(n) ? n : null; }
    if (tipo === "date") {
      if (typeof v === "number" && v > 20000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 86400000)); return d.toISOString().slice(0, 10); }
      const m = txt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (m) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
      const iso = txt.match(/^(\d{4}-\d{2}-\d{2})/); return iso ? iso[1] : null;
    }
    return txt;
  };
  const numero = v => {
    if (typeof v === "number") return v;
    const t = String(v == null ? "" : v).replace(/[R$\s]/g, "");
    if (!t) return null;
    const n = Number(t.includes(",") ? t.replace(/\./g, "").replace(",", ".") : t);
    return isFinite(n) ? n : null;
  };

  $("#arq-frota").addEventListener("change", async e => {
    const arq = e.target.files[0]; e.target.value = "";
    if (!arq) return;
    const el = $("#imp-msg");
    if (VEICULOS.length && !window.confirm("A carga inicial já foi feita (" + VEICULOS.length + " veículos no sistema). Importar a planilha de novo vai atualizar os carros com os dados dela. Continuar?")) return;
    try {
      msg(el, "Lendo " + arq.name + "…");
      const wb = XLSX.read(await arq.arrayBuffer(), { type: "array", cellDates: false });
      const ordem = wb.SheetNames.slice().sort((a, b) => (sem(b) === "base ativa") - (sem(a) === "base ativa"));
      let linhas = null, cab = null, aba = "";
      for (const nome of ordem) {
        const m = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: true });
        const r = m.slice(0, 40).findIndex(l => l.some(c => sem(c) === "placa"));
        if (r >= 0) { cab = m[r].map(sem); linhas = m.slice(r + 1); aba = nome; break; }
      }
      if (!linhas) return msg(el, "Não achei nenhuma aba com a coluna Placa.", "erro");
      const idx = {};
      // cada coluna da planilha vai para um campo só (a primeira regra que reconhecer o título)
      const usadas = new Set();
      MAPA.forEach(([campo, teste]) => { const i = cab.findIndex((h, j) => !usadas.has(j) && teste(h)); if (i >= 0) { idx[campo] = i; usadas.add(i); } });
      const porPlaca = new Map();
      linhas.forEach(l => {
        const placa = String(l[idx.placa] || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
        if (placa.length < 6) return;
        const reg = { placa, atualizado_em: new Date().toISOString() };
        MAPA.forEach(([campo, , tipo]) => {
          if (campo === "placa" || idx[campo] === undefined) return;
          reg[campo] = converter(tipo, l[idx[campo]]);
        });
        porPlaca.set(placa, reg);
      });
      const regs = Array.from(porPlaca.values());
      if (!regs.length) return msg(el, "A aba " + aba + " não tem placas preenchidas.", "erro");
      // condutor: a planilha da Tamires é a base principal. Quando o condutor muda, a troca
      // entra sozinha na linha do tempo (data da importação). Célula de condutor vazia não apaga quem está no sistema.
      const existentes = new Map(VEICULOS.map(v => [v.placa, v]));
      let trocas = 0, novos = 0;
      regs.forEach(r => {
        const ex = existentes.get(r.placa);
        if (!ex) { novos++; return; }
        if ("condutor" in r && !r.condutor) { delete r.condutor; return; }
        if ("condutor" in r && sem(r.condutor) !== sem(ex.condutor)) trocas++;
        else delete r.condutor;   // mesmo condutor: não mexe na linha do tempo
      });
      msg(el, `Enviando ${regs.length} veículos da aba "${aba}"… (pode levar até 1 minuto)`);
      // a planilha só preenche (célula vazia não apaga o que já está no sistema); troca de condutor entra na linha do tempo
      for (let i = 0; i < regs.length; i += 100) {
        msg(el, `Enviando veículos ${i + 1} a ${Math.min(i + 100, regs.length)} de ${regs.length} da aba "${aba}"…`);
        await api("carga_inicial", { regs: regs.slice(i, i + 100) });
      }
      const faltou = CAMPOS.filter(f => idx[f.c] === undefined).map(f => f.r);
      msg(el, `✓ ${regs.length} veículos importados da aba "${aba}".` + (novos ? ` ${novos} ${novos === 1 ? "carro novo" : "carros novos"}.` : "") + (trocas ? ` ${trocas} ${trocas === 1 ? "troca" : "trocas"} de condutor ${trocas === 1 ? "registrada" : "registradas"} na linha do tempo.` : " Nenhuma troca de condutor.") + (faltou.length ? " Colunas que não estavam na planilha: " + faltou.join(", ") + "." : " Todas as colunas encontradas."), "ok");
    } catch (err) {
      msg(el, "Erro ao importar: " + (err.message || err), "erro");
    }
  });

  // ------------------------------------------------------------------ usuários (admin)
  function desenharUsuarios() {
    const data = USUARIOS.slice().sort((a, b) => String(a.nome).localeCompare(String(b.nome)));
    $("#tb-usuarios").innerHTML = data.map(u => `<tr>
      <td><b>${esc(u.nome)}</b></td><td>${esc(u.email)}</td>
      <td>${esc({ admin: "Administrador", gerente: "Gerente", coordenador: "Coordenador", supervisor: "Supervisor" }[u.perfil] || u.perfil)}</td>
      <td>${esc(u.perfil === "admin" ? "Todos" : (u.perfil === "gerente" ? u.gestor_operacao : u.perfil === "coordenador" ? u.coordenador : u.supervisor)) || "—"}</td>
      <td>${u.perfil === "admin" || !listaAcessos(u).length ? '<span class="tag-ac">Tudo</span>' : listaAcessos(u).map(k => `<span class="tag-ac">${esc(ACESSOS[k] || k)}</span>`).join("")}</td>
      <td>${u.ativo ? "Sim" : "Não"}${u.tem_senha ? "" : ' <small style="color:#b45309">(sem senha ainda)</small>'}</td>
      <td><button class="bt bt-sec" data-acao="editar" data-email="${esc(u.email)}">Editar</button> ${String(u.email).toLowerCase() === String(EU.email).toLowerCase() ? "" : `<button class="bt bt-sec" data-acao="ativar" data-email="${esc(u.email)}" data-ativo="${u.ativo}">${u.ativo ? "Desativar" : "Reativar"}</button>
        <button class="bt bt-sec" data-acao="zerar" data-email="${esc(u.email)}" title="A pessoa cria uma senha nova em Primeiro acesso">Zerar senha</button>`}</td></tr>`).join("");
    $$("#tb-usuarios button").forEach(b => b.addEventListener("click", async () => {
      try {
        if (b.dataset.acao === "editar") {
          const u = USUARIOS.find(x => x.email === b.dataset.email); if (!u) return;
          $("#u-email").value = u.email; $("#u-nome").value = u.nome; $("#u-perfil").value = u.perfil;
          $("#u-equipe").value = u.perfil === "gerente" ? u.gestor_operacao : u.perfil === "coordenador" ? u.coordenador : u.perfil === "supervisor" ? u.supervisor : "";
          const ac = listaAcessos(u);
          $$("#u-acessos input").forEach(x => { x.checked = !ac.length || ac.indexOf(x.value) >= 0; });
          $("#u-salvar").textContent = "Salvar alterações"; $("#form-user").scrollIntoView({ behavior: "smooth" });
          return;
        }
        if (b.dataset.acao === "zerar") {
          if (!window.confirm("Zerar a senha de " + b.dataset.email + "? A pessoa vai criar outra em Primeiro acesso.")) return;
          await api("zerar_senha", { dados: { email: b.dataset.email } });
          msg($("#u-msg"), "✓ Senha zerada. A pessoa entra em Primeiro acesso e cria uma nova.", "ok");
        } else await api("ativar_usuario", { dados: { email: b.dataset.email }, ativo: b.dataset.ativo !== "true" });
      } catch (err) { msg($("#u-msg"), err.message, "erro"); }
    }));
  }
  $("#u-ac-todos").addEventListener("click", () => $$("#u-acessos input").forEach(x => { x.checked = true; }));
  $("#u-ac-nada").addEventListener("click", () => $$("#u-acessos input").forEach(x => { x.checked = false; }));
  $("#u-perfil").addEventListener("change", () => { $("#u-acessos").disabled = $("#u-perfil").value === "admin"; });
  $("#form-user").addEventListener("submit", async e => {
    e.preventDefault();
    const perfil = $("#u-perfil").value, equipe = $("#u-equipe").value.trim();
    if (perfil !== "admin" && !equipe) return msg($("#u-msg"), "Informe quem essa pessoa vai ver: o Gestor operação (gerente), o Coordenador (coordenador) ou o Supervisor (supervisor).", "erro");
    const marcados = $$("#u-acessos input").filter(x => x.checked).map(x => x.value);
    const todos = $$("#u-acessos input").length;
    const reg = { email: $("#u-email").value.trim().toLowerCase(), nome: $("#u-nome").value.trim(), perfil, equipe,
      acessos: perfil === "admin" || marcados.length === todos ? "" : marcados.join(",") };
    if (perfil !== "admin" && !marcados.length) return msg($("#u-msg"), "Marque pelo menos uma coisa que a pessoa pode acessar.", "erro");
    try { await api("salvar_usuario", { dados: reg }); }
    catch (err) { return msg($("#u-msg"), err.message, "erro"); }
    msg($("#u-msg"), `✓ ${reg.nome} salvo. Agora ele entra pela tela de login em "Primeiro acesso".`, "ok");
    e.target.reset(); $$("#u-acessos input").forEach(x => { x.checked = true; }); $("#u-salvar").textContent = "Salvar";
  });
})();
