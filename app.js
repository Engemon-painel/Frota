// =====================================================================
// GESTÃO DE FROTAS — telas (login, Veículos, Usuários)
// Dados e permissões ficam no Supabase: cada login só recebe os carros
// da própria equipe (a trava está no banco, não nesta tela).
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
  const configurado = /^https:\/\/.+\.supabase\.co/.test(CFG.SUPABASE_URL || "") && (CFG.SUPABASE_ANON_KEY || "").length > 20;
  const sb = configurado ? window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY) : null;

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
  const fofo = { status: "", tipo: "", busca: "" };
  const TIPOS_OCOR = ["Batida", "Avaria", "Multa", "Roubo/Furto", "Pane mecânica", "Desvio de rota", "Uso fora do horário", "Uso indevido", "Abastecimento irregular", "Outro"];
  const GARAGENS = ["Própria", "Alugada", "Casa do condutor", "Cliente"];
  const dataHora = v => v ? new Date(v).toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "—";
  const paraInput = d => { const x = new Date(d.getTime() - d.getTimezoneOffset() * 60000); return x.toISOString().slice(0, 16); };
  const podeGerir = () => EU && (EU.perfil === "admin" || EU.perfil === "gerente");
  const podeSup = () => EU && EU.perfil !== "supervisor";      // supervisor do carro: gerente, coordenador ou admin

  // ------------------------------------------------------------------ login
  function mostrarLogin(texto, tipo) {
    $("#tela-app").hidden = true; $("#tela-login").hidden = false;
    msg($("#login-msg"), texto, tipo);
  }
  if (!configurado) {
    mostrarLogin("Falta configurar o Supabase no arquivo config.js.", "erro");
    $("#bt-entrar").disabled = true;
    return;
  }

  $("#form-login").addEventListener("submit", async e => {
    e.preventDefault();
    const email = $("#login-email").value.trim(), senha = $("#login-senha").value;
    msg($("#login-msg"), "Entrando…");
    const { error } = await sb.auth.signInWithPassword({ email, password: senha });
    if (error) msg($("#login-msg"), /invalid/i.test(error.message) ? "E-mail ou senha incorretos." : error.message, "erro");
  });
  $("#bt-primeiro").addEventListener("click", async () => {
    const email = $("#login-email").value.trim(), senha = $("#login-senha").value;
    if (!email || senha.length < 6) return msg($("#login-msg"), "Digite o e-mail e a senha que quer usar (mínimo 6 caracteres) e clique de novo.", "erro");
    const { data, error } = await sb.auth.signUp({ email, password: senha });
    if (error) return msg($("#login-msg"), /registered/i.test(error.message) ? "Este e-mail já tem senha. Use Entrar ou Esqueci a senha." : error.message, "erro");
    msg($("#login-msg"), data.session ? "Senha criada. Entrando…" : "Senha criada. Confirme pelo link enviado ao seu e-mail e depois clique em Entrar.", "ok");
  });
  $("#bt-esqueci").addEventListener("click", async () => {
    const email = $("#login-email").value.trim();
    if (!email) return msg($("#login-msg"), "Digite o seu e-mail e clique de novo em Esqueci a senha.", "erro");
    const { error } = await sb.auth.resetPasswordForEmail(email, { redirectTo: location.origin + location.pathname });
    msg($("#login-msg"), error ? error.message : "Enviamos um link para o seu e-mail para criar uma senha nova.", error ? "erro" : "ok");
  });
  $("#bt-sair").addEventListener("click", () => sb.auth.signOut());

  sb.auth.onAuthStateChange(async (evento, sessao) => {
    if (evento === "PASSWORD_RECOVERY") {
      const nova = window.prompt("Digite a nova senha (mínimo 6 caracteres):");
      if (nova && nova.length >= 6) {
        const { error } = await sb.auth.updateUser({ password: nova });
        window.alert(error ? "Não foi possível trocar a senha: " + error.message : "Senha trocada.");
      }
    }
    if (sessao) entrar(sessao); else { entrando = ""; EU = null; mostrarLogin(""); }
  });

  let entrando = "";
  async function entrar(sessao) {
    const email = String(sessao.user.email || "").toLowerCase();
    if (entrando === email) return;      // o Supabase avisa o mesmo login mais de uma vez
    entrando = email;
    const { data, error } = await sb.from("usuarios").select("*").ilike("email", email).maybeSingle();
    if (error || !data || !data.ativo) {
      entrando = "";
      await sb.auth.signOut();
      return mostrarLogin("Seu e-mail ainda não tem acesso. Peça ao administrador para cadastrar você em Usuários.", "erro");
    }
    EU = data;
    $("#user-nome").textContent = EU.nome;
    $("#user-perfil").textContent = { admin: "Administrador", gerente: "Gerente", coordenador: "Coordenador", supervisor: "Supervisor" }[EU.perfil] || EU.perfil;
    $$("[data-so-admin]").forEach(el => { el.hidden = EU.perfil !== "admin"; });
    $("#tela-login").hidden = true; $("#tela-app").hidden = false;
    await carregarVeiculos();
    carregarOcorrencias();
    carregarHistorico();
    if (EU.perfil === "admin") carregarUsuarios();
  }

  // ------------------------------------------------------------------ abas
  $$(".aba").forEach(b => b.addEventListener("click", () => {
    $$(".aba").forEach(x => x.classList.toggle("on", x === b));
    $$(".aba-conteudo").forEach(c => { c.hidden = c.id !== "aba-" + b.dataset.aba; });
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

  async function carregarVeiculos() {
    const { data, error } = await sb.from("veiculos").select("*").order("placa").range(0, 9999);
    if (error) { msg($("#imp-msg"), "Erro ao carregar veículos: " + error.message, "erro"); return; }
    VEICULOS = data || [];
    montarFiltros();
    desenhar();
  }

  function opcoes(sel, valores, primeiro) {
    const atual = sel.value;
    sel.innerHTML = `<option value="">${esc(primeiro)}</option>` + valores.map(v => `<option>${esc(v)}</option>`).join("");
    sel.value = valores.includes(atual) ? atual : "";
  }
  const unicos = campo => Array.from(new Set(VEICULOS.map(v => (v[campo] || "").trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, "pt-BR"));
  function montarFiltros() {
    opcoes($("#f-contrato"), unicos("operacao"), "Todos os contratos");
    opcoes($("#f-gestor"), unicos("gestor_operacao"), "Todos os gestores");
    opcoes($("#f-status"), unicos("status"), "Todos os status");
    const dl = (id, campos, extra) => { $(id).innerHTML = Array.from(new Set([].concat(...campos.map(unicos), extra || []))).sort((a, b) => a.localeCompare(b, "pt-BR")).map(v => `<option value="${esc(v)}">`).join(""); };
    dl("#dl-operacao", ["operacao"]); dl("#dl-gestor", ["gestor_operacao"]);
    dl("#dl-status", ["status"], ["Ativo", "Manutenção", "Funilaria", "Disponível", "Devolvido"]);
    dl("#dl-tipo", ["tipo"]); dl("#dl-comb", ["combustivel"], ["Flex", "Diesel", "Gasolina", "Etanol", "Elétrico"]);
    dl("#dl-pessoas", ["supervisor", "coordenador", "gestor_operacao"]);
    $("#lista-equipes").innerHTML = Array.from(new Set(unicos("gestor_operacao").concat(unicos("coordenador"), unicos("supervisor")))).map(v => `<option value="${esc(v)}">`).join("");
  }
  [["#f-contrato", "contrato"], ["#f-gestor", "gestor"], ["#f-status", "status"]].forEach(([id, k]) =>
    $(id).addEventListener("change", e => { filtro[k] = e.target.value; desenhar(); }));
  $("#f-busca").addEventListener("input", e => { filtro.busca = sem(e.target.value); desenhar(); });
  $("#f-garagem").addEventListener("change", e => { filtro.garagem = e.target.value; desenhar(); });

  function filtrados(ignorarStatus) {
    return VEICULOS.filter(v =>
      (!filtro.contrato || v.operacao === filtro.contrato) &&
      (!filtro.gestor || v.gestor_operacao === filtro.gestor) &&
      (ignorarStatus || !filtro.status || v.status === filtro.status) &&
      (!filtro.garagem || (filtro.garagem === "sem" ? !v.garagem_tipo : v.garagem_tipo === filtro.garagem)) &&
      (!filtro.busca || sem([v.placa, v.condutor, v.modelo, v.marca, v.operacao, v.gestor_operacao, v.supervisor, v.coordenador, v.garagem_endereco].join(" ")).includes(filtro.busca)));
  }

  function desenhar() {
    const base = filtrados(true);
    const porStatus = {};
    base.forEach(v => { const s = v.status || "Sem status"; porStatus[s] = (porStatus[s] || 0) + 1; });
    const kp = [{ r: "Total de veículos", v: base.length, cor: "#0e7c86", s: "" }]
      .concat(Object.keys(porStatus).sort((a, b) => porStatus[b] - porStatus[a]).map(s => ({ r: s, v: porStatus[s], cor: COR_STATUS(s), s })));
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
        if (f.c === "status") return `<td>${v.status ? `<span class="st" style="background:${COR_STATUS(v.status)}">${esc(v.status)}</span>` : "—"}</td>`;
        return `<td${f.t === "num" ? ' class="num"' : ""}>${mostrar(f, v[f.c])}</td>`;
      }).join("") +
      `<td>${esc(v.coordenador) || '<span style="color:#b45309;font-weight:700">preencher</span>'}</td>
      <td>${esc(v.supervisor) || '<span style="color:#b45309;font-weight:700">preencher</span>'}</td>
      <td>${v.garagem_tipo ? `${esc(v.garagem_tipo)}<small style="display:block;color:#475569">${esc(v.garagem_endereco || "")}</small>` : '<span style="color:#b91c1c;font-weight:800">⚠ informar</span>'}</td></tr>`).join("")
      : `<tr><td colspan="${CAMPOS.length + 3}" style="text-align:center;color:#475569;padding:24px">${VEICULOS.length ? "Nenhum veículo com esses filtros." : (EU && EU.perfil === "admin" ? "Nenhum veículo ainda. Clique em Carga inicial (planilha) para trazer a base de hoje, ou em + Nova placa." : "Nenhum veículo da sua equipe cadastrado ainda.")}</td></tr>`;
  }

  $("#tb-veiculos").addEventListener("click", e => { const tr = e.target.closest("tr[data-placa]"); if (tr) abrirVeiculo(tr.dataset.placa); });

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
        <div><small>Status</small>${esc(v.status) || "—"}</div><div><small>Tipo</small>${esc(v.tipo) || "—"}</div>
      </div>
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
      const { error } = await sb.from("veiculos").update(r).eq("placa", v.placa);
      if (error) return msg($("#m-msg0"), error.message, "erro");
      msg($("#m-msg0"), "✓ Salvo", "ok"); await carregarVeiculos();
    });
    $("#m-salva-condutor").addEventListener("click", async () => {
      const nome = $("#m-condutor").value.trim(), desde = $("#m-desde").value;
      if (!nome) return msg($("#m-msg1"), "Digite o nome do novo condutor.", "erro");
      const { error } = await sb.rpc("trocar_condutor", { p_placa: v.placa, p_condutor: nome, p_desde: new Date(desde || Date.now()).toISOString() });
      if (error) return msg($("#m-msg1"), error.message, "erro");
      await carregarVeiculos(); carregarHistorico(); abrirVeiculo(v.placa);
    });
    if (podeSup()) $("#m-salva-equipe").addEventListener("click", async () => {
      const mud = { supervisor: $("#m-sup").value.trim() || null };
      if (gerir) mud.coordenador = $("#m-coord").value.trim() || null;
      const { error } = await sb.from("veiculos").update(mud).eq("placa", v.placa);
      if (error) return msg($("#m-msg2"), error.message, "erro");
      msg($("#m-msg2"), "✓ Salvo", "ok"); await carregarVeiculos();
    });
    $("#m-salva-gar").addEventListener("click", async () => {
      const tipo = $("#m-gar").value, end = $("#m-end").value.trim();
      if (!tipo || !end) return msg($("#m-msg3"), "Escolha a garagem e escreva o endereço.", "erro");
      const { error } = await sb.from("veiculos").update({ garagem_tipo: tipo, garagem_endereco: end }).eq("placa", v.placa);
      if (error) return msg($("#m-msg3"), error.message, "erro");
      msg($("#m-msg3"), "✓ Salvo", "ok"); await carregarVeiculos();
    });
    $("#m-nova-ocor").addEventListener("click", () => novaOcorrencia(v.placa));

    const [hist, oc] = await Promise.all([
      sb.from("historico_condutor").select("*").eq("placa", v.placa).order("inicio", { ascending: false }),
      sb.from("ocorrencias").select("*").eq("placa", v.placa).order("data_hora", { ascending: false })
    ]);
    if (!$("#m-tempo")) return;
    const h = hist.data || [];
    $("#m-tempo").innerHTML = h.length ? h.map(x => `<li><b>${esc(x.condutor)}</b><span>${dataHora(x.inicio)} → ${x.fim ? dataHora(x.fim) : '<span class="atual">com o carro agora</span>'}</span><span class="info">registrado por ${esc(x.registrado_por || "—")}</span></li>`).join("")
      : "<li>Nenhuma troca registrada ainda.</li>";
    const o = oc.data || [];
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
      const { error } = await sb.from("veiculos").insert(r);
      if (error) return msg($("#cv-msg"), error.message, "erro");
      await carregarVeiculos(); carregarHistorico(); abrirVeiculo(r.placa);
    });
  });

  // ------------------------------------------------------------------ histórico de uso (linha do tempo de todos os carros)
  let HIST = [];
  const fh = { busca: "", de: "", ate: "" };
  async function carregarHistorico() {
    const { data, error } = await sb.from("historico_condutor").select("*").order("inicio", { ascending: false }).range(0, 49999);
    if (error) { $("#h-info").textContent = "Erro: " + error.message; return; }
    HIST = data || [];
    const placas = VEICULOS.map(v => `<option>${esc(v.placa)}</option>`).join("");
    $("#h-q-placa").innerHTML = placas;
    if (!$("#h-q-data").value) $("#h-q-data").value = paraInput(new Date());
    $("#dl-pessoas").innerHTML += Array.from(new Set(HIST.map(h => h.condutor))).map(c => `<option value="${esc(c)}">`).join("");
    desenharHistorico();
  }
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
      </div>
      <div class="linha-bt"><button class="bt bt-pri" id="tc-salva">Registrar troca</button><span id="tc-msg" class="msg"></span></div>`);
    $("#tc-salva").addEventListener("click", async () => {
      const nome = $("#tc-nome").value.trim(); if (!nome) return msg($("#tc-msg"), "Digite o nome do novo condutor.", "erro");
      const { error } = await sb.rpc("trocar_condutor", { p_placa: $("#tc-placa").value, p_condutor: nome, p_desde: new Date($("#tc-desde").value || Date.now()).toISOString() });
      if (error) return msg($("#tc-msg"), error.message, "erro");
      fecharModal(); await carregarVeiculos(); carregarHistorico();
    });
  });

  // ------------------------------------------------------------------ multas e batidas: cruza a base com o histórico
  let CZ = { cab: [], linhas: [], iPlaca: -1, quando: [] };
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
      if (!HIST.length) await carregarHistorico();
      const wb = XLSX.read(await arq.arrayBuffer(), { type: "array", cellDates: true });
      let m = null, r = -1;
      for (const nome of wb.SheetNames) {
        const mm = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: true });
        const rr = mm.slice(0, 40).findIndex(l => l.some(c => sem(c) === "placa") && l.some(c => sem(c).startsWith("data")));
        if (rr >= 0) { m = mm; r = rr; break; }
      }
      if (!m) return msg(el, "Não achei uma aba com as colunas Placa e Data.", "erro");
      const cab = m[r].map(c => String(c).trim()), cs = cab.map(sem);
      const iPlaca = cs.indexOf("placa");
      const iData = cs.findIndex(c => c === "data" || c.startsWith("data da") || c.startsWith("data do") || c.startsWith("data/hora") || c.startsWith("data e hora") || c.startsWith("data"));
      const iHora = cs.findIndex(c => c === "hora" || c.startsWith("hora da") || c.startsWith("hora do"));
      const linhas = m.slice(r + 1).filter(l => String(l[iPlaca] || "").trim());
      let achou = 0, semData = 0;
      const quando = [], cond = [];
      linhas.forEach(l => {
        const placa = String(l[iPlaca]).toUpperCase().replace(/[^A-Z0-9]/g, "");
        const d = dataDe(l[iData], iHora >= 0 ? l[iHora] : null);
        quando.push(d);
        const c = d ? quemLocal(placa, d) : "";
        if (!d) semData++; else if (c) achou++;
        cond.push(c);
      });
      CZ = { cab, linhas, iPlaca, iData, quando, cond };
      $("#cz-cab").innerHTML = "<tr>" + cab.map(c => `<th>${esc(c)}</th>`).join("") + '<th style="background:#dcfce7">Condutor no dia</th></tr>';
      const fmt = v => v instanceof Date ? v.toLocaleString("pt-BR") : esc(v);
      $("#tb-cruzar").innerHTML = linhas.map((l, i) => "<tr>" + cab.map((_, j) => `<td>${fmt(l[j])}</td>`).join("") +
        `<td style="font-weight:800;background:#f0fdf4">${cond[i] ? esc(cond[i]) : '<span style="color:#b91c1c">não identificado</span>'}</td></tr>`).join("");
      msg(el, `✓ ${linhas.length} linhas · condutor identificado em ${achou}` + (linhas.length - achou - semData ? ` · ${linhas.length - achou - semData} sem condutor registrado na data (ou placa fora da sua equipe)` : "") + (semData ? ` · ${semData} sem data válida` : ""), "ok");
      ["#bt-cruzar-exp", "#bt-cruzar-ocor", "#cz-tipo"].forEach(id => { $(id).hidden = false; });
    } catch (err) { msg(el, "Erro ao ler a base: " + (err.message || err), "erro"); }
  });
  $("#bt-cruzar-exp").addEventListener("click", () => {
    const aoa = [CZ.cab.concat("Condutor no dia")].concat(CZ.linhas.map((l, i) => CZ.cab.map((_, j) => l[j]).concat(CZ.cond[i] || "não identificado")));
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
    for (let i = 0; i < regs.length; i += 200) {
      const { error } = await sb.from("ocorrencias").insert(regs.slice(i, i + 200));
      if (error) return msg($("#cz-msg"), "Erro: " + error.message, "erro");
    }
    msg($("#cz-msg"), `✓ ${regs.length} ocorrências registradas.`, "ok"); carregarOcorrencias();
  });

  // ------------------------------------------------------------------ exportar Excel
  $("#bt-exportar").addEventListener("click", async () => {
    const bt = $("#bt-exportar"); bt.disabled = true; const t0 = bt.textContent; bt.textContent = "Gerando…";
    try {
      const [hist, oc] = await Promise.all([
        sb.from("historico_condutor").select("*").order("placa").order("inicio").range(0, 49999),
        sb.from("ocorrencias").select("*").order("data_hora", { ascending: false }).range(0, 49999)
      ]);
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
  async function carregarOcorrencias() {
    const { data, error } = await sb.from("ocorrencias").select("*").order("data_hora", { ascending: false }).range(0, 4999);
    if (error) { $("#o-info").textContent = "Erro: " + error.message; return; }
    OCORRENCIAS = data || [];
    $("#q-placa").innerHTML = VEICULOS.map(v => `<option>${esc(v.placa)}</option>`).join("");
    if (!$("#q-data").value) $("#q-data").value = paraInput(new Date());
    desenharOcorrencias();
  }
  [["#o-status", "status"], ["#o-tipo", "tipo"]].forEach(([id, k]) => $(id).addEventListener("change", e => { fofo[k] = e.target.value; desenharOcorrencias(); }));
  $("#o-busca").addEventListener("input", e => { fofo.busca = sem(e.target.value); desenharOcorrencias(); });
  const COR_OCOR = { "Aberta": "#b91c1c", "Em tratativa": "#b45309", "Encerrada": "#15803d" };
  function desenharOcorrencias() {
    const l = OCORRENCIAS.filter(o => (!fofo.status || o.status === fofo.status) && (!fofo.tipo || o.tipo === fofo.tipo) &&
      (!fofo.busca || sem([o.placa, o.condutor, o.descricao, o.local, o.registrado_por].join(" ")).includes(fofo.busca)));
    $("#o-info").textContent = `${l.length} de ${OCORRENCIAS.length} ocorrências`;
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
    const { data, error } = await sb.rpc("quem_estava", { p_placa: placa, p_quando: new Date(q).toISOString() });
    $("#q-resp").innerHTML = error ? esc(error.message) : (data ? `👤 <b>${esc(data)}</b> estava com o ${esc(placa)} em ${dataHora(q)}` : `Nenhum condutor registrado para o ${esc(placa)} nessa data.`);
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
      const { data } = await sb.rpc("quem_estava", { p_placa: $("#n-placa").value, p_quando: new Date($("#n-data").value || Date.now()).toISOString() });
      $("#n-condutor").value = data || "";
      $("#n-condutor").placeholder = data ? "" : "nenhum condutor registrado nessa data — digite se souber";
    };
    $("#n-placa").addEventListener("change", achar); $("#n-data").addEventListener("change", achar); achar();
    $("#n-salva").addEventListener("click", async () => {
      const desc = $("#n-desc").value.trim();
      if (!desc) return msg($("#n-msg"), "Descreva o que aconteceu.", "erro");
      const { error } = await sb.from("ocorrencias").insert({
        placa: $("#n-placa").value, data_hora: new Date($("#n-data").value || Date.now()).toISOString(), tipo: $("#n-tipo").value,
        local: $("#n-local").value.trim() || null, condutor: $("#n-condutor").value.trim() || null, descricao: desc
      });
      if (error) return msg($("#n-msg"), error.message, "erro");
      fecharModal(); carregarOcorrencias();
      $$(".aba").forEach(x => x.classList.toggle("on", x.dataset.aba === "ocorrencias"));
      $$(".aba-conteudo").forEach(c => { c.hidden = c.id !== "aba-ocorrencias"; });
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
      const { error } = await sb.from("ocorrencias").update({ status: $("#t-status").value, tratativa: $("#t-texto").value.trim() || null }).eq("id", o.id);
      if (error) return msg($("#t-msg"), error.message, "erro");
      fecharModal(); carregarOcorrencias();
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
      msg(el, `Enviando ${regs.length} veículos da aba "${aba}"…`);
      // envia em grupos com as mesmas colunas (assim nenhuma coluna que ficou de fora é apagada)
      const grupos = {};
      regs.forEach(r => { const k = Object.keys(r).sort().join(","); (grupos[k] = grupos[k] || []).push(r); });
      for (const g of Object.values(grupos)) {
        for (let i = 0; i < g.length; i += 500) {
          const { error } = await sb.from("veiculos").upsert(g.slice(i, i + 500), { onConflict: "placa" });
          if (error) throw error;
        }
      }
      const faltou = CAMPOS.filter(f => idx[f.c] === undefined).map(f => f.r);
      msg(el, `✓ ${regs.length} veículos importados da aba "${aba}".` + (novos ? ` ${novos} ${novos === 1 ? "carro novo" : "carros novos"}.` : "") + (trocas ? ` ${trocas} ${trocas === 1 ? "troca" : "trocas"} de condutor ${trocas === 1 ? "registrada" : "registradas"} na linha do tempo.` : " Nenhuma troca de condutor.") + (faltou.length ? " Colunas que não estavam na planilha: " + faltou.join(", ") + "." : " Todas as colunas encontradas."), "ok");
      await carregarVeiculos();
    } catch (err) {
      msg(el, "Erro ao importar: " + (err.message || err), "erro");
    }
  });

  // ------------------------------------------------------------------ usuários (admin)
  async function carregarUsuarios() {
    const { data, error } = await sb.from("usuarios").select("*").order("nome");
    if (error) return msg($("#u-msg"), error.message, "erro");
    $("#tb-usuarios").innerHTML = (data || []).map(u => `<tr>
      <td><b>${esc(u.nome)}</b></td><td>${esc(u.email)}</td>
      <td>${esc({ admin: "Administrador", gerente: "Gerente", coordenador: "Coordenador", supervisor: "Supervisor" }[u.perfil] || u.perfil)}</td>
      <td>${esc(u.perfil === "admin" ? "Todos" : (u.perfil === "gerente" ? u.gestor_operacao : u.perfil === "coordenador" ? u.coordenador : u.supervisor)) || "—"}</td>
      <td>${u.ativo ? "Sim" : "Não"}</td>
      <td>${u.email.toLowerCase() === EU.email.toLowerCase() ? "" : `<button class="bt bt-sec" data-email="${esc(u.email)}" data-ativo="${u.ativo}">${u.ativo ? "Desativar" : "Reativar"}</button>`}</td></tr>`).join("");
    $$("#tb-usuarios button").forEach(b => b.addEventListener("click", async () => {
      const { error } = await sb.from("usuarios").update({ ativo: b.dataset.ativo !== "true" }).eq("email", b.dataset.email);
      if (error) msg($("#u-msg"), error.message, "erro"); else carregarUsuarios();
    }));
  }
  $("#form-user").addEventListener("submit", async e => {
    e.preventDefault();
    const perfil = $("#u-perfil").value, equipe = $("#u-equipe").value.trim();
    if (perfil !== "admin" && !equipe) return msg($("#u-msg"), "Informe quem essa pessoa vai ver: o Gestor operação (gerente), o Coordenador (coordenador) ou o Supervisor (supervisor).", "erro");
    const reg = {
      email: $("#u-email").value.trim().toLowerCase(), nome: $("#u-nome").value.trim(), perfil, ativo: true,
      gestor_operacao: perfil === "gerente" ? equipe : null, coordenador: perfil === "coordenador" ? equipe : null,
      supervisor: perfil === "supervisor" ? equipe : null
    };
    const { error } = await sb.from("usuarios").upsert(reg, { onConflict: "email" });
    if (error) return msg($("#u-msg"), error.message, "erro");
    msg($("#u-msg"), `✓ ${reg.nome} salvo. Agora ele entra pela tela de login em "Primeiro acesso".`, "ok");
    e.target.reset();
    carregarUsuarios();
  });
})();
