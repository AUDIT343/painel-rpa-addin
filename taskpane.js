/* ============================================================
   Painel RPA - Task pane logic (v0.4)
   - Tabela RPA_Estado_Tbl com 6 colunas (sem Notificado — anti-dup faz-se no flow via Audit Log)
   - Migração robusta: se folha existe sem tabela ou com estrutura errada, recria limpa
   - Trata corretamente o caso veryHidden + delete em Excel Online
   - NOVO v0.4: upsert para lista Estado PT no site SharePoint do cliente
   ============================================================ */
 
const ESTADO_CONFIG = {
  "Em Curso":   { comentario: "opcional",    modalTitle: "Marcar Em Curso",        modalSub: "PT volta a ser editável." },
  "Pendente":   { comentario: "obrigatorio", modalTitle: "Marcar como Pendente",   modalSub: "PT fica pausado a aguardar resposta do cliente." },
  "Concluído":  { comentario: "obrigatorio", modalTitle: "Marcar como Concluído",  modalSub: "PT segue para revisão pelo Sócio." },
  "Devolvido":  { comentario: "obrigatorio", modalTitle: "Devolver ao Auditor",    modalSub: "PT volta para correção. Notas obrigatórias." },
  "Revisto":    { comentario: "obrigatorio", modalTitle: "Aprovar como Revisto",   modalSub: "PT fica bloqueado para edição." }
};
 
const ESTADO_CSS_CLASS = {
  "Em Curso": "emcurso",
  "Pendente": "pendente",
  "Concluído": "concluido",
  "Revisto": "revisto",
  "Devolvido": "devolvido"
};
 
const NOME_FOLHA_ESTADO = "_RPA_Estado";
const NOME_TABELA_ESTADO = "RPA_Estado_Tbl";
const COLUNAS_TABELA = ["Timestamp", "Utilizador", "EstadoAnterior", "EstadoNovo", "Iteracao", "Comentario"];
const NUM_COLUNAS = COLUNAS_TABELA.length;
 
const NOME_LISTA_ESTADO_PT = "Estado PT";
const COLUNAS_LISTA = {
  Title: "Title",
  Ficheiro: "Ficheiro",
  Caminho: "Caminho",
  Link: "Link",
  Ano: "Ano",
  Estado: "Estado",
  Auditor: "Auditor",
  Iteracao: "Itera_x00e7__x00e3_o",
  Comentario: "Coment_x00e1_rio",
  UltimaAtualizacao: "_x00da_ltimaatualiza_x00e7__x00e"
};
 
let utilizadorEmail = "";
let utilizadorNome = "";
let estadoAtualLido = "";
let timestampUltimoEstado = "";
let fileContext = null;
 
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    mostrarErro("Este Add-in foi desenhado para Excel.");
    return;
  }
 
  try {
    if (Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
      utilizadorEmail = Office.context.mailbox.userProfile.emailAddress || "";
      utilizadorNome  = Office.context.mailbox.userProfile.displayName || utilizadorEmail;
    }
  } catch (e) { /* silencioso */ }
 
  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });
 
  document.getElementById("novo-estado").addEventListener("change", onEstadoChange);
  document.getElementById("btn-submeter").addEventListener("click", abrirModal);
  document.getElementById("btn-cancelar").addEventListener("click", fecharModal);
  document.getElementById("btn-confirmar").addEventListener("click", confirmarAlteracao);
 
  const btnSync = document.getElementById("btn-sync-mapa");
  if (btnSync) btnSync.addEventListener("click", sincronizarMapaManual);
 
  fileContext = parseFileContext();
 
  if (fileContext && fileContext.siteUrl) {
    try {
      const user = await getCurrentUser(fileContext.siteUrl);
      if (user) {
        utilizadorEmail = user.email || utilizadorEmail;
        utilizadorNome = user.title || utilizadorNome || utilizadorEmail;
      }
    } catch (e) {
      console.warn("Não foi possível obter utilizador via REST:", e.message);
    }
  }
 
  if (!utilizadorNome) {
    utilizadorNome = "(identificado pelo flow no save)";
  }
 
  document.getElementById("utilizador").textContent = utilizadorNome;
 
  await carregarContexto();
});
 
function switchTab(tabName) {
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t.dataset.tab === tabName));
  document.querySelectorAll(".tab-content").forEach(c => c.classList.toggle("active", c.id === `tab-${tabName}`));
}
 
function parseFileContext() {
  try {
    const url = Office.context.document.url || "";
    if (!url) return null;
 
    const matchSite = url.match(/^(https:\/\/[^/]+\/sites\/[^/]+)/i);
    if (!matchSite) return null;
    const siteUrl = matchSite[1];
 
    const afterSite = url.substring(siteUrl.length);
    const decoded = decodeURIComponent(afterSite);
 
    const lastSlash = decoded.lastIndexOf("/");
    if (lastSlash < 0) return null;
 
    const ficheiro = decoded.substring(lastSlash + 1).split("?")[0];
    const caminho = decoded.substring(0, lastSlash + 1);
 
    const matchAno = caminho.match(/\/(20\d{2})\//);
    const ano = matchAno ? parseInt(matchAno[1]) : null;
 
    const link = url.split("?")[0];
 
    const ctx = { siteUrl, caminho, ficheiro, ano, link };
    console.log("FileContext:", ctx);
    return ctx;
  } catch (e) {
    console.warn("parseFileContext falhou:", e.message);
    return null;
  }
}
 
async function getRequestDigest(siteUrl) {
  const resp = await fetch(`${siteUrl}/_api/contextinfo`, {
    method: "POST",
    credentials: "include",
    headers: {
      "Accept": "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose"
    }
  });
  if (!resp.ok) throw new Error(`contextinfo HTTP ${resp.status}`);
  const data = await resp.json();
  return data.d.GetContextWebInformation.FormDigestValue;
}
 
async function getCurrentUser(siteUrl) {
  const resp = await fetch(`${siteUrl}/_api/web/currentuser?$select=Id,Email,Title,LoginName`, {
    method: "GET",
    credentials: "include",
    headers: { "Accept": "application/json;odata=verbose" }
  });
  if (!resp.ok) return null;
  const data = await resp.json();
  return {
    id: data.d.Id,
    email: data.d.Email,
    title: data.d.Title,
    loginName: data.d.LoginName
  };
}
 
async function findExistingItem(siteUrl, ano, caminho, ficheiro) {
  const filter = encodeURIComponent(
    `${COLUNAS_LISTA.Ano} eq ${ano} and ${COLUNAS_LISTA.Ficheiro} eq '${ficheiro.replace(/'/g, "''")}'`
  );
  const url = `${siteUrl}/_api/web/lists/getbytitle('${NOME_LISTA_ESTADO_PT}')/items?$filter=${filter}&$select=Id,${COLUNAS_LISTA.Caminho}&$top=10`;
 
  const resp = await fetch(url, {
    method: "GET",
    credentials: "include",
    headers: { "Accept": "application/json;odata=verbose" }
  });
  if (!resp.ok) throw new Error(`findItem HTTP ${resp.status}`);
  const data = await resp.json();
 
  const matches = data.d.results || [];
  const exact = matches.find(m => (m[COLUNAS_LISTA.Caminho] || "") === caminho);
  return exact || null;
}
 
async function getListItemEntityType(siteUrl) {
  const url = `${siteUrl}/_api/web/lists/getbytitle('${NOME_LISTA_ESTADO_PT}')?$select=ListItemEntityTypeFullName`;
  const resp = await fetch(url, {
    method: "GET",
    credentials: "include",
    headers: { "Accept": "application/json;odata=verbose" }
  });
  if (!resp.ok) throw new Error(`getListType HTTP ${resp.status}`);
  const data = await resp.json();
  return data.d.ListItemEntityTypeFullName;
}
 
async function upsertEstadoPT(estado, comentario, iteracao, autor) {
  if (!fileContext || !fileContext.siteUrl) {
    throw new Error("Contexto do ficheiro não disponível.");
  }
  if (!fileContext.ano) {
    throw new Error("Ano não detetado no caminho do ficheiro.");
  }
 
  const { siteUrl, ano, caminho, ficheiro, link } = fileContext;
  const nomePT = ficheiro.replace(/\.(xlsx|xlsm|xls|docx|doc)$/i, "");
 
  const digest = await getRequestDigest(siteUrl);
  const entityType = await getListItemEntityType(siteUrl);
 
  const existing = await findExistingItem(siteUrl, ano, caminho, ficheiro);
 
  const payload = {
    "__metadata": { "type": entityType },
    [COLUNAS_LISTA.Title]: nomePT,
    [COLUNAS_LISTA.Ficheiro]: ficheiro,
    [COLUNAS_LISTA.Caminho]: caminho,
    [COLUNAS_LISTA.Link]: { "__metadata": { "type": "SP.FieldUrlValue" }, "Url": link, "Description": nomePT },
    [COLUNAS_LISTA.Ano]: ano,
    [COLUNAS_LISTA.Estado]: estado,
    [COLUNAS_LISTA.Auditor]: autor || "",
    [COLUNAS_LISTA.Iteracao]: iteracao,
    [COLUNAS_LISTA.Comentario]: comentario || "",
    [COLUNAS_LISTA.UltimaAtualizacao]: new Date().toISOString()
  };
 
  let url, method, headers;
 
  if (existing) {
    url = `${siteUrl}/_api/web/lists/getbytitle('${NOME_LISTA_ESTADO_PT}')/items(${existing.Id})`;
    method = "POST";
    headers = {
      "Accept": "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*"
    };
  } else {
    url = `${siteUrl}/_api/web/lists/getbytitle('${NOME_LISTA_ESTADO_PT}')/items`;
    method = "POST";
    headers = {
      "Accept": "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest
    };
  }
 
  const resp = await fetch(url, {
    method,
    credentials: "include",
    headers,
    body: JSON.stringify(payload)
  });
 
  if (!resp.ok) {
    let errMsg = `HTTP ${resp.status}`;
    try {
      const errData = await resp.json();
      if (errData.error && errData.error.message) {
        errMsg = errData.error.message.value || errMsg;
      }
    } catch (_) {}
    throw new Error(errMsg);
  }
 
  return existing ? "updated" : "created";
}
 
/**
 * Garante que existe a folha _RPA_Estado com tabela RPA_Estado_Tbl com NUM_COLUNAS colunas.
 */
async function garantirEstrutura(context) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
 
  let folha = sheets.items.find(s => s.name === NOME_FOLHA_ESTADO);
 
  if (folha) {
    const tables = folha.tables;
    tables.load("items/name");
    await context.sync();
 
    let tabelaOK = false;
 
    if (tables.items.length > 0) {
      const tabela = folha.tables.getItem(NOME_TABELA_ESTADO);
      try {
        const cols = tabela.columns;
        cols.load("items/name");
        await context.sync();
 
        const nomesCol = cols.items.map(c => c.name);
        const todasPresentes = COLUNAS_TABELA.every(nome => nomesCol.includes(nome));
 
        if (todasPresentes && nomesCol.length === NUM_COLUNAS) {
          tabelaOK = true;
        }
      } catch (e) {
        tabelaOK = false;
      }
    }
 
    if (!tabelaOK) {
      try {
        folha.visibility = Excel.SheetVisibility.visible;
        await context.sync();
      } catch (e) { /* silencioso */ }
 
      try {
        folha.delete();
        await context.sync();
        folha = null;
      } catch (e) {
        console.error("Delete folha falhou, a tentar limpar:", e.message);
        try {
          folha.getUsedRange().clear();
          const tabs = folha.tables;
          tabs.load("items/name");
          await context.sync();
          for (const t of tabs.items) {
            try { t.delete(); } catch (_) {}
          }
          await context.sync();
        } catch (e2) {
          mostrarErro("Não foi possível limpar a estrutura. Apague manualmente a folha _RPA_Estado.");
          throw e2;
        }
      }
    }
  }
 
  if (!folha) {
    folha = sheets.add(NOME_FOLHA_ESTADO);
  }
 
  const usedRange = folha.getUsedRangeOrNullObject();
  usedRange.load("rowCount, columnCount");
  await context.sync();
 
  if (usedRange.isNullObject || usedRange.rowCount === 0) {
    const colLetterEnd = String.fromCharCode(65 + NUM_COLUNAS - 1);
    const range = folha.getRange(`A1:${colLetterEnd}1`);
    range.values = [COLUNAS_TABELA];
 
    const tabela = folha.tables.add(`A1:${colLetterEnd}1`, true);
    tabela.name = NOME_TABELA_ESTADO;
 
    await context.sync();
  }
 
  folha.visibility = Excel.SheetVisibility.veryHidden;
  await context.sync();
 
  return folha;
}
 
async function lerTabela(context, folha) {
  try {
    const tabela = folha.tables.getItem(NOME_TABELA_ESTADO);
    const dataRange = tabela.getDataBodyRange();
    dataRange.load("values, rowCount");
    await context.sync();
    return dataRange.values || [];
  } catch (e) {
    return [];
  }
}
 
async function carregarContexto() {
  try {
    const url = Office.context.document.url || "";
    const partes = url.split("/");
    const nomeFicheiro = decodeURIComponent(partes[partes.length - 1] || "(desconhecido)");
    document.getElementById("ficheiro").textContent = nomeFicheiro;
 
    const matchSite = url.match(/\/sites\/([^/]+)\//);
    if (matchSite && matchSite[1]) {
      const siteName = decodeURIComponent(matchSite[1]);
      document.getElementById("cliente").textContent = parseClienteFromSite(siteName);
    } else {
      document.getElementById("cliente").textContent = "(local — sem SharePoint)";
    }
 
    await Excel.run(async (context) => {
      const folha = await garantirEstrutura(context);
      const linhas = await lerTabela(context, folha);
 
      if (linhas.length === 0) {
        estadoAtualLido = "";
        timestampUltimoEstado = "";
        document.getElementById("iteracao").textContent = "1";
      } else {
        const ultima = linhas[linhas.length - 1];
        estadoAtualLido = ultima[3] || "";
        timestampUltimoEstado = ultima[0] || "";
        const iteracao = ultima[4] || "1";
        document.getElementById("iteracao").textContent = iteracao;
      }
 
      atualizarBadgeEstado(estadoAtualLido, timestampUltimoEstado);
      renderHistorico(linhas);
    });
  } catch (err) {
    mostrarErro("Erro ao carregar contexto: " + err.message);
  }
}
 
function parseClienteFromSite(siteName) {
  const matchRPA = siteName.match(/^(.+?)(RPA\d+)$/);
  if (matchRPA) {
    let nome = matchRPA[1].replace(/\.$/, "").trim();
    nome = nome.replace(/(S\.?A\.?)$/i, ", S.A.");
    nome = nome.replace(/(Lda\.?)$/i, ", Lda");
    return `${nome} [${matchRPA[2]}]`;
  }
  return siteName;
}
 
function renderHistorico(linhas) {
  const lista = document.getElementById("historico-lista");
 
  if (!linhas || linhas.length === 0) {
    lista.innerHTML = '<div class="historico-empty">Sem alterações registadas.</div>';
    return;
  }
 
  const linhasInvertidas = [...linhas].reverse();
 
  const html = linhasInvertidas.map(l => {
    const [timestamp, utilizador, estadoAnt, estadoNovo, iteracao, comentario] = l;
    const estadoClass = ESTADO_CSS_CLASS[estadoNovo] || "vazio";
    const tempo = formatarTempo(timestamp);
    const comentarioHtml = comentario && String(comentario).trim()
      ? `<div class="historico-comentario">"${escapeHtml(String(comentario))}"</div>`
      : `<div class="historico-comentario empty">(sem comentário)</div>`;
 
    return `
      <div class="historico-item bg-${estadoClass}">
        <div class="historico-header">
          <span class="historico-badge estado-${estadoClass}">${escapeHtml(estadoNovo || "(vazio)")}</span>
          <span class="historico-time">${tempo}</span>
        </div>
        <div class="historico-meta">${escapeHtml(utilizador || "—")} · Iteração ${iteracao || 1}</div>
        ${comentarioHtml}
      </div>
    `;
  }).join("");
 
  lista.innerHTML = html;
}
 
function formatarTempo(timestamp) {
  if (!timestamp) return "—";
  const s = String(timestamp);
  const hoje = new Date();
  const hojeStr = hoje.toLocaleDateString("pt-PT", { day: "2-digit", month: "2-digit", year: "numeric" });
  if (s.startsWith(hojeStr)) {
    const partes = s.split(",");
    return partes.length > 1 ? partes[1].trim() : s;
  }
  return s;
}
 
function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = String(str);
  return div.innerHTML;
}
 
function atualizarBadgeEstado(estado, timestamp) {
  const el = document.getElementById("estado-atual");
  const since = document.getElementById("estado-since");
  const cls = "estado-" + (ESTADO_CSS_CLASS[estado] || "vazio");
  const txt = estado || "(vazio)";
  el.innerHTML = `<span class="estado-badge ${cls}">${txt}</span>`;
  since.textContent = timestamp ? `desde ${formatarTempo(timestamp)}` : "";
}
 
function onEstadoChange() {
  const valor = document.getElementById("novo-estado").value;
  document.getElementById("btn-submeter").disabled = !valor;
}
 
function abrirModal() {
  const novoEstado = document.getElementById("novo-estado").value;
  if (!novoEstado) return;
 
  const cfg = ESTADO_CONFIG[novoEstado];
  document.getElementById("modal-title").textContent = cfg.modalTitle;
  document.getElementById("modal-subtitle").textContent = cfg.modalSub;
 
  const labelEl = document.getElementById("modal-label");
  if (cfg.comentario === "obrigatorio") {
    labelEl.innerHTML = 'Comentário <span class="obrigatorio">(obrigatório)</span>';
  } else {
    labelEl.innerHTML = 'Comentário <span class="opcional">(opcional)</span>';
  }
 
  document.getElementById("modal-comentario").value = "";
  document.getElementById("modal-comentario").style.border = "1px solid #ccc";
  document.getElementById("modal").classList.add("active");
  setTimeout(() => document.getElementById("modal-comentario").focus(), 50);
}
 
function fecharModal() {
  document.getElementById("modal").classList.remove("active");
}
 
async function confirmarAlteracao() {
  const novoEstado = document.getElementById("novo-estado").value;
  const comentario = document.getElementById("modal-comentario").value.trim();
  const cfg = ESTADO_CONFIG[novoEstado];
 
  if (cfg.comentario === "obrigatorio" && !comentario) {
    const ta = document.getElementById("modal-comentario");
    ta.style.border = "1px solid #C13838";
    ta.placeholder = "Comentário é obrigatório.";
    ta.focus();
    return;
  }
 
  let iteracaoUsada = 1;
 
  try {
    iteracaoUsada = await gravarEstado(novoEstado, comentario);
    fecharModal();
    mostrarSucesso(`Estado alterado para "${novoEstado}".`);
 
    document.getElementById("novo-estado").value = "";
    document.getElementById("btn-submeter").disabled = true;
 
    await carregarContexto();
 
    try {
      Office.context.document.settings.saveAsync();
    } catch (e) { /* silencioso */ }
 
  } catch (err) {
    mostrarErro("Erro ao gravar: " + err.message);
    return;
  }
 
  try {
    await upsertEstadoPT(novoEstado, comentario, iteracaoUsada, utilizadorNome);
  } catch (err) {
    console.warn("Mapa não atualizado:", err.message);
    mostrarWarning("Mapa não atualizado: " + err.message + ". Use 'Sincronizar mapa' para tentar de novo.");
  }
}
 
async function sincronizarMapaManual() {
  if (!fileContext) {
    fileContext = parseFileContext();
  }
  if (!fileContext || !fileContext.siteUrl) {
    mostrarErro("Não foi possível identificar o site SharePoint deste ficheiro.");
    return;
  }
  if (!fileContext.ano) {
    mostrarErro("Ano não detetado no caminho do ficheiro. Verifique se o ficheiro está em /YYYY/.");
    return;
  }
 
  let estadoAtual = "";
  let comentarioAtual = "";
  let iteracaoAtual = 1;
 
  try {
    await Excel.run(async (context) => {
      const folha = await garantirEstrutura(context);
      const linhas = await lerTabela(context, folha);
      if (linhas.length === 0) {
        throw new Error("Tabela local vazia. Defina o estado primeiro.");
      }
      const ultima = linhas[linhas.length - 1];
      estadoAtual = ultima[3] || "";
      comentarioAtual = ultima[5] || "";
      iteracaoAtual = parseInt(ultima[4]) || 1;
    });
  } catch (err) {
    mostrarErro("Erro ao ler estado local: " + err.message);
    return;
  }
 
  if (!estadoAtual) {
    mostrarErro("Estado atual vazio. Defina o estado primeiro.");
    return;
  }
 
  const btn = document.getElementById("btn-sync-mapa");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "A sincronizar...";
  }
 
  try {
    const result = await upsertEstadoPT(estadoAtual, comentarioAtual, iteracaoAtual, utilizadorNome);
    mostrarSucesso(result === "created" ? "Mapa criado." : "Mapa atualizado.");
  } catch (err) {
    mostrarErro("Falha ao sincronizar mapa: " + err.message);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Sincronizar mapa";
    }
  }
}
 
async function gravarEstado(novoEstado, comentario) {
  let iteracaoUsada = 1;
 
  await Excel.run(async (context) => {
    const folha = await garantirEstrutura(context);
    const linhasExistentes = await lerTabela(context, folha);
 
    let estadoAnterior = "";
    let iteracao = 1;
 
    if (linhasExistentes.length > 0) {
      const ultima = linhasExistentes[linhasExistentes.length - 1];
      estadoAnterior = ultima[3] || "";
      const iterAnt = parseInt(ultima[4]) || 1;
      if ((estadoAnterior === "Devolvido" && novoEstado === "Em Curso") ||
          (estadoAnterior === "Revisto" && novoEstado === "Em Curso")) {
        iteracao = iterAnt + 1;
      } else {
        iteracao = iterAnt;
      }
    }
 
    iteracaoUsada = iteracao;
 
    const timestamp = new Date().toLocaleString("pt-PT", {
      day: "2-digit", month: "2-digit", year: "numeric",
      hour: "2-digit", minute: "2-digit"
    });
 
    const tabela = folha.tables.getItem(NOME_TABELA_ESTADO);
    tabela.rows.add(null, [[timestamp, utilizadorNome, estadoAnterior, novoEstado, iteracao, comentario]]);
 
    await context.sync();
  });
 
  return iteracaoUsada;
}
 
function mostrarSucesso(msg) {
  const el = document.getElementById("feedback");
  el.className = "feedback success";
  el.textContent = msg;
  setTimeout(() => { el.className = "feedback"; el.textContent = ""; }, 4000);
}
 
function mostrarWarning(msg) {
  const el = document.getElementById("feedback");
  el.className = "feedback warning";
  el.textContent = msg;
  setTimeout(() => { el.className = "feedback"; el.textContent = ""; }, 8000);
}
 
function mostrarErro(msg) {
  const el = document.getElementById("feedback");
  el.className = "feedback error";
  el.textContent = msg;
}
 
