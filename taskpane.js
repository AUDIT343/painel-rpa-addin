/* ============================================================
   Painel RPA - Task pane logic (v0.4.1)
   - Tabela RPA_Estado_Tbl com 6 colunas (sem Notificado — anti-dup faz-se no flow via Audit Log)
   - Migração robusta: se folha existe sem tabela ou com estrutura errada, recria limpa
   - Trata corretamente o caso veryHidden + delete em Excel Online
   - v0.4.1: botão "Sincronizar mapa" força save do ficheiro (flow F-Sync2 trata do upsert)
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

let utilizadorEmail = "";
let utilizadorNome = "";
let estadoAtualLido = "";
let timestampUltimoEstado = "";

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

  if (!utilizadorEmail) {
    utilizadorNome = "(identificado pelo flow no save)";
  }

  document.getElementById("utilizador").textContent = utilizadorNome;

  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });

  document.getElementById("novo-estado").addEventListener("change", onEstadoChange);
  document.getElementById("btn-submeter").addEventListener("click", abrirModal);
  document.getElementById("btn-cancelar").addEventListener("click", fecharModal);
  document.getElementById("btn-confirmar").addEventListener("click", confirmarAlteracao);

  const btnSync = document.getElementById("btn-sync-mapa");
  if (btnSync) btnSync.addEventListener("click", forcarSave);

  await carregarContexto();
});

function switchTab(tabName) {
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t.dataset.tab === tabName));
  document.querySelectorAll(".tab-content").forEach(c => c.classList.toggle("active", c.id === `tab-${tabName}`));
}

/**
 * Garante que existe a folha _RPA_Estado com tabela RPA_Estado_Tbl com NUM_COLUNAS colunas.
 * Migração robusta: tenta primeiro adaptar, se falhar apaga e recria.
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
  return String(timestamp);
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

  try {
    await gravarEstado(novoEstado, comentario);
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
  }
}

async function forcarSave() {
  const btn = document.getElementById("btn-sync-mapa");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "A guardar...";
  }

  try {
    await new Promise((resolve, reject) => {
      Office.context.document.settings.saveAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error ? result.error.message : "Erro desconhecido"));
        }
      });
    });
    mostrarSucesso("Ficheiro guardado. Mapa será atualizado em segundos.");
  } catch (err) {
    mostrarErro("Erro ao guardar: " + err.message);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Sincronizar mapa";
    }
  }
}

async function gravarEstado(novoEstado, comentario) {
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

    const timestamp = new Date().toLocaleString("pt-PT", {
      day: "2-digit", month: "2-digit", year: "numeric",
      hour: "2-digit", minute: "2-digit"
    });

    const tabela = folha.tables.getItem(NOME_TABELA_ESTADO);
    tabela.rows.add(null, [[timestamp, utilizadorNome, estadoAnterior, novoEstado, iteracao, comentario]]);

    await context.sync();
  });
}

function mostrarSucesso(msg) {
  const el = document.getElementById("feedback");
  el.className = "feedback success";
  el.textContent = msg;
  setTimeout(() => { el.className = "feedback"; el.textContent = ""; }, 4000);
}

function mostrarErro(msg) {
  const el = document.getElementById("feedback");
  el.className = "feedback error";
  el.textContent = msg;
}
