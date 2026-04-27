/* ============================================================
   Painel RPA - Task pane logic (PoC v0.2)
   Mudanças vs v0.1:
   - Layout com 2 tabs (Estado, Histórico)
   - Histórico mostra comentários
   - Bug fix: cabeçalho da folha já não aparece no histórico
   - Logo RPA no cabeçalho
   - Cliente a negrito
   - "Estado desde" mostra timestamp da última transição
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

let utilizadorEmail = "";
let utilizadorNome = "";
let estadoAtualLido = "";
let timestampUltimoEstado = "";

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    mostrarErro("Este Add-in foi desenhado para Excel.");
    return;
  }

  // Identificar utilizador
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

  // Wire tabs
  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });

  // Wire eventos
  document.getElementById("novo-estado").addEventListener("change", onEstadoChange);
  document.getElementById("btn-submeter").addEventListener("click", abrirModal);
  document.getElementById("btn-cancelar").addEventListener("click", fecharModal);
  document.getElementById("btn-confirmar").addEventListener("click", confirmarAlteracao);

  await carregarContexto();
});

function switchTab(tabName) {
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t.dataset.tab === tabName));
  document.querySelectorAll(".tab-content").forEach(c => c.classList.toggle("active", c.id === `tab-${tabName}`));
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
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      let folha = sheets.items.find(s => s.name === NOME_FOLHA_ESTADO);

      if (!folha) {
        folha = sheets.add(NOME_FOLHA_ESTADO);
        folha.visibility = Excel.SheetVisibility.veryHidden;
        folha.getRange("A1:F1").values = [["Timestamp", "Utilizador", "EstadoAnterior", "EstadoNovo", "Iteracao", "Comentario"]];
        await context.sync();

        estadoAtualLido = "";
        timestampUltimoEstado = "";
        document.getElementById("iteracao").textContent = "1";
      } else {
        const usedRange = folha.getUsedRange();
        usedRange.load("values, rowCount");
        await context.sync();

        if (usedRange.rowCount > 1) {
          const ultimaLinha = usedRange.values[usedRange.rowCount - 1];
          estadoAtualLido = ultimaLinha[3] || "";
          timestampUltimoEstado = ultimaLinha[0] || "";
          const iteracao = ultimaLinha[4] || "1";
          document.getElementById("iteracao").textContent = iteracao;
        } else {
          estadoAtualLido = "";
          timestampUltimoEstado = "";
          document.getElementById("iteracao").textContent = "1";
        }
      }

      atualizarBadgeEstado(estadoAtualLido, timestampUltimoEstado);
      await renderHistorico(folha, context);
    });
  } catch (err) {
    mostrarErro("Erro ao carregar contexto: " + err.message);
  }
}

function parseClienteFromSite(siteName) {
  // Tentar extrair "Xcability, S.A. [RPA062]" de "XcabilityS.A.RPA062"
  // Heurística: separar onde há transições maiúscula+ponto+RPA
  const matchRPA = siteName.match(/^(.+?)(RPA\d+)$/);
  if (matchRPA) {
    let nome = matchRPA[1].replace(/\.$/, "").trim();
    // Tentar inserir vírgula antes de S.A./Lda.
    nome = nome.replace(/(S\.?A\.?)$/i, ", S.A.");
    nome = nome.replace(/(Lda\.?)$/i, ", Lda");
    return `${nome} [${matchRPA[2]}]`;
  }
  return siteName;
}

async function renderHistorico(folha, context) {
  const lista = document.getElementById("historico-lista");
  try {
    const usedRange = folha.getUsedRange();
    usedRange.load("values, rowCount");
    await context.sync();

    if (usedRange.rowCount <= 1) {
      lista.innerHTML = '<div class="historico-empty">Sem alterações registadas.</div>';
      return;
    }

    // Filtrar cabeçalho (fix do bug v0.1) — saltar linha 0 que é sempre cabeçalho
    const linhas = usedRange.values.slice(1).reverse();  // mais recente primeiro

    const html = linhas.map(l => {
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
  } catch (e) {
    lista.innerHTML = '<div class="historico-empty">Erro a carregar histórico.</div>';
  }
}

function formatarTempo(timestamp) {
  // Receber p.ex. "27/04/2026, 21:55" e mostrar de forma compacta
  if (!timestamp) return "—";
  const s = String(timestamp);
  // Se for o dia de hoje, só mostra hora
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
    // Marcar visualmente o textarea
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
    // Reset border do textarea
    document.getElementById("modal-comentario").style.border = "1px solid #ccc";

    await carregarContexto();

    try {
      Office.context.document.settings.saveAsync();
    } catch (e) { /* silencioso */ }

  } catch (err) {
    mostrarErro("Erro ao gravar: " + err.message);
  }
}

async function gravarEstado(novoEstado, comentario) {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    let folha = sheets.items.find(s => s.name === NOME_FOLHA_ESTADO);
    if (!folha) {
      folha = sheets.add(NOME_FOLHA_ESTADO);
      folha.visibility = Excel.SheetVisibility.veryHidden;
      folha.getRange("A1:F1").values = [["Timestamp", "Utilizador", "EstadoAnterior", "EstadoNovo", "Iteracao", "Comentario"]];
      await context.sync();
    }

    const usedRange = folha.getUsedRange();
    usedRange.load("rowCount, values");
    await context.sync();

    let estadoAnterior = "";
    let iteracao = 1;

    if (usedRange.rowCount > 1) {
      const ultima = usedRange.values[usedRange.rowCount - 1];
      estadoAnterior = ultima[3] || "";
      const iterAnt = parseInt(ultima[4]) || 1;
      // Modelo B: incrementam só Devolvido→Em Curso (F7) e Revisto→Em Curso (F8)
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

    const novaLinha = folha.getRange(`A${usedRange.rowCount + 1}:F${usedRange.rowCount + 1}`);
    novaLinha.values = [[timestamp, utilizadorNome, estadoAnterior, novoEstado, iteracao, comentario]];

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
