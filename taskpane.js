/* ============================================================
   Painel RPA - Task pane logic (PoC v0.1)
   - Lê metadata do ficheiro Excel
   - Combo-box de estados + modal de comentário
   - Grava em folha oculta _RPA_Estado
   - Fecha o ficheiro (se possível na PoC, senão pede ao utilizador)
   ============================================================ */

// Configuração estados → comportamento
const ESTADO_CONFIG = {
  "Em Curso":   { comentario: "opcional",  modalTitle: "Marcar Em Curso",        modalSub: "PT volta a ser editável." },
  "Pendente":   { comentario: "obrigatorio", modalTitle: "Marcar como Pendente",  modalSub: "PT fica pausado a aguardar resposta do cliente." },
  "Concluído":  { comentario: "obrigatorio", modalTitle: "Marcar como Concluído", modalSub: "PT segue para revisão pelo Sócio." },
  "Devolvido":  { comentario: "obrigatorio", modalTitle: "Devolver ao Auditor",   modalSub: "PT volta para correção. Notas obrigatórias." },
  "Revisto":    { comentario: "obrigatorio", modalTitle: "Aprovar como Revisto",  modalSub: "PT fica bloqueado para edição." }
};

const NOME_FOLHA_ESTADO = "_RPA_Estado";

let utilizadorEmail = "";
let utilizadorNome = "";
let estadoAtualLido = "";

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
  } catch (e) {
    // Excel não tem mailbox context. Vamos usar fallback.
  }

  // Fallback: pedir ao Office.context.document
  if (!utilizadorEmail) {
    try {
      // Em Excel, não há API direta para email do utilizador atual.
      // Solução: usar Office.context.auth.getAccessTokenAsync (SSO) — fica para v2.
      // PoC: deixar "Utilizador" para o Power Automate identificar via Editor do trigger.
      utilizadorNome = "(identificado pelo flow no save)";
    } catch (e) {
      utilizadorNome = "(desconhecido)";
    }
  }

  document.getElementById("utilizador").textContent = utilizadorNome;

  // Ler metadata do ficheiro
  await carregarContexto();

  // Wire eventos
  document.getElementById("novo-estado").addEventListener("change", onEstadoChange);
  document.getElementById("btn-submeter").addEventListener("click", abrirModal);
  document.getElementById("btn-cancelar").addEventListener("click", fecharModal);
  document.getElementById("btn-confirmar").addEventListener("click", confirmarAlteracao);
});

async function carregarContexto() {
  try {
    // Nome do ficheiro
    const url = Office.context.document.url || "";
    const partes = url.split("/");
    const nomeFicheiro = decodeURIComponent(partes[partes.length - 1] || "(desconhecido)");
    document.getElementById("ficheiro").textContent = nomeFicheiro;

    // Cliente — extrair do URL (procura padrão /sites/XXX/)
    const matchSite = url.match(/\/sites\/([^/]+)\//);
    if (matchSite && matchSite[1]) {
      // Heurística: substring entre "sites/" e o "."
      const siteName = decodeURIComponent(matchSite[1]);
      // Tenta extrair "Xcability" de "XcabilityS.A.RPA062"
      const matchCliente = siteName.match(/^([A-Za-z]+)/);
      if (matchCliente) {
        document.getElementById("cliente").textContent = siteName.replace(/RPA\d+$/, "").replace(/\.$/, "");
      } else {
        document.getElementById("cliente").textContent = siteName;
      }
    } else {
      document.getElementById("cliente").textContent = "(local — sem SharePoint)";
    }

    // Ler estado atual da folha _RPA_Estado (se existir) ou assumir vazio
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      let folha = sheets.items.find(s => s.name === NOME_FOLHA_ESTADO);

      if (!folha) {
        // Cria folha oculta na primeira utilização
        folha = sheets.add(NOME_FOLHA_ESTADO);
        folha.visibility = Excel.SheetVisibility.veryHidden;

        const headers = folha.getRange("A1:F1");
        headers.values = [["Timestamp", "Utilizador", "EstadoAnterior", "EstadoNovo", "Iteracao", "Comentario"]];
        await context.sync();

        estadoAtualLido = "";
        document.getElementById("iteracao").textContent = "1";
      } else {
        // Lê última linha
        const usedRange = folha.getUsedRange();
        usedRange.load("values, rowCount");
        await context.sync();

        if (usedRange.rowCount > 1) {
          const ultimaLinha = usedRange.values[usedRange.rowCount - 1];
          estadoAtualLido = ultimaLinha[3] || "";  // EstadoNovo
          const iteracao = ultimaLinha[4] || "1";
          document.getElementById("iteracao").textContent = iteracao;
        } else {
          estadoAtualLido = "";
          document.getElementById("iteracao").textContent = "1";
        }
      }

      atualizarBadgeEstado(estadoAtualLido);
      await renderHistorico(folha, context);
    });
  } catch (err) {
    mostrarErro("Erro ao carregar contexto: " + err.message);
  }
}

async function renderHistorico(folha, context) {
  try {
    const usedRange = folha.getUsedRange();
    usedRange.load("values, rowCount");
    await context.sync();

    if (usedRange.rowCount <= 1) {
      document.getElementById("historico").textContent = "Sem alterações nesta sessão.";
      return;
    }

    const linhas = usedRange.values.slice(-3).reverse();  // últimas 3
    const html = linhas
      .filter(l => l[0])
      .map(l => `${l[0]} — ${l[1]} → ${l[3]}`)
      .join("<br>");
    document.getElementById("historico").innerHTML = html || "Sem alterações nesta sessão.";
  } catch (e) {
    // silencioso
  }
}

function atualizarBadgeEstado(estado) {
  const el = document.getElementById("estado-atual");
  const classMap = {
    "Em Curso": "estado-emcurso",
    "Pendente": "estado-pendente",
    "Concluído": "estado-concluido",
    "Revisto": "estado-revisto",
    "Devolvido": "estado-devolvido"
  };
  const cls = classMap[estado] || "estado-vazio";
  const txt = estado || "(vazio)";
  el.innerHTML = `<span class="estado-badge ${cls}">${txt}</span>`;
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
    alert("Comentário é obrigatório para esta transição.");
    return;
  }

  try {
    await gravarEstado(novoEstado, comentario);
    fecharModal();
    mostrarSucesso(`Estado alterado para "${novoEstado}". O ficheiro será guardado.`);

    // Reset combo
    document.getElementById("novo-estado").value = "";
    document.getElementById("btn-submeter").disabled = true;

    // Recarregar contexto
    await carregarContexto();

    // Tentar fechar/guardar o ficheiro
    // Office.js não tem document.close() universal — só Office.context.document.close() em algumas hosts
    // PoC: forçar save e mostrar mensagem
    try {
      Office.context.document.settings.saveAsync();
    } catch (e) {
      // ignorar
    }

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
      // Modelo B: F7/F8 incrementam (Devolvido → Em Curso, ou Revisto → Em Curso)
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
  setTimeout(() => { el.className = "feedback"; el.textContent = ""; }, 5000);
}

function mostrarErro(msg) {
  const el = document.getElementById("feedback");
  el.className = "feedback error";
  el.textContent = msg;
}
