# Instalação do Painel RPA — Passo a passo para AUDIT343

URL final que vai usar: **https://audit343.github.io/painel-rpa-addin/**

---

## PARTE 1 — Criar repositório GitHub e fazer upload

### 1.1 Criar repositório

1. Vá a https://github.com/new (já com sessão iniciada como AUDIT343)
2. **Repository name:** `painel-rpa-addin` (exatamente assim — em minúsculas)
3. **Description:** "Painel de workflow para papéis de trabalho RPA SROC"
4. **Visibility:** ⚠️ **Public** (necessário para GitHub Pages gratuito)
5. ✅ Marcar **Add a README file**
6. Clicar **Create repository**

### 1.2 Upload dos ficheiros

1. No repositório acabado de criar, clicar **Add file → Upload files**
2. Arrastar para a janela os seguintes ficheiros (da pasta `rpa-addin` que descarregou):
   - `manifest.xml`
   - `taskpane.html`
   - `taskpane.js`
   - `README.md` (substitui o que GitHub criou)
3. Arrastar também a **pasta `assets`** inteira (com os 3 PNGs)
4. Em "Commit changes" escrever: `Initial commit - PoC v0.1`
5. Clicar **Commit changes**

Verificar que ficou assim:
```
painel-rpa-addin/
├── README.md
├── manifest.xml
├── taskpane.html
├── taskpane.js
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png
```

### 1.3 Ativar GitHub Pages

1. No repositório → **Settings** (separador no topo)
2. Menu lateral esquerdo → **Pages**
3. **Source:** escolher `Deploy from a branch`
4. **Branch:** `main` / `/ (root)` → **Save**
5. Esperar 1-2 minutos
6. Recarregar a página Settings → Pages — deve aparecer:
   `Your site is live at https://audit343.github.io/painel-rpa-addin/`

### 1.4 Validar que está acessível

Abrir no browser: **https://audit343.github.io/painel-rpa-addin/taskpane.html**

Deve mostrar o painel RPA (sem dados, porque está fora do contexto Excel — é normal). Se mostrar erro 404, esperar mais 1-2 minutos. Se persistir, verificar Settings → Pages.

Também testar: **https://audit343.github.io/painel-rpa-addin/assets/icon-32.png**
Deve mostrar o quadrado azul com "RPA".

---

## PARTE 2 — Sideload do Add-in no Excel Online

### 2.1 Abrir Excel Online

1. Abrir https://www.office.com (ou https://www.microsoft365.com)
2. Login com a sua conta `ricardopereira@rpa-sroc.pt`
3. Abrir um qualquer ficheiro `.xlsx` para teste
   - **Recomendação:** criar **Novo livro em branco** primeiro para testar; só depois aplicar a um PT real

### 2.2 Carregar o manifest

1. Menu **Inserir → Suplementos** (ou no menu Base depende da versão)
2. Janela "Suplementos do Office" → separador **OS MEUS SUPLEMENTOS**
3. Clicar **Gerir os meus suplementos** → **Carregar Os Meus Suplementos** (link no fundo da janela)
4. **Procurar** → escolher o ficheiro `manifest.xml` no seu computador
5. Clicar **Carregar**

### 2.3 Confirmar instalação

- No separador **Base** da ribbon do Excel deve aparecer um grupo **RPA** com botão **Painel RPA**
- Clicar no botão → task pane abre à direita

---

## PARTE 3 — Primeiros testes

Com um livro de teste em branco aberto:

### Teste 1 — Painel abre
- ✅ Painel à direita
- ✅ Cabeçalho "Painel RPA" com quadrado azul
- ✅ Cliente, Papel de trabalho, Utilizador, Iteração visíveis (alguns "(local)" se for ficheiro local)
- ✅ Estado atual = "(vazio)"

### Teste 2 — Marcar Em Curso (sem comentário)
1. Combo → "Em Curso"
2. Botão "Submeter alteração" fica ativo
3. Clicar → modal abre com label "(opcional)"
4. Deixar comentário vazio → "Confirmar"
5. ✅ Mensagem verde "Estado alterado para Em Curso"
6. ✅ Badge muda para "Em Curso" (amarelo)
7. ✅ Histórico mostra a entrada

### Teste 3 — Marcar Pendente sem comentário
1. Combo → "Pendente" → Submeter → modal abre com "(obrigatório)"
2. Deixar vazio → Confirmar → ✅ Alert "Comentário é obrigatório"

### Teste 4 — Marcar Pendente com comentário
1. Mesmo passo, agora escrever "Aguardo confirmação saldo Cliente XYZ"
2. Confirmar → ✅ Sucesso, badge "Pendente" (vermelho claro)

### Teste 5 — Confirmar folha oculta criada
1. Clique direito em qualquer tab de folha → **Mostrar**
2. ✅ `_RPA_Estado` aparece na lista
3. Mostrar → ver que tem 3 linhas (cabeçalho + 2 entradas)
4. Voltar a ocultar (clique direito → Ocultar)

### Teste 6 — Modelo B de iteração
1. Combo → "Em Curso" → Confirmar (resolver pendência) — iteração mantém-se em 1
2. Combo → "Pendente" com comentário → Confirmar
3. Combo → "Em Curso" → Confirmar — iteração mantém-se em 1 ✅
   *(Pendente → Em Curso não incrementa, é F2b)*

### Teste 7 — Devolvido → Em Curso incrementa
1. Combo → "Concluído" com comentário → Confirmar
2. Combo → "Devolvido" com comentário → Confirmar
3. Combo → "Em Curso" → Confirmar — ✅ **iteração passa a 2**
   *(Devolvido → Em Curso é F7, modelo B)*

### Teste 8 — Reabrir o ficheiro
1. Fechar o ficheiro
2. Voltar a abrir
3. Abrir Painel RPA
4. ✅ Estado lido corretamente da última entrada da folha
5. ✅ Iteração correta

---

## Se algo correr mal

| Problema | Solução |
|---|---|
| 404 no taskpane.html | Esperar 2 min, verificar Settings → Pages, branch main, root |
| "Não foi possível carregar suplemento" | Verificar manifest.xml na sua versão tem URLs `audit343.github.io` (não `RPA_BASE_URL`) |
| Botão RPA não aparece | Fechar e reabrir Excel Online; ver Suplementos → Os Meus Suplementos |
| Painel abre mas vazio/erro | F12 no browser → Consola → ver erros JavaScript e enviar-me |
| Folha _RPA_Estado não aparece nem em "Mostrar" | É `veryHidden`. Para a ver: Visual Basic Editor (Alt+F11) → propriedade Visible |

---

## Quando isto estiver a funcionar

Avise-me e passamos para:

1. **Replicar para Word** (Add-in idêntico, com ajustes)
2. **Modificar F1-F8** para ler folha `_RPA_Estado` em vez do trigger SharePoint convencional
3. **Resolver o "fechar ficheiro automaticamente"** com workaround viável
4. **Bloqueio Nível 3** no F4 (BreakRoleInheritance)
5. **Deploy centralizado** via M365 Admin Center para os 8 utilizadores

---

**Notas importantes:**

- ⚠️ **GitHub Pages cache 5-10 min** — depois de atualizar ficheiros pode demorar a refletir. Hard refresh no browser (Ctrl+Shift+R)
- ⚠️ **Sideload é manual e por utilizador** — só você terá o Add-in agora; quando validarmos, fazemos deploy central
- ⚠️ **PoC não toca em SharePoint ainda** — só grava em folha oculta. Os flows F1-F8 atuais continuam a funcionar normalmente em paralelo. Não há risco de partir nada.
- ✅ **Pode usar em qualquer xlsx** — não só do Xcability — para os testes iniciais

---

Bom trabalho. Avise quando tiver feito a Parte 1, antes de avançar para a Parte 2.
