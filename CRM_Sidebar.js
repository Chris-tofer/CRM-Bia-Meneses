<!DOCTYPE html>
<!--
CRM_Sidebar.html - v2.9.1
NOVO: Autosync da aba ativa otimizado (cache no client + intervalo 5s com execução imediata ao abrir).
NOVO: Ajuste visual dos botões no painel RESPONDI (padding menor e mesmo estilo no "Ir para este lead no CRM").
v2.8:
NOVO: Botão "Ir para este lead no CRM" no painel RESPONDI.
v2.7:
NOVO: Informativo sobre como cadastrar novo formulário.
Opção de limpar Interesse padrão de um formulário.
Botão e texto de ajuda, na sidebar.
AMPLIAÇÃO DA BUSCA: 
- Agora é possível pesquisar formulários por nome ou parte do nome, além do ID que já era possível
- Título da busca corrigido
- Placeholder atualizado
-->
<html>
<head>
  <base target="_top" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">

<style>
  body { font-family: 'Roboto', Arial, sans-serif; padding: 0; background: #ffffff; color: #333; margin: 0; }
  
  /* ABAS PRINCIPAIS */
  .tabs { display: flex; background: #ffffff; padding: 0; border-bottom: 1px solid #e0e0e0; }
  .tab { flex: 1; border: 0; padding: 12px 5px; font-weight: 700; font-size: 10px; text-transform: uppercase; cursor: pointer; background: #ffffff; color: #5f6368; transition: all 0.2s ease; display: flex; flex-direction: column; align-items: center; gap: 4px; }
  .tab .material-icons { font-size: 20px; }
  .tab.active, .tab:hover { background: #6200ee; color: #ffffff; }

  /* SUB-ABAS CONFIG */
  .sub-tabs { display: flex; gap: 10px; padding: 10px 15px; background: #f1f3f4; border-bottom: 1px solid #e0e0e0; }
  .sub-tab { flex: 1; padding: 8px; border: 1px solid #dadce0; border-radius: 20px; background: #fff; font-size: 10px; font-weight: 700; text-transform: uppercase; cursor: pointer; color: #5f6368; text-align: center; }
  .sub-tab.active { background: #6200ee; color: #fff; border-color: #6200ee; }

  /* HISTÓRICO */
  .history-bar { display: flex; gap: 15px; padding: 8px 16px; background: #f8f9fa; border-bottom: 1px solid #e0e0e0; align-items: center; }
  .btnHist { background: none; border: none; cursor: pointer; color: #6200ee; display: flex; align-items: center; padding: 4px; border-radius: 50%; transition: background 0.2s; }
  .btnHist:hover:not(:disabled) { background: #eee; }
  .btnHist:disabled { color: #ccc; cursor: default; }

  /* NAVEGAÇÃO */
  .section { background: #fff; margin-bottom: 5px; }
  .secTitle { padding: 12px 16px; font-size: 11px; font-weight: 700; text-transform: uppercase; color: #70757a; display: flex; cursor: pointer; align-items: center; justify-content: space-between; }
  .btnNav { width: 100%; border: 0; padding: 10px 16px; display: flex; gap: 12px; align-items: center; cursor: pointer; background: transparent; color: #3c4043; font-weight: 500; }
  .btnNav.active, .btnNav:hover { background: #f3e5f5; color: #6200ee; border-right: 4px solid #6200ee; }
  .ico { color: #6200ee; }

  /* CARDS E FORMULÁRIOS */
  .panel { display: none; padding: 12px; }
  .panel.active { display: block; }
  
  /* Ajuste no Card: Espaçamento entre cards */
  .card { background: #ffffff; border: 1px solid #e0e0e0; border-radius: 12px; padding: 16px; margin: 0 12px 12px 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
  
  /* Ajuste na Pergunta: Fonte maior e mais espaço inferior (estilo antigo) */
  .q { font-weight: 700; font-size: 15px; color: #1a1c1e; margin-bottom: 10px; display: block; line-height: 1.4; }
  
  /* Ajuste na Resposta: Fonte maior e melhor altura de linha (estilo antigo) */
  .a { font-size: 14px; color: #3c4043; line-height: 1.5; white-space: pre-wrap; word-break: break-word; }
  
  .pill { background: #ffffff; border: 1px solid #dadce0; padding: 4px 10px; font-size: 11px; border-radius: 4px; color: #5f6368; font-weight: 500; margin-right: 4px; margin-bottom: 4px; display: inline-block; }
  
  label { display: block; font-size: 11px; font-weight: 700; color: #5f6368; margin-bottom: 4px; margin-top: 12px; text-transform: uppercase; }
  input, select, textarea { width: 100%; padding: 10px; border: 1px solid #dadce0; border-radius: 6px; box-sizing: border-box; font-family: inherit; font-size: 13px; }
  
  /* Ajuste no Botão: padding menor (mais "magro") */
  .btn-action {
    width: 100%;
    background: #6200ee;
    color: white;
    border: none;
    padding: 8px 10px; /* <- era 10px */
    border-radius: 6px;
    font-weight: 700;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    margin-top: 10px;
    transition: all 0.2s;
  }
  .btn-action .material-icons{
    font-size: 18px;
    line-height: 1;
  }

  .btn-action:hover { opacity: 0.9; }

  .hidden { display: none; }

  /* HELP MODAL (não altera o layout existente; só aparece quando aberto) */
  .helpOverlay { position: fixed; inset: 0; background: rgba(0,0,0,0.35); display: none; z-index: 9998; }
  .helpModal { position: fixed; left: 50%; top: 50%; transform: translate(-50%, -50%); width: calc(100% - 32px); max-width: 420px;
               background: #fff; border: 1px solid #e0e0e0; border-radius: 12px; box-shadow: 0 8px 24px rgba(0,0,0,0.18);
               display: none; z-index: 9999; }
  .helpHeader { display:flex; align-items:center; justify-content:space-between; padding: 12px 14px; border-bottom: 1px solid #eee; }
  .helpTitle { font-weight: 700; font-size: 12px; text-transform: uppercase; color: #5f6368; }
  .helpClose { background: none; border: none; cursor: pointer; color: #6200ee; display:flex; align-items:center; padding: 4px; border-radius: 50%; }
  .helpClose:hover { background: #eee; }
  .helpBody { padding: 12px 14px 14px 14px; font-size: 13px; color: #3c4043; line-height: 1.5; white-space: normal; max-height: 65vh; overflow-y: auto;}
  .helpBody h3 { margin: 10px 0 6px 0; font-size: 13px; color:#1a1c1e; }
  .helpBody ul { margin: 6px 0 0 18px; padding: 0; }
  .helpBody li { margin: 4px 0; }
</style>
</head>

<body>
  <div class="tabs">
    <button id="tabNav" class="tab active" onclick="showTab('nav')">
      <span class="material-icons">explore</span>MENU
    </button>
    <button id="tabLead" class="tab" onclick="showTab('lead')">
      <span class="material-icons">assignment</span>RESPONDI
    </button>
    <button id="tabCfg" class="tab" onclick="showTab('cfg')">
      <span class="material-icons">settings</span>CONFIG
    </button>
  </div>

  <div class="history-bar">
    <button id="btnBack" class="btnHist" onclick="historyBack()" disabled title="Voltar">
      <span class="material-icons">arrow_back</span>
    </button>
    <button id="btnForward" class="btnHist" onclick="historyForward()" disabled title="Avançar">
      <span class="material-icons">arrow_forward</span>
    </button>
    <span style="font-size: 10px; color: #999; text-transform: uppercase; font-weight: 700;">Histórico</span>
    <button id="btnHelp" class="btnHist" onclick="openHelp()" title="Ajuda">
      <span class="material-icons">help_outline</span>
    </button>
  </div>

  <div id="panelNav" class="panel active">
    <div id="navContent"></div>
  </div>

  <div id="panelLead" class="panel" style="background:#f8f9fa; min-height:100vh; padding:0;">
    <div style="padding:16px">
      <button class="btn-action" style="background:#fff; color:#6200ee; border:1px solid #6200ee" onclick="loadSelected()">
        <span class="material-icons">refresh</span>Atualizar Formulário
      </button>

      <!-- >>> INSERIDO (v2.8): Botão para ir ao CRM pelo respondent_id da seleção <<< -->
      <button class="btn-action" id="btnGoCrm" style="background:#fff; color:#6200ee; border:1px solid #6200ee" onclick="void(0)">
        <span class="material-icons">search</span>Ir para este lead no CRM
      </button>

      <div id="goCrmMsg" style="margin-top:8px;font-size:12px;opacity:.85;"></div>

      <script>
        (function(){
          const btn = document.getElementById('btnGoCrm');
          const msg = document.getElementById('goCrmMsg');

          btn.addEventListener('click', () => {
            msg.textContent = 'Procurando no CRM...';
            google.script.run
              .withSuccessHandler(res => {
                if (!res || !res.ok) {
                  msg.textContent = (res && res.error) ? res.error : 'Falha ao procurar.';
                  return;
                }
                msg.textContent = `OK: respondent_id ${res.respondent_id} (linha ${res.row})`;
              })
              .withFailureHandler(err => {
                msg.textContent = (err && err.message) ? err.message : String(err);
              })
              .uiGoToCrmByRespondentIdFromSelection();
          });
        })();
      </script>
      <!-- >>> FIM INSERIDO <<< -->
    </div>
    <div id="meta" style="padding: 0 16px 12px 16px"></div>
    <div id="content"></div>
  </div>

  <div id="panelCfg" class="panel" style="padding:0">
    <div class="sub-tabs">
      <div id="subTabForms" class="sub-tab active" onclick="showSubCfg('forms')">Formulários</div>
      <div id="subTabLists" class="sub-tab" onclick="showSubCfg('lists')">Listas do CRM</div>
    </div>

    <div style="padding:15px">
      <div id="cfgForms">
        <label>Pesquisar ID ou NOME do Formulário</label>
        <div style="display:flex; gap:5px">
          <input type="text" id="cfgSearchId" placeholder="Ex: 7SQRLgDL ou Mentoria">
          <button class="sub-tab active" style="border-radius:6px; flex:none; padding:0 12px" onclick="searchFormMapping()">
            <span class="material-icons" style="font-size:18px">search</span>
          </button>
        </div
        ><!-- Ajuda: como cadastrar/editar formulário -->
        <div style="display:flex; align-items:center; justify-content:space-between; margin-top:10px; padding:10px 12px; border:1px solid #eee; border-radius:10px; background:#fff;">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-icons" style="font-size:18px; color:#6200ee;">info</span>
            <span style="font-size:12px; font-weight:700; color:#3c4043;">Como adicionar ou editar um formulário</span>
          </div>

          <button type="button" class="btnHist" onclick="openCfgFormsHelp()" title="Ver passo a passo" style="padding:6px;">
            <span class="material-icons">help_outline</span>
          </button>
        </div>
        <hr style="border:0; border-top:1px solid #eee; margin: 15px 0">
        <label>ID Form</label>
        <input type="text" id="cfgFormId" readonly style="background:#f8f9fa">
        <label>Nome do Formulário</label>
        <input type="text" id="cfgFormName">
        <label>Funil Destino</label>
        <select id="cfgFunnel"></select>
        <label>Interesse Padrão</label>
        <select id="cfgInterest"></select>
        <button class="btn-action" onclick="saveFormMapping()"><span class="material-icons">save</span>Salvar Mapeamento</button>
      </div>

      <div id="cfgLists" class="hidden">
        <label>Lista para Editar</label>
        <select id="cfgListSelector" onchange="loadListContent()">
          <option value="" disabled selected>Escolha...</option>
          <option value="Infoprodutos">Infoprodutos</option>
          <option value="Objeções">Objeções</option>
          <option value="Origem_Lead">Origem_Lead</option>
          <option value="Responsável">Responsável</option>
          <option value="Segmento_Perfil">Segmento_Perfil</option>
          <option value="Servicos_Agencia">Servicos_Agencia</option>
          <option value="Status_Funil">Status_Funil</option>
          <option value="Tipo_Relacao">Tipo_Relacao</option>
        </select>
        <label>Itens (Um por linha)</label>
        <textarea id="cfgListTextArea" style="height:180px; width:100%; border:1px solid #dadce0; border-radius:6px; padding:10px;"></textarea>
        <button class="btn-action" onclick="saveConfigList()"><span class="material-icons">cloud_upload</span>Atualizar Lista</button>
      </div>
    </div>
  </div>

  <!-- HELP MODAL -->
  <div id="helpOverlay" class="helpOverlay" onclick="closeHelp()"></div>
  <div id="helpModal" class="helpModal" role="dialog" aria-modal="true" aria-labelledby="helpTitle">
    <div class="helpHeader">
      <div id="helpTitle" class="helpTitle">Ajuda</div>
      <button class="helpClose" onclick="closeHelp()" title="Fechar">
        <span class="material-icons">close</span>
      </button>
    </div>
    <div id="helpBody" class="helpBody"></div>
  </div>

<script>
  let historyStack = [], historyIndex = -1, lastDetectedSheet = "", syncInterval = null, allLists = {};

  // AJUDA (por aba)
  const HELP_BY_SHEET = {
    "CRM_GERAL": `

      <h3>CRM_GERAL</h3>
      <ul>
        <li>Use esta aba para acompanhar o lead no funil (status, temperatura, responsável, etc.).</li>
        <li>As colunas E, F, G, M, N, O, P e AE usam validação via listas da aba <b>CONFIG</b>.</li>
        <li>O lead entra com <b>raw_id</b> (col B), <b>respondent_id</b> (col C) e <b>lead_key</b> (col A).</li>
      </ul>
      <h3>Dicas rápidas</h3>
      <ul>
        <li>Se editar @/WhatsApp/Email, mantenha o formato limpo (sem texto extra).</li>
        <li>Evite apagar raw_id/respondent_id: isso quebra rastreio com a RAW.</li>
      </ul>
    `,
    "RAW_RESPONDI": `

      <h3>RAW_RESPONDI</h3>
      <ul>
        <li>É o “banco de dados” bruto: cada resposta de formulário vira uma linha.</li>
        <li>Cada resposta de formulário vira uma linha aqui.</li>
      </ul>
      <h3><b>⚠️ ATENÇÃO</b></h3>
      <ul>
        <li><b>Jamais edite ou apague informações nesta aba.</b></li>
        <li>Qualquer alteração aqui pode causar perda de dados ou erros no CRM.</li>
      </ul>
    `,
    "INFOPRODUTOS": `

      <h3>INFOPRODUTOS</h3>
      <ul>
        <li><b>Somente visualização.</b> Mostra apenas leads interessados em <b>infoprodutos</b>.</li>
        <li>Serve para acompanhar esse tipo de lead sem precisar aplicar filtros.</li>
      </ul>
      <h3>O que você pode fazer aqui</h3>
      <ul>
        <li>Consultar leads, status e prioridade.</li>
        <li>Entender rapidamente quem está mais quente.</li>
      </ul>
      <h3>Importante</h3>
      <ul>
        <li>Para editar qualquer informação do lead, use a aba <b>CRM_GERAL</b>.</li>
      </ul>
    `,
    "SERVICOS_AGENCIA": `

      <h3>SERVICOS_AGENCIA</h3>
      <ul>
        <li><b>Somente visualização.</b> Mostra apenas leads de <b>serviços da agência</b>.</li>
        <li>Serve para acompanhar oportunidades de serviço sem precisar aplicar filtros.</li>
      </ul>
      <h3>O que você pode fazer aqui</h3>
      <ul>
        <li>Consultar leads, status e follow-ups.</li>
        <li>Priorizar negociações e propostas.</li>
      </ul>
      <h3>Importante</h3>
      <ul>
        <li>Para editar qualquer informação do lead, use a aba <b>CRM_GERAL</b>.</li>
      </ul>
    `,
    "MAPEAMENTO_FORMS": `

      <h3>MAPEAMENTO_FORMS</h3>
      <ul>
        <li><b>Somente visualização.</b> Mostra como cada formulário alimenta o CRM.</li>
        <li>Serve para conferir funil, interesse e destino padrão de cada <b>form</b>.</li>
      </ul>
      <h3>Importante</h3>
      <ul>
        <li>Se precisar mudar algo, use a <b>aba Config</b> do <b>menu lateral</b>.</li>
        <li>Evite editar diretamente nesta planilha.</li>
      </ul>
    `,
    "CONFIG": `

      <h3>CONFIG</h3>
      <ul>
        <li><b>Somente visualização.</b> Área técnica de configuração do CRM.</li>
        <li>Aqui ficam regras e listas usadas nas validações do sistema.</li>
      </ul>
      <h3>Importante</h3>
      <ul>
        <li>Qualquer ajuste deve ser feito pela <b>aba Config</b> do <b>menu lateral</b>, não por aqui.</li>
      </ul>
    `
  };

  function getHelpHtmlForSheet_(sheetName) {
    if (!sheetName) return "<p>Não consegui detectar a aba ativa.</p>";
    if (HELP_BY_SHEET[sheetName]) return HELP_BY_SHEET[sheetName];

    // Abas de funil (FUNIL_*)
    if (sheetName.startsWith("FUNIL_")) {
      return `
        <h3>${sheetName}</h3>
        <ul>
          <li><b>Somente visualização.</b> Mostra leads que entraram neste funil.</li>
          <li>Serve para acompanhar esse funil sem precisar aplicar filtros.</li>
        </ul>
        <h3>O que você pode fazer aqui</h3>
        <ul>
          <li>Ver quem entrou neste funil.</li>
          <li>Ler um resumo curto para identificar rapidamente os leads mais promissores.</li>
          <li>Consultar leads, status e prioridade.</li>
        </ul>
        <h3>Como ver o formulário completo</h3>
        <ul>
          <li>Clique em qualquer célula da linha do lead que quiser ver todas as respostas do formulário.
          Neste menu lateral clique em <b>Respondi</b>. Para ver outro, clique novamente na linha desejada e em <b>Atualizar Seleção</b>.</li>
        </ul>
        <h3>Importante</h3>
        <ul>
          <li>Para editar qualquer informação do lead, use a aba <b>CRM_GERAL</b>.</li>
        </ul>
      `;
    }

    return `
      <h3>${sheetName}</h3>
      <ul>
        <li>Ainda não existe um texto de ajuda específico para esta aba.</li>
        <li>Se quiser, me diga o objetivo desta aba e eu te devolvo a ajuda pronta para colar no mapa.</li>
      </ul>
    `;
  }

  function openHelp() {
    // preferir a aba detectada pelo autosync, mas confirmar pelo servidor quando possível
    const localGuess = lastDetectedSheet || "";
    google.script.run.withSuccessHandler(sheetName => {
      const name = sheetName || localGuess;
      document.getElementById("helpTitle").textContent = "Ajuda • " + (name || "Aba");
      document.getElementById("helpBody").innerHTML = getHelpHtmlForSheet_(name);
      document.getElementById("helpOverlay").style.display = "block";
      document.getElementById("helpModal").style.display = "block";
    }).getCurrentSheetName();
  }

  function closeHelp() {
    document.getElementById("helpOverlay").style.display = "none";
    document.getElementById("helpModal").style.display = "none";
  }

  function openCfgFormsHelp() {
    const html = `
      <h3>Adicionar novo formulário (cadastro)</h3>
      <ul>
        <li><b>Copie o ID do formulário</b> no Respondi.</li>
        <li>No menu lateral, vá em <b>Config → Formulários</b>.</li>
        <li>No campo <b>Pesquisar ID ou NOME</b>, cole o ID e clique na <b>lupa</b>.</li>
        <li>Se aparecer “Nenhum formulário encontrado”, o sistema mantém o ID no campo <b>ID Form</b> para você cadastrar.</li>
        <li>Preencha o <b>Nome do Formulário</b></li>
        <li>Escolha o <b>Funil Destino</b>.</li>
        <li>Escolha o <b>Interesse Padrão</b> (ou <b>- Nenhum (Limpar) -</b> se não quiser fixar).</li>
        <li>Clique em <b>Salvar Mapeamento</b>.</li>
        <li>Pronto: a próxima resposta desse formulário já entra no funil/CRM com esse padrão.</li>
      </ul>

      <h3>Editar formulário existente</h3>
      <ul>
        <li>No campo <b>Pesquisar ID ou NOME</b>, digite o <b>ID</b> ou parte do <b>nome</b> e clique na <b>lupa</b>.</li>
        <li>Se aparecer 1 resultado, confirme para carregar.</li>
        <li>Se aparecer mais de um, <b>digite o número</b> correspondente ao correto e clique em "OK"</li>
        <li>Com os campos carregados, ajuste <b>Nome</b>, <b>Funil Destino</b> e/ou <b>Interesse Padrão</b>.</li>
        <li>Clique em <b>Salvar Mapeamento</b> para gravar a alteração.</li>
      </ul>

      <h3>Observações rápidas</h3>
      <ul>
        <li>O campo <b>ID Form</b> é <b>somente leitura</b> (vem da busca).</li>
        <li>Se você mudar o <b>Funil Destino</b>, isso afeta <b>novas entradas</b> (não reorganiza leads antigos automaticamente).</li>
        <li><b>- Nenhum (Limpar) -</b> remove o interesse padrão (o interesse pode vir em branco ou ser detectado pelo sistema).</li>
      </ul>
    `;

    document.getElementById("helpTitle").textContent = "Ajuda • Formulários (Config)";
    document.getElementById("helpBody").innerHTML = html;
    document.getElementById("helpOverlay").style.display = "block";
    document.getElementById("helpModal").style.display = "block";
  }

  // INICIALIZAÇÃO
  window.onload = function() { refreshNav(); };

  function showTab(which){
    ["nav", "lead", "cfg"].forEach(t => {
      document.getElementById("tab" + t.charAt(0).toUpperCase() + t.slice(1)).classList.toggle("active", which === t);
      document.getElementById("panel" + t.charAt(0).toUpperCase() + t.slice(1)).classList.toggle("active", which === t);
    });
   
    if(which === "nav") { refreshNav(); startAutoSync(); } 
    else { stopAutoSync(); if(which === "lead") loadSelected(); if(which === "cfg") loadConfigData(); }
  }

  function showSubCfg(type) {
    document.getElementById("subTabForms").classList.toggle("active", type === 'forms');
    document.getElementById("subTabLists").classList.toggle("active", type === 'lists');
    document.getElementById("cfgForms").classList.toggle("hidden", type !== 'forms');
    document.getElementById("cfgLists").classList.toggle("hidden", type !== 'lists');
  }

  function loadConfigData() {
    google.script.run.withSuccessHandler(lists => {
      allLists = lists;
      if(lists.Funis_Entrada) {
        document.getElementById("cfgFunnel").innerHTML = lists.Funis_Entrada.map(i => `<option value="${i}">${i}</option>`).join("");
      }
      if(lists.Interesse) {
        // Cria a opção "Nenhum" no topo manualmente
        const emptyOpt = `<option value="">- Nenhum (Limpar) -</option>`;
        
        // Gera as opções vindas da planilha
        const listOpts = lists.Interesse.map(i => `<option value="${i}">${i}</option>`).join("");
        
        // Junta tudo e coloca no HTML
        document.getElementById("cfgInterest").innerHTML = emptyOpt + listOpts;
      }
    }).uiGetConfigLists();
  }

  function loadListContent() {
    const listName = document.getElementById("cfgListSelector").value;
    document.getElementById("cfgListTextArea").value = (allLists[listName] || []).join("\n");
  }

  function saveConfigList() {
    const listName = document.getElementById("cfgListSelector").value;
    const content = document.getElementById("cfgListTextArea").value;
    google.script.run.withSuccessHandler(msg => alert(msg)).uiSaveConfigList(listName, content);
  }

function searchFormMapping() {
    const term = document.getElementById("cfgSearchId").value.trim();
    if (!term) return;

    google.script.run.withSuccessHandler(results => {
      if (results && results.length > 0) {
        let selectedForm = null;

        if (results.length === 1) {
          // Se houver apenas um, confirma diretamente
          if(confirm("Encontrado: " + results[0].form_name + "\n\nDeseja carregar os dados?")) {
            selectedForm = results[0];
          }
        } else {
          // Se houver vários, cria uma lista para escolha simples (prompt ou confirm expandido)
          const names = results.map((r, i) => (i + 1) + ") " + r.form_name).join("\n");
          const choice = prompt("Múltiplos formulários encontrados. Digite o número correspondente:\n\n" + names);
          const index = parseInt(choice) - 1;
          
          if (results[index]) {
            selectedForm = results[index];
          }
        }

        if (selectedForm) {
          document.getElementById("cfgFormId").value = selectedForm.form_id;
          document.getElementById("cfgFormName").value = selectedForm.form_name;
          document.getElementById("cfgFunnel").value = selectedForm.funnel_default;
          document.getElementById("cfgInterest").value = selectedForm.interest_default;
        }
      } else {
        alert("Nenhum formulário encontrado com o termo: " + term);
        // Se for um ID novo, mantém no campo de ID para cadastro
        document.getElementById("cfgFormId").value = term;
      }
    }).uiGetFormMapping(term);
  }

  function saveFormMapping() {
    const data = {
      form_id: document.getElementById("cfgFormId").value,
      form_name: document.getElementById("cfgFormName").value,
      funnel_default: document.getElementById("cfgFunnel").value,
      interest_default: document.getElementById("cfgInterest").value
    };
    google.script.run.withSuccessHandler(msg => alert(msg)).uiSaveFormMapping(data);
  }

  // NAVEGAÇÃO E RENDERIZAÇÃO
  function refreshNav(){
    google.script.run.withSuccessHandler(renderNav).navGetConfig();
  }

  function renderNav(cfg){
    const root = document.getElementById("navContent");
    const sections = cfg.sections || [];
    root.innerHTML = sections.map((sec, i) => `
      <div class="section">
        <div class="secTitle" onclick="toggle('items_${i}')">
          <span>${sec.title}</span>
          <span class="material-icons" style="font-size:16px">expand_more</span>
        </div>
        <div class="items" id="items_${i}">
          ${(sec.items || []).map(it => `
            <button class="btnNav" data-sheet="${it.sheet}" onclick="go('${it.sheet}')">
              <span class="ico">${it.icon}</span>
              <span class="lbl">${it.label}</span>
            </button>
          `).join("")}
        </div>
      </div>
    `).join("");
    if (cfg.activeSheet) {
      syncActiveSheetHighlight(cfg.activeSheet);
    }

    startAutoSync();
  }

  function go(sheet){
    google.script.run.navGoToSheet(sheet);
    if(historyStack[historyIndex] !== sheet) {
      historyStack = historyStack.slice(0, historyIndex + 1);
      historyStack.push(sheet);
      historyIndex++;
      updateHistoryButtons();
    }
  }

  function historyBack(){ if(historyIndex > 0) goHistory(historyStack[--historyIndex]); }
  function historyForward(){ if(historyIndex < historyStack.length - 1) goHistory(historyStack[++historyIndex]); }
  function goHistory(s){ google.script.run.navGoToSheet(s); updateHistoryButtons(); }
  function updateHistoryButtons(){
    document.getElementById("btnBack").disabled = historyIndex <= 0;
    document.getElementById("btnForward").disabled = historyIndex >= historyStack.length - 1;
  }

  function loadSelected(){
    google.script.run.withSuccessHandler(renderLead).uiFetchLeadFromActiveSelection();
  }

  function renderLead(d){
    const meta = document.getElementById("meta"), el = document.getElementById("content");
    if(!d || d.error) { el.innerHTML = `<div class="card">${d?.error || "Erro"}</div>`; return; }
    meta.innerHTML = `<div class="pill" style="border-color:#6200ee; color:#6200ee; font-weight:700">📋 ${d.form_name}</div><span class="pill">aba: ${d._sheet}</span><span class="pill">linha: ${d._row}</span>`;
    el.innerHTML = (d.qa || []).map(it => `<div class="card"><span class="q">${it.q}</span><div class="a">${it.a}</div></div>`).join("");
  }

function syncActiveSheetHighlight(sheetName) {
  const buttons = document.querySelectorAll(".btnNav");
  buttons.forEach(btn => {
    btn.classList.toggle(
      "active",
      btn.getAttribute("data-sheet") === sheetName
    );
  });
}

function startAutoSync() {
  if (syncInterval) return;

  function pollActiveSheet() {
    google.script.run
      .withSuccessHandler(sheetName => {
        if (!sheetName) return;

        // Cache no client: só atualiza se mudou
        if (sheetName === lastDetectedSheet) return;

        lastDetectedSheet = sheetName;
        syncActiveSheetHighlight(sheetName);
      })
      .getCurrentSheetName();
  }

  // roda uma vez ao abrir a sidebar (resposta imediata)
  pollActiveSheet();

  // e depois repete em intervalos maiores (menos chamadas ao servidor)
  syncInterval = setInterval(pollActiveSheet, 7000);
}

function stopAutoSync() {
  if (syncInterval) {
    clearInterval(syncInterval);
    syncInterval = null;
  }
}

  function toggle(id){ const el = document.getElementById(id); if(el) el.classList.toggle("hidden"); }
</script>
</body>
</html>
