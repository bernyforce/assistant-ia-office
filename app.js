/* ===========================================
   ASSISTANT IA DOCUMENTS - Office Add-in
   =========================================== */

// ── Stockage global ──
const DocumentStore = {
  documents: [], // { name, type, content, charCount }
  
  add(name, type, content) {
    // Éviter les doublons
    this.documents = this.documents.filter(d => d.name !== name);
    this.documents.push({
      name,
      type,
      content,
      charCount: content.length,
      addedAt: new Date().toLocaleTimeString()
    });
    this.updateUI();
  },

  remove(name) {
    this.documents = this.documents.filter(d => d.name !== name);
    this.updateUI();
  },

  clear() {
    this.documents = [];
    this.updateUI();
  },

  updateUI() {
    const countEl = document.getElementById('docCount');
    const listEl = document.getElementById('docsList');
    countEl.textContent = this.documents.length;
    
    listEl.innerHTML = this.documents.map(doc => `
      <div class="doc-item">
        <span class="doc-name" title="${doc.name}">
          ${getFileIcon(doc.type)} ${doc.name}
        </span>
        <span class="doc-size">${formatSize(doc.charCount)}</span>
        <span class="doc-remove" onclick="DocumentStore.remove('${doc.name.replace(/'/g, "\\'")}')">✕</span>
      </div>
    `).join('');
  },

  // Construire le contexte pour le LLM
  buildContext() {
    if (this.documents.length === 0) return '';
    
    return this.documents.map((doc, i) => {
      // Tronquer les très gros documents pour respecter la fenêtre de contexte
      const maxChars = 15000; // ~4000 tokens par doc
      const content = doc.content.length > maxChars 
        ? doc.content.substring(0, maxChars) + '\n\n[... document tronqué ...]'
        : doc.content;
      
      return `\n═══════════════════════════════════════\n` +
             `📄 DOCUMENT ${i + 1}: "${doc.name}" (${doc.type})\n` +
             `═══════════════════════════════════════\n` +
             content;
    }).join('\n\n');
  }
};

// ── Configuration ──
const Config = {
  _defaults: {
    llmProvider: 'openrouter',
    apiUrl: 'https://openrouter.ai/api/v1/chat/completions',
    apiKey: '',
    modelName: 'moonshotai/kimi-k2.5',
    responseLang: 'fr',
    maxTokens: 2048
  },

  load() {
    try {
      const saved = localStorage.getItem('assistantIA_config');
      return saved ? { ...this._defaults, ...JSON.parse(saved) } : { ...this._defaults };
    } catch {
      return { ...this._defaults };
    }
  },

  save(config) {
    localStorage.setItem('assistantIA_config', JSON.stringify(config));
  },

  applyToUI() {
    const cfg = this.load();
    document.getElementById('llmProvider').value = cfg.llmProvider;
    document.getElementById('apiUrl').value = cfg.apiUrl;
    document.getElementById('apiKey').value = cfg.apiKey;
    document.getElementById('modelName').value = cfg.modelName;
    document.getElementById('responseLang').value = cfg.responseLang;
    document.getElementById('maxTokens').value = cfg.maxTokens;
  },

  readFromUI() {
    return {
      llmProvider: document.getElementById('llmProvider').value,
      apiUrl: document.getElementById('apiUrl').value.trim(),
      apiKey: document.getElementById('apiKey').value.trim(),
      modelName: document.getElementById('modelName').value.trim(),
      responseLang: document.getElementById('responseLang').value,
      maxTokens: parseInt(document.getElementById('maxTokens').value) || 2048
    };
  }
};

// ── Extracteurs de texte ──
const TextExtractor = {
  
  async extractFromFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    setStatus(`Extraction de ${file.name}...`);
    
    switch (ext) {
      case 'pdf':     return await this.extractPDF(file);
      case 'docx':
      case 'doc':     return await this.extractDOCX(file);
      case 'xlsx':
      case 'xls':
      case 'csv':     return await this.extractExcel(file);
      case 'pptx':
      case 'ppt':     return await this.extractPPTX(file);
      case 'txt':
      case 'md':      return await this.extractText(file);
      default:
        throw new Error(`Format non supporté: .${ext}`);
    }
  },

  // PDF via pdf.js
  async extractPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';
    
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(item => item.str).join(' ');
      fullText += `\n--- Page ${i} ---\n${pageText}`;
    }
    return fullText.trim();
  },

  // DOCX via Mammoth.js
  async extractDOCX(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    return result.value;
  },

  // Excel via SheetJS
  async extractExcel(file) {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    let fullText = '';
    
    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const csv = XLSX.utils.sheet_to_csv(sheet);
      fullText += `\n--- Feuille: ${sheetName} ---\n${csv}`;
    });
    return fullText.trim();
  },

  // PPTX (fichier ZIP contenant des XML)
  async extractPPTX(file) {
    // Approche : utiliser JSZip pour décompresser et extraire le texte
    // On charge JSZip dynamiquement si nécessaire
    if (typeof JSZip === 'undefined') {
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
    }
    
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    let fullText = '';
    let slideNum = 0;
    
    // Les slides sont dans ppt/slides/slide1.xml, slide2.xml, etc.
    const slideFiles = Object.keys(zip.files)
      .filter(name => name.match(/ppt\/slides\/slide\d+\.xml$/))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)/)[1]);
        const numB = parseInt(b.match(/slide(\d+)/)[1]);
        return numA - numB;
      });

    for (const slidePath of slideFiles) {
      slideNum++;
      const xmlStr = await zip.files[slidePath].async('text');
      // Extraire le texte des balises XML <a:t>
      const textMatches = xmlStr.match(/<a:t[^>]*>([^<]*)<\/a:t>/g) || [];
      const slideText = textMatches
        .map(m => m.replace(/<[^>]+>/g, ''))
        .join(' ');
      fullText += `\n--- Slide ${slideNum} ---\n${slideText}`;
    }
    return fullText.trim() || '[Aucun texte extractible de ce fichier PowerPoint]';
  },

  // Texte brut
  async extractText(file) {
    return await file.text();
  }
};

// ── Communication avec le LLM ──
const LLMClient = {
  
  async ask(question, context) {
    const cfg = Config.load();
    
    if (!cfg.apiKey && cfg.llmProvider !== 'ollama-remote') {
      throw new Error('Clé API non configurée. Allez dans l\'onglet ⚙️ Config.');
    }

    const langInstructions = {
      'fr': 'Réponds en français.',
      'en': 'Answer in English.',
      'ar': 'أجب باللغة العربية.'
    };

    const systemPrompt = `Tu es un assistant expert qui analyse des documents. 
${langInstructions[cfg.responseLang] || langInstructions['fr']}

RÈGLES IMPORTANTES:
1. Base tes réponses UNIQUEMENT sur les documents fournis ci-dessous.
2. Pour chaque information, CITE LE NOM DU DOCUMENT SOURCE entre crochets [NomDuFichier].
3. Si la réponse ne se trouve dans aucun document, dis-le clairement.
4. Sois précis et structuré dans tes réponses.
5. Si plusieurs documents contiennent des infos pertinentes, cite-les tous.

═══════════════════════════════════════
DOCUMENTS CHARGÉS:
═══════════════════════════════════════
${context}
═══════════════════════════════════════`;

    const body = {
      model: cfg.modelName,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: question }
      ],
      max_tokens: cfg.maxTokens,
      temperature: 0.3,  // Basse température pour plus de précision factuelle
      stream: true        // Streaming pour réduire la latence perçue
    };

    const headers = {
      'Content-Type': 'application/json'
    };

    // Ajouter l'auth selon le fournisseur
    if (cfg.llmProvider === 'openrouter') {
      headers['Authorization'] = `Bearer ${cfg.apiKey}`;
      headers['HTTP-Referer'] = 'https://office-assistant-ia.local';
      headers['X-Title'] = 'Office AI Assistant';
    } else if (cfg.llmProvider === 'ollama-remote') {
      // Ollama n'a pas besoin d'auth par défaut
      if (cfg.apiKey) headers['Authorization'] = `Bearer ${cfg.apiKey}`;
      body.stream = false; // Ollama streaming format différent
    } else {
      headers['Authorization'] = `Bearer ${cfg.apiKey}`;
    }

    const response = await fetch(cfg.apiUrl, {
      method: 'POST',
      headers,
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Erreur API (${response.status}): ${errText}`);
    }

    // Si streaming activé
    if (body.stream && response.body) {
      return this.handleStream(response);
    }
    
    // Sinon, réponse complète
    const data = await response.json();
    return {
      text: data.choices?.[0]?.message?.content 
            || data.message?.content 
            || 'Aucune réponse reçue.',
      stream: false
    };
  },

  async handleStream(response) {
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    
    return {
      stream: true,
      async *[Symbol.asyncIterator]() {
        let buffer = '';
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';
          
          for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed || trimmed === 'data: [DONE]') continue;
            if (!trimmed.startsWith('data: ')) continue;
            
            try {
              const json = JSON.parse(trimmed.slice(6));
              const content = json.choices?.[0]?.delta?.content;
              if (content) yield content;
            } catch { /* skip malformed chunks */ }
          }
        }
      }
    };
  }
};

// ── Lecture du document Office ouvert ──
async function readCurrentOfficeDocument() {
  return new Promise((resolve, reject) => {
    if (!Office || !Office.context || !Office.context.document) {
      reject(new Error('Office API non disponible'));
      return;
    }

    const host = Office.context.host;
    
    if (host === Office.HostType.Word) {
      Word.run(async (context) => {
        const body = context.document.body;
        body.load('text');
        await context.sync();
        resolve({
          name: 'Document Word ouvert',
          type: 'docx',
          content: body.text
        });
      }).catch(reject);
    } 
    else if (host === Office.HostType.Excel) {
      Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();
        
        let fullText = '';
        for (const sheet of sheets.items) {
          const range = sheet.getUsedRange();
          range.load('values');
          try {
            await context.sync();
            fullText += `\n--- Feuille: ${sheet.name} ---\n`;
            fullText += range.values.map(row => row.join('\t')).join('\n');
          } catch {
            fullText += `\n--- Feuille: ${sheet.name} (vide) ---\n`;
          }
        }
        resolve({
          name: 'Classeur Excel ouvert',
          type: 'xlsx',
          content: fullText.trim()
        });
      }).catch(reject);
    }
    else if (host === Office.HostType.PowerPoint) {
      // L'API PowerPoint JS est plus limitée
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: Office.ValueFormat.Unformatted },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve({
              name: 'Présentation PowerPoint ouverte',
              type: 'pptx',
              content: result.value || '[Sélectionnez du texte dans la présentation]'
            });
          } else {
            // Essayer de lire tout le document
            Office.context.document.getFileAsync(Office.FileType.Text, (fileResult) => {
              if (fileResult.status === Office.AsyncResultStatus.Succeeded) {
                const file = fileResult.value;
                const sliceCount = file.sliceCount;
                let fullText = '';
                let slicesReceived = 0;
                
                for (let i = 0; i < sliceCount; i++) {
                  file.getSliceAsync(i, (sliceResult) => {
                    if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                      fullText += sliceResult.value.data;
                    }
                    slicesReceived++;
                    if (slicesReceived === sliceCount) {
                      file.closeAsync();
                      resolve({
                        name: 'Présentation PowerPoint ouverte',
                        type: 'pptx',
                        content: fullText || '[Contenu non extractible directement]'
                      });
                    }
                  });
                }
              } else {
                resolve({
                  name: 'Présentation PowerPoint ouverte',
                  type: 'pptx',
                  content: '[Utilisez le bouton charger fichier pour les PPT]'
                });
              }
            });
          }
        }
      );
    }
    else {
      reject(new Error('Type de document Office non reconnu'));
    }
  });
}

// ── Utilitaires ──
function getFileIcon(type) {
  const icons = {
    'pdf': '📕', 'docx': '📘', 'doc': '📘',
    'xlsx': '📗', 'xls': '📗', 'csv': '📗',
    'pptx': '📙', 'ppt': '📙',
    'txt': '📄', 'md': '📄'
  };
  return icons[type] || '📄';
}

function formatSize(charCount) {
  if (charCount < 1000) return `${charCount} car.`;
  if (charCount < 1000000) return `${(charCount / 1000).toFixed(1)}K car.`;
  return `${(charCount / 1000000).toFixed(1)}M car.`;
}

function setStatus(text, type = '') {
  const bar = document.getElementById('statusBar');
  bar.textContent = text;
  bar.style.background = type === 'error' ? '#dc2626' : 
                          type === 'success' ? '#166534' : '#1e293b';
}

function addChatMessage(role, content) {
  const container = document.getElementById('chatMessages');
  const div = document.createElement('div');
  div.className = `message ${role}`;
  div.innerHTML = formatMessageContent(content, role);
  container.appendChild(div);
  container.scrollTop = container.scrollHeight;
  return div;
}

function formatMessageContent(text, role) {
  if (role !== 'assistant') return escapeHtml(text);
  
  // Formatter la réponse de l'assistant
  let html = escapeHtml(text);
  
  // Mettre en évidence les références [NomDeFichier]
  html = html.replace(/\[([^\]]+)\]/g, '<span class="source-tag">📄 $1</span>');
  
  // Basique: gras **texte** et retours à la ligne
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/\n/g, '<br>');
  
  return html;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

async function loadScript(url) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

// ── Initialisation ──
Office.onReady((info) => {
  console.log('Office ready:', info.host, info.platform);
  initApp();
});

// Fallback si on teste hors d'Office
if (typeof Office === 'undefined' || !Office.onReady) {
  document.addEventListener('DOMContentLoaded', initApp);
}

function initApp() {
  // Charger la config
  Config.applyToUI();

  // Gestion des onglets
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(tc => tc.style.display = 'none');
      tab.classList.add('active');
      document.getElementById(`tab-${tab.dataset.tab}`).style.display = 'block';
    });
  });

  // Pré-remplir l'URL selon le provider
  document.getElementById('llmProvider').addEventListener('change', (e) => {
    const urlInput = document.getElementById('apiUrl');
    switch (e.target.value) {
      case 'openrouter':
        urlInput.value = 'https://openrouter.ai/api/v1/chat/completions';
        break;
      case 'ollama-remote':
        urlInput.value = 'http://localhost:11434/api/chat';
        break;
      case 'openai-compatible':
        urlInput.value = 'https://api.openai.com/v1/chat/completions';
        break;
    }
  });

  // Sauvegarder la config
  document.getElementById('btnSaveConfig').addEventListener('click', () => {
    const cfg = Config.readFromUI();
    Config.save(cfg);
    const status = document.getElementById('configStatus');
    status.textContent = '✅ Configuration sauvegardée !';
    status.className = 'status success';
    setTimeout(() => { status.textContent = ''; status.className = 'status'; }, 3000);
  });

  // Charger des fichiers locaux
  document.getElementById('btnLoadFiles').addEventListener('click', async () => {
    const files = document.getElementById('fileInput').files;
    if (files.length === 0) {
      document.getElementById('fileInput').click();
      return;
    }

    for (const file of files) {
      try {
        setStatus(`Extraction: ${file.name}...`);
        const content = await TextExtractor.extractFromFile(file);
        const ext = file.name.split('.').pop().toLowerCase();
        DocumentStore.add(file.name, ext, content);
        setStatus(`✅ ${file.name} chargé`, 'success');
      } catch (err) {
        console.error(err);
        setStatus(`❌ Erreur: ${file.name} - ${err.message}`, 'error');
      }
    }
    
    document.getElementById('fileInput').value = '';
    setStatus(`${DocumentStore.documents.length} document(s) chargé(s)`, 'success');
  });

  // Lire le document Office courant
  document.getElementById('btnReadCurrent').addEventListener('click', async () => {
    try {
      setStatus('Lecture du document ouvert...');
      const doc = await readCurrentOfficeDocument();
      DocumentStore.add(doc.name, doc.type, doc.content);
      setStatus(`✅ "${doc.name}" lu avec succès`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`❌ ${err.message}`, 'error');
    }
  });

  // Vider les documents
  document.getElementById('btnClearDocs').addEventListener('click', () => {
    DocumentStore.clear();
    setStatus('Documents supprimés');
  });

  // Poser une question
  document.getElementById('btnAsk').addEventListener('click', handleAskQuestion);
  document.getElementById('questionInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      handleAskQuestion();
    }
  });
}

async function handleAskQuestion() {
  const questionInput = document.getElementById('questionInput');
  const question = questionInput.value.trim();
  if (!question) return;

  const btnAsk = document.getElementById('btnAsk');
  btnAsk.disabled = true;

  // Ajouter le doc courant si demandé
  if (document.getElementById('chkIncludeCurrentDoc').checked) {
    try {
      const currentDoc = await readCurrentOfficeDocument();
      DocumentStore.add(currentDoc.name, currentDoc.type, currentDoc.content);
    } catch { /* pas grave si ça échoue */ }
  }

  // Vérifier qu'il y a des documents
  if (DocumentStore.documents.length === 0) {
    addChatMessage('error', 'Aucun document chargé. Allez dans l\'onglet 📁 pour ajouter des documents.');
    btnAsk.disabled = false;
    return;
  }

  // Afficher la question
  addChatMessage('user', question);
  questionInput.value = '';

  // Afficher l'indicateur de chargement
  const loadingMsg = addChatMessage('assistant', 'Réflexion en cours');
  loadingMsg.querySelector('.message') || loadingMsg;
  loadingMsg.classList.add('loading-dots');

  try {
    setStatus('Envoi au LLM...');
    const context = DocumentStore.buildContext();
    const result = await LLMClient.ask(question, context);

    if (result.stream) {
      // Streaming : afficher token par token
      loadingMsg.classList.remove('loading-dots');
      loadingMsg.innerHTML = '';
      let fullText = '';
      
      for await (const chunk of result) {
        fullText += chunk;
        loadingMsg.innerHTML = formatMessageContent(fullText, 'assistant');
        document.getElementById('chatMessages').scrollTop = 
          document.getElementById('chatMessages').scrollHeight;
      }
      setStatus('✅ Réponse complète', 'success');
    } else {
      // Réponse complète
      loadingMsg.classList.remove('loading-dots');
      loadingMsg.innerHTML = formatMessageContent(result.text, 'assistant');
      setStatus('✅ Réponse reçue', 'success');
    }
  } catch (err) {
    console.error(err);
    loadingMsg.classList.remove('loading-dots');
    loadingMsg.className = 'message error';
    loadingMsg.textContent = `❌ Erreur: ${err.message}`;
    setStatus(`Erreur: ${err.message}`, 'error');
  }

  btnAsk.disabled = false;
}
