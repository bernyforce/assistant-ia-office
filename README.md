# 🤖 Assistant IA Documents — Office 365 Add-in

> Interrogez vos documents locaux (Word, PDF, Excel, PowerPoint) directement
> depuis Office 365 grâce à un LLM de votre choix (Kimi K2.5, Claude, GPT-4o, 
> Llama, DeepSeek, etc.)

---

## 📋 Table des matières

- [Présentation](#-présentation)
- [Fonctionnalités](#-fonctionnalités)
- [Prérequis](#-prérequis)
- [Architecture](#-architecture)
- [Installation](#-installation)
  - [Méthode 1 — Serveur local (Développement)](#méthode-1--serveur-local-développement)
  - [Méthode 2 — GitHub Pages (Production gratuite)](#méthode-2--github-pages-production-gratuite)
  - [Méthode 3 — Script Lab (Zéro installation)](#méthode-3--script-lab-zéro-installation)
- [Configuration](#-configuration)
  - [Fournisseurs LLM supportés](#fournisseurs-llm-supportés)
  - [Obtenir une clé API](#obtenir-une-clé-api)
- [Utilisation](#-utilisation)
  - [Charger des documents](#1-charger-des-documents)
  - [Poser des questions](#2-poser-des-questions)
  - [Bonnes pratiques de prompting](#3-bonnes-pratiques-de-prompting)
- [Structure du projet](#-structure-du-projet)
- [Personnalisation](#-personnalisation)
- [Limites connues](#-limites-connues)
- [Dépannage](#-dépannage)
- [Licence](#-licence)

---

## 🎯 Présentation

**Assistant IA Documents** est un complément Office 365 (Add-in) qui s'intègre
nativement dans **Word**, **Excel** et **PowerPoint**. Il vous permet de :

- Charger vos cours, notes, documents locaux (PDF, DOCX, XLSX, PPTX, TXT, CSV)
- Les interroger en langage naturel via un LLM en ligne
- Obtenir des réponses sourcées indiquant dans quel fichier chaque information
  a été trouvée

**Aucune installation logicielle requise** sur le poste. Vous ajoutez simplement
le complément à votre session Office 365 et renseignez une clé API.

---

## ✨ Fonctionnalités

| Fonctionnalité | Description |
|---|---|
| 📄 Multi-format | PDF, DOCX, XLSX, PPTX, TXT, MD, CSV |
| 📖 Lecture doc ouvert | Lit le contenu du document actuellement ouvert dans Office |
| 🔌 Multi-LLM | OpenRouter, Ollama distant, tout endpoint compatible OpenAI |
| ⚡ Streaming | Réponses affichées en temps réel, token par token |
| 📎 Sources citées | Chaque réponse indique le fichier source `[NomDuFichier]` |
| 🌍 Multilingue | Réponses en français, anglais ou arabe |
| 💾 Config persistante | Sauvegarde locale via `localStorage` |
| 🎨 Interface soignée | Task Pane avec onglets Chat / Documents / Config |
| 🔒 Vie privée | Aucun serveur tiers — communication directe navigateur ↔ API LLM |

---

## 📦 Prérequis

- **Office 365** (Word, Excel ou PowerPoint) — version desktop ou web
- **Un compte API LLM** avec clé (voir [Configuration](#-configuration))
- Pour l'installation locale uniquement : **Node.js 18+** (optionnel)

---

## 🏗 Architecture

Copy
┌──────────────────────────────────────────────────┐ │ Office 365 (Word / Excel / PPT) │ │ │ │ ┌──────────────────────────────────────────┐ │ │ │ Task Pane (HTML/JS/CSS) │ │ │ │ │ │ │ │ ┌──────────┐ ┌──────────┐ ┌──────────┐ │ │ │ │ │ Config │ │Documents │ │ Chat │ │ │ │ │ │ API Key │ │ Upload │ │ Question │ │ │ │ │ │ Modèle │ │ Parse │ │ Réponse │ │ │ │ │ └──────────┘ └────┬─────┘ └────┬─────┘ │ │ │ │ │ │ │ │ │ │ DocumentStore │ │ │ │ │ (mémoire JS) │ │ │ │ │ │ │ │ │ │ │ └─────┬──────┘ │ │ │ │ │ │ │ │ │ Contexte + Question │ │ │ │ │ │ │ │ └──────────────────────────┼───────────────┘ │ │ │ │ └──────────────────────────────┼───────────────────┘ │ HTTPS ▼ ┌─────────────────────┐ │ API LLM │ │ (OpenRouter / │ │ Ollama / autre) │ └─────────────────────┘


**Flux de données :**

1. L'utilisateur charge des fichiers via l'interface (File API du navigateur)
2. Les bibliothèques JS côté client extraient le texte (pdf.js, mammoth.js,
   SheetJS, JSZip)
3. Le texte extrait est stocké en mémoire JavaScript (`DocumentStore`)
4. Quand l'utilisateur pose une question, le script construit un prompt
   contenant le contexte documentaire + la question
5. Le prompt est envoyé directement à l'API LLM via `fetch()`
6. La réponse (en streaming) est affichée dans le chat

---

## 🚀 Installation

### Méthode 1 — Serveur local (Développement)

Idéale pour tester et développer. Nécessite Node.js.

```bash
# Cloner ou créer le dossier du projet
mkdir assistant-ia-office && cd assistant-ia-office

# Placer les 4 fichiers : manifest.xml, index.html, style.css, app.js

# Installer les certificats SSL de développement Office
npx office-addin-dev-certs install

# Lancer le serveur HTTPS local
npx http-server . -S \
  -C ~/.office-addin-dev-certs/localhost.crt \
  -K ~/.office-addin-dev-certs/localhost.key \
  -p 3000

# Le serveur tourne sur https://localhost:3000
Chargement dans Office (Sideloading) :

Windows :

1. Ouvrir Word, Excel ou PowerPoint
2. Fichier → Options → Centre de gestion de la confidentialité
   → Paramètres → Catalogues de compléments approuvés
3. Ajouter le chemin du dossier contenant manifest.xml
   Exemple : \\%USERPROFILE%\assistant-ia-office
4. Cocher "Afficher dans le menu"
5. Redémarrer Office
6. Insertion → Mes compléments → Dossier partagé → Assistant IA Documents
macOS :

1. Copier manifest.xml dans :
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   (remplacer Word par Excel ou Powerpoint selon l'application)
2. Redémarrer l'application Office
3. Insertion → Compléments → Mes compléments → Assistant IA Documents
Office Web (navigateur) :

1. Ouvrir un document dans Office Online (office.com)
2. Insertion → Compléments → Charger mon complément
3. Sélectionner le fichier manifest.xml
4. Le Task Pane s'ouvre automatiquement
Méthode 2 — GitHub Pages (Production gratuite)
Pas besoin de serveur local permanent. Hébergement gratuit sur GitHub.

Copy# 1. Créer un dépôt GitHub (public ou privé avec Pages activé)
gh repo create assistant-ia-office --public

# 2. Pousser les fichiers
git init
git add manifest.xml index.html style.css app.js
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/VOTRE-USER/assistant-ia-office.git
git push -u origin main

# 3. Activer GitHub Pages
#    Settings → Pages → Source: "main" → Dossier: / (root) → Save
#    URL obtenue : https://VOTRE-USER.github.io/assistant-ia-office/
Mise à jour du manifest :

Ouvrir manifest.xml et remplacer toutes les occurrences de https://localhost:3000 par votre URL GitHub Pages :

Copy<!-- AVANT -->
<SourceLocation DefaultValue="https://localhost:3000/index.html"/>

<!-- APRÈS -->
<SourceLocation DefaultValue="https://VOTRE-USER.github.io/assistant-ia-office/index.html"/>
Puis sideloader le manifest mis à jour dans Office (même procédure que Méthode 1).

Méthode 3 — Script Lab (Zéro installation)
Si vous ne voulez absolument rien installer ni héberger.

1. Ouvrir Word, Excel ou PowerPoint (desktop ou web)
2. Insertion → Compléments → Store
3. Chercher "Script Lab" → Installer
4. Ouvrir Script Lab → Code
5. Créer un nouveau snippet
6. Coller le contenu de index.html (onglet HTML), 
   style.css (onglet CSS) et app.js (onglet Script)
7. Cliquer "Run" pour exécuter
Note : Script Lab est limité en terme d'interface. Les méthodes 1 et 2 offrent une bien meilleure expérience utilisateur.

⚙️ Configuration
Au premier lancement, ouvrez l'onglet ⚙️ Config dans le Task Pane.

Fournisseurs LLM supportés
Fournisseur	URL API	Clé requise	Modèles recommandés
OpenRouter	https://openrouter.ai/api/v1/chat/completions	Oui	moonshotai/kimi-k2.5, anthropic/claude-sonnet-4, google/gemini-2.5-pro
Ollama (distant)	http://VOTRE-IP:11434/api/chat	Non*	llama3.3, mistral, qwen2.5
OpenAI	https://api.openai.com/v1/chat/completions	Oui	gpt-4o, gpt-4o-mini
Mistral AI	https://api.mistral.ai/v1/chat/completions	Oui	mistral-large-latest
DeepSeek	https://api.deepseek.com/v1/chat/completions	Oui	deepseek-chat
Tout endpoint compatible OpenAI	Votre URL	Selon config	—
*Ollama n'exige pas de clé par défaut. Si vous exposez Ollama sur internet, configurez une authentification.

Obtenir une clé API
OpenRouter (recommandé — accès à 200+ modèles) :

1. Aller sur https://openrouter.ai
2. Créer un compte
3. Dashboard → API Keys → Create Key
4. Copier la clé (commence par sk-or-v1-...)
5. Coller dans l'onglet Config du add-in
Tarif indicatif OpenRouter : Kimi K2.5 ≈ $0.60/M tokens en entrée, soit environ 0,01€ par question avec un document de 10 pages.

📖 Utilisation
1. Charger des documents
Via le sélecteur de fichiers :

Onglet 📁 Documents → Cliquer sur "Charger les fichiers"
→ Sélectionner un ou plusieurs fichiers
→ Les fichiers apparaissent dans la liste avec leur taille en caractères
Via le document Office ouvert :

Onglet 📁 Documents → Cliquer sur "Lire doc ouvert"
→ Le contenu du document actif est ajouté automatiquement
Formats supportés :

Format	Extension	Bibliothèque utilisée
PDF	.pdf	pdf.js (Mozilla)
Word	.docx, .doc	mammoth.js
Excel	.xlsx, .xls, .csv	SheetJS
PowerPoint	.pptx, .ppt	JSZip + parsing XML
Texte	.txt, .md	API File native
2. Poser des questions
Onglet 💬 Chat → Écrire votre question → Cliquer "Envoyer" (ou Ctrl+Entrée)
La case "Inclure doc ouvert" est cochée par défaut : le contenu du document actuellement ouvert dans Office sera automatiquement ajouté au contexte à chaque question.

3. Bonnes pratiques de prompting
Soyez spécifique sur ce que vous cherchez :

❌  "Résume mes documents"
✅  "Résume les points clés du chapitre sur la photosynthèse 
     dans mes cours de biologie"
Demandez explicitement les sources :

❌  "Quelle est la définition du PIB ?"
✅  "Quelle est la définition du PIB donnée dans mes documents ? 
     Cite le fichier source exact."
Comparez entre documents :

✅  "Compare les définitions de la démocratie entre le cours de 
     droit constitutionnel et le cours de science politique. 
     Y a-t-il des différences ?"
Posez des questions de synthèse :

✅  "En me basant sur l'ensemble de mes notes de cours, quels sont 
     les 5 thèmes les plus importants à réviser pour l'examen ? 
     Pour chaque thème, indique dans quel fichier j'ai des notes."
Demandez des explications :

✅  "L'exercice 3 du TD de mathématiques utilise la méthode de Gauss. 
     Explique-moi cette méthode en te basant sur mon cours, puis 
     montre-moi comment l'appliquer à cet exercice."
Posez des questions en chaîne :

Message 1 : "Liste les chapitres couverts dans mon cours d'économie."
Message 2 : "Détaille le chapitre 3 sur la politique monétaire."
Message 3 : "Quels exercices dans mes TDs portent sur ce chapitre ?"
📁 Structure du projet
📁 assistant-ia-office/
│
├── 📄 manifest.xml      ← Déclaration du add-in Office
│                           Définit les métadonnées, les hôtes (Word, Excel, PPT),
│                           les boutons du ruban et l'URL du Task Pane.
│
├── 📄 index.html         ← Interface utilisateur du Task Pane
│                           Structure HTML avec 3 onglets : Chat, Documents, Config.
│                           Charge les dépendances CDN (Office.js, pdf.js, etc.)
│
├── 📄 style.css          ← Styles de l'interface
│                           Design responsive adapté au Task Pane étroit d'Office.
│
├── 📄 app.js             ← Logique applicative principale
│                           - DocumentStore : stockage en mémoire des documents
│                           - Config : gestion de la configuration (localStorage)
│                           - TextExtractor : extraction de texte multi-format
│                           - LLMClient : communication avec l'API (streaming)
│                           - Intégration Office JS API (lecture doc ouvert)
│
└── 📄 README.md          ← Ce fichier
🔧 Personnalisation
Modifier la taille maximale des documents
Dans app.js, fonction DocumentStore.buildContext() :

Copyconst maxChars = 15000; // Modifier cette valeur
// 15000 chars ≈ 4000 tokens ≈ ~6 pages
// 50000 chars ≈ 13000 tokens ≈ ~20 pages
// Pour Kimi K2.5 (128K contexte) : vous pouvez monter à 100000+
Modifier le prompt système
Dans app.js, fonction LLMClient.ask(), modifiez la variable systemPrompt pour changer le comportement de l'assistant (langue, style, instructions).

Ajouter un nouveau fournisseur LLM
Ajouter une option dans le <select id="llmProvider"> de index.html
Ajouter le cas dans le switch de l'événement change du provider dans app.js
Ajouter les headers spécifiques dans LLMClient.ask()
Changer les icônes du ruban
Remplacez les URLs Icon16, Icon32, Icon80 dans manifest.xml par vos propres images PNG (16x16, 32x32, 80x80 pixels).

⚠️ Limites connues
Limite	Détail	Contournement
PDF scannés	Les PDF basés image (scans) ne sont pas lisibles par pdf.js	Utiliser un outil OCR externe pour convertir en PDF texte
Taille du contexte	Dépend du modèle LLM (8K à 128K+ tokens)	Ajuster maxChars dans app.js ou utiliser un modèle à grande fenêtre
Fichiers .doc anciens	Le format binaire .doc (pre-2007) n'est pas toujours bien lu	Convertir en .docx via Office avant de charger
PowerPoint via Office API	L'API JS de PowerPoint est limitée pour la lecture de texte	Charger le fichier .pptx via le bouton Upload plutôt que "Lire doc ouvert"
CORS	Certaines API peuvent bloquer les appels depuis un add-in Office	OpenRouter et les principaux fournisseurs acceptent les appels cross-origin
Pas de mémoire entre sessions	Les documents chargés sont perdus à la fermeture du Task Pane	Recharger les fichiers à chaque session
Pas d'historique de conversation côté LLM	Chaque question est indépendante (pas de suivi multi-tours)	Reformuler les questions de manière autonome
🔍 Dépannage
Le complément n'apparaît pas dans Office :

Vérifiez que le serveur HTTPS est lancé et accessible
Vérifiez que le certificat SSL est approuvé par votre système
Sur macOS, vérifiez le chemin exact du dossier wef
Redémarrez complètement Office (pas juste le document)
Erreur "Failed to fetch" ou erreur réseau :

Vérifiez l'URL de l'API dans la configuration
Vérifiez que la clé API est correcte et active
Testez l'API dans un outil comme curl ou Postman
Si Ollama : vérifiez que le serveur est accessible et que CORS est activé (OLLAMA_ORIGINS=*)
Les PDF ne sont pas lus :

Vérifiez que le PDF contient du texte (pas un scan)
Essayez d'ouvrir le PDF dans un navigateur et de sélectionner le texte
Si aucun texte n'est sélectionnable, le PDF est une image
La réponse est tronquée :

Augmentez la valeur "Tokens max" dans l'onglet Config
Reformulez la question pour qu'elle soit plus ciblée
Latence élevée :

Le streaming est activé par défaut pour réduire la latence perçue
Choisissez un modèle plus rapide (Kimi K2.5 et GPT-4o-mini sont rapides)
Réduisez le nombre de documents chargés
Réduisez maxChars dans app.js
📊 Coûts estimatifs
En utilisant OpenRouter comme passerelle :

Usage	Tokens estimés	Coût approximatif
1 question + 5 pages de contexte	~3 000 tokens	~$0.002
1 question + 20 pages de contexte	~10 000 tokens	~$0.006
1 question + 50 pages de contexte	~30 000 tokens	~$0.018
Session de révision (20 questions)	~100 000 tokens	~$0.060
Basé sur les tarifs de Kimi K2.5 via OpenRouter. Les tarifs varient selon le modèle choisi.

🛡️ Sécurité et vie privée
Aucun serveur intermédiaire : vos documents transitent directement de votre navigateur vers l'API LLM choisie
Aucun stockage distant : les documents restent en mémoire JavaScript locale et disparaissent à la fermeture
Clé API en local : stockée dans localStorage de votre navigateur, jamais transmise à un tiers autre que le fournisseur LLM
Vous contrôlez le fournisseur : choisissez un fournisseur dont vous acceptez la politique de confidentialité
⚠️ Attention : les contenus envoyés à l'API LLM sont traités selon la politique du fournisseur choisi. Pour des documents sensibles, privilégiez Ollama auto-hébergé ou un fournisseur avec engagement de non-rétention des données.

📄 Licence
MIT License — Libre d'utilisation, modification et distribution.

MIT License

Copyright (c) 2026

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
Fait avec 🧠 pour les étudiants et professionnels qui veulent exploiter leurs documents intelligemment.


Ce README couvre l'intégralité du projet : présentation, installation détaillée pour chaque plateforme et méthode, configuration de chaque fournisseur LLM, guide d'utilisation avec les meilleures pratiques de prompting, personnalisation, dépannage, et considérations de sécurité. Il est prêt à être déposé tel quel dans ton dépôt.