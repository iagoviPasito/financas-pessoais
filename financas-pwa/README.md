# 📱 Controle Financeiro — Deploy no Vercel

## Passo a passo pelo iPhone (10 minutos)

### 1. Criar conta no Vercel
- Acesse **vercel.com** no Safari
- Clique em "Sign Up"
- Use "Continue with GitHub" (crie uma conta GitHub gratuita se não tiver)

### 2. Fazer upload do projeto
- No Vercel, clique em **"Add New Project"**
- Escolha **"Browse"** para upload manual
- Selecione a pasta `financas-pwa` (ou o arquivo ZIP)
- O Vercel detecta automaticamente que é React

### 3. Configurar o build
Vercel preenche automaticamente:
- Framework: **Create React App**
- Build Command: `npm run build`
- Output Directory: `build`

Clique em **Deploy** e aguarde ~2 minutos.

### 4. Instalar no iPhone
- Após o deploy, você recebe um link tipo `financas-pessoais.vercel.app`
- Abra esse link no **Safari** (não no Chrome — só o Safari permite instalar PWA no iPhone)
- Toque no ícone de **compartilhar** (quadrado com seta ↑)
- Role para baixo e toque em **"Adicionar à Tela de Início"**
- Confirme o nome e toque em **"Adicionar"**

✅ Pronto! O app aparece na sua home com ícone dourado.

---

## Observações

- **Dados salvos localmente**: as transações ficam no armazenamento do iPhone (localStorage), não sobem pra nuvem. Se limpar o Safari, os dados são apagados. Sempre exporte o Excel como backup.
- **Voz**: use o botão 🎙 na aba Lançamentos. O Safari pedirá permissão ao microfone na primeira vez.
- **Offline**: depois do primeiro acesso, o app funciona sem internet (exceto a função de voz que precisa de IA).

---

## Estrutura do projeto

```
financas-pwa/
├── public/
│   ├── index.html        ← Metatags iOS PWA
│   ├── manifest.json     ← Configuração do PWA
│   ├── service-worker.js ← Cache offline
│   ├── icon-192.png      ← Ícone do app
│   └── icon-512.png      ← Ícone splash screen
├── src/
│   ├── index.js          ← Entry point
│   └── App.jsx           ← App completo
├── package.json
└── vercel.json
```
