# 🌟 ভাষা মিত্র - Bangla Language Assistant for Microsoft Word

AI-powered Bangla spell checker, grammar checker, and writing assistant powered by Google Gemini.

## ✨ Features

- ✅ **বানান পরীক্ষা** - Spell checking with suggestions
- ✅ **লেখার ভাব বিশ্লেষণ** - Tone analysis and improvements
- ✅ **সাধু-চলিত মিশ্রণ সনাক্তকরণ** - Detect and fix language style mixing
- ✅ **বিরাম চিহ্ন পরামর্শ** - Punctuation suggestions
- ✅ **শ্রুতিমধুরতা উন্নতি** - Euphony improvements
- ✅ **লেখার ধরন বিশ্লেষণ** - Content type analysis
- ✅ **অনুপস্থিত উপাদান চিহ্নিত** - Missing elements detection
- ✅ **উন্নতির পরামর্শ** - Improvement suggestions

## 🚀 Local Development

### Prerequisites

- Node.js 16+
- Microsoft Word (Desktop)

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/youngsunson/version3.git
   cd bhashamitra
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Install development certificates**
   ```bash
   npx office-addin-dev-certs install --machine
   ```
   Click "Yes" when prompted.

4. **Start development server**
   ```bash
   npm run dev
   ```

5. **Load add-in in Word**
   
   **Option A: Automatic (Recommended)**
   ```bash
   npm run start
   ```
   
   **Option B: Manual**
   - Open Microsoft Word
   - Go to **Insert** → **Add-ins** → **Get Add-ins**
   - Click **MY ADD-INS** → **Upload My Add-in**
   - Select `manifest-dev.xml`
   - Click **Upload**

6. **Use the add-in**
   - Go to **Home** tab in Word
   - Click **"বানান পরীক্ষক"** button
   - Enter your Google Gemini API Key in settings
   - Start checking your Bangla text!

## 📦 Production Deployment (GitHub Pages)

### 1. Install gh-pages
```bash
npm install
```

### 2. Build and Deploy
```bash
npm run deploy
```

This will:
- Build the project to `dist/` folder
- Deploy to `gh-pages` branch
- Make it available at: `https://youngsunson.github.io/version3/`

### 3. GitHub Pages Settings

1. Go to your repository on GitHub
2. **Settings** → **Pages**
3. **Source**: Deploy from a branch
4. **Branch**: Select `gh-pages`
5. **Folder**: `/ (root)`
6. Click **Save**

Wait 2-3 minutes for deployment.

### 4. Load Production Add-in in Word

Use the production `manifest.xml` file:
- **Insert** → **Add-ins** → **Upload My Add-in**
- Select `manifest.xml` (NOT manifest-dev.xml)
- The add-in will load from GitHub Pages

## 🔑 Getting Google Gemini API Key

1. Visit [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Click **"Create API Key"**
3. Copy the API key
4. Paste it in the add-in settings (⚙️ icon)

## 📝 Available Scripts

- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm run deploy` - Build and deploy to GitHub Pages
- `npm run start` - Load add-in in Word (development)
- `npm run start:prod` - Load add-in in Word (production)
- `npm run validate` - Validate production manifest
- `npm run validate:dev` - Validate development manifest

## 📂 Project Structure

```
bhashamitra/
├── src/
│   ├── index.tsx       # Main React application
│   └── index.css       # Styles
├── public/
│   └── assets/         # Icons
├── manifest.xml        # Production manifest (GitHub Pages)
├── manifest-dev.xml    # Development manifest (localhost)
├── package.json        # Dependencies and scripts
├── vite.config.ts      # Vite configuration
├── tsconfig.json       # TypeScript configuration
└── README.md           # This file
```

## 🤝 AI Models Supported

- **Gemini 2.5 Flash** - Latest and best (Recommended)
- **Gemini 2.0 Flash** - New and fastest
- **Gemini 1.5 Pro** - Best quality
- **Gemini 1.5 Flash** - Fast
- **Gemini Pro** - Standard

## 📄 License

MIT License - See LICENSE file for details

## 👨‍💻 Author

**Bhasha Mitra Team**

## 🐛 Issues & Support

For issues and support, please visit: [GitHub Issues](https://github.com/youngsunson/version3/issues)

---

Made with ❤️ for Bangla language lovers
