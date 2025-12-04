import { useState, useEffect } from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';

// Types
interface SpellingError {
  wrong: string;
  suggestions: string[];
  position?: number;
}

interface ToneImprovement {
  current: string;
  suggestions: string[];
  reason: string;
}

interface LanguageStyleMixing {
  detected: boolean;
  recommendedStyle?: string;
  reason?: string;
  corrections?: Array<{
    current: string;
    suggestion: string;
    type: string;
  }>;
}

interface PunctuationIssue {
  issue: string;
  currentSentence: string;
  correctedSentence: string;
  explanation: string;
}

interface EuphonyImprovement {
  current: string;
  suggestions: string[];
  reason: string;
}

interface ContentAnalysis {
  contentType: string;
  description?: string;
  missingElements?: string[];
  suggestions?: string[];
}

// Main App Component
function App() {
  const [apiKey, setApiKey] = useState(localStorage.getItem('gemini_api_key') || '');
  const [selectedModel, setSelectedModel] = useState(localStorage.getItem('gemini_model') || 'gemini-2.5-flash');
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  const [showInstructions, setShowInstructions] = useState(false);
  
  const [corrections, setCorrections] = useState<SpellingError[]>([]);
  const [toneImprovements, setToneImprovements] = useState<ToneImprovement[]>([]);
  const [languageStyleMixing, setLanguageStyleMixing] = useState<LanguageStyleMixing | null>(null);
  const [punctuationIssues, setPunctuationIssues] = useState<PunctuationIssue[]>([]);
  const [euphonyImprovements, setEuphonyImprovements] = useState<EuphonyImprovement[]>([]);
  const [contentAnalysis, setContentAnalysis] = useState<ContentAnalysis | null>(null);
  
  const [stats, setStats] = useState({ totalWords: 0, errorCount: 0, accuracy: 100 });
  const [currentText, setCurrentText] = useState('');

  useEffect(() => {
    Office.onReady(() => {
      console.log('Office is ready!');
    });
  }, []);

  const showMessage = (text: string, type: 'success' | 'error') => {
    setMessage({ text, type });
    setTimeout(() => setMessage(null), 3000);
  };

  const saveSettings = () => {
    if (apiKey) {
      localStorage.setItem('gemini_api_key', apiKey);
    }
    localStorage.setItem('gemini_model', selectedModel);
    showMessage('সেটিংস সংরক্ষিত হয়েছে! ✓', 'success');
    setShowSettings(false);
  };

  const getTextFromWord = async (): Promise<string> => {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        const body = context.document.body;
        body.load('text');
        await context.sync();
        resolve(body.text);
      }).catch((error) => {
        console.error('Error reading from Word:', error);
        resolve('');
      });
    });
  };

  const highlightTextInWord = async (searchText: string, highlightColor: string) => {
    await Word.run(async (context) => {
      const searchResults = context.document.body.search(searchText, { matchCase: true, matchWholeWord: false });
      searchResults.load('font');
      await context.sync();

      for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = highlightColor;
      }
      await context.sync();
    }).catch((error) => {
      console.error('Error highlighting in Word:', error);
    });
  };

  const scrollToTextInWord = async (searchText: string) => {
    await Word.run(async (context) => {
      const searchResults = context.document.body.search(searchText, { matchCase: true, matchWholeWord: false });
      searchResults.load();
      await context.sync();

      if (searchResults.items.length > 0) {
        searchResults.items[0].select();
      }
    }).catch((error) => {
      console.error('Error scrolling in Word:', error);
    });
  };

  const replaceTextInWord = async (searchText: string, replaceText: string) => {
    await Word.run(async (context) => {
      const searchResults = context.document.body.search(searchText, { matchCase: false, matchWholeWord: false });
      searchResults.load();
      await context.sync();

      searchResults.items.forEach((item) => {
        item.insertText(replaceText, Word.InsertLocation.replace);
      });
      await context.sync();
    }).catch((error) => {
      console.error('Error replacing in Word:', error);
    });
  };

  const clearHighlights = async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.font.highlightColor = "none"; // ✅ FIXED: was null → now "none"
      await context.sync();
    }).catch((error) => {
      console.error('Error clearing highlights:', error);
    });
  };

  const checkSpelling = async () => {
    if (!apiKey) {
      showMessage('অনুগ্রহ করে প্রথমে API Key দিন এবং সংরক্ষণ করুন', 'error');
      return;
    }

    const text = await getTextFromWord();
    
    if (!text || text.trim().length === 0) {
      showMessage('অনুগ্রহ করে Word document এ কিছু টেক্সট লিখুন', 'error');
      return;
    }

    setCurrentText(text);
    setIsLoading(true);

    try {
      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{
              parts: [{
                text: `আপনি একজন বাংলা ভাষা বিশেষজ্ঞ এবং সাহিত্যিক। নিচের বাংলা টেক্সট সম্পূর্ণভাবে বিশ্লেষণ করুন।

টেক্সট: "${text}"

নির্দেশনা:

1. **বানান ভুল** চিহ্নিত করুন

2. **Tone/ভাব** চিহ্নিত করুন এবং সেই অনুযায়ী শব্দ পরিবর্তনের পরামর্শ দিন

3. **সাধু-চলিত ভাষা মিশ্রণ** পরীক্ষা করুন:
   - লেখায় যদি সাধু ও চলিত উভয় রীতির শব্দ মিশ্রিত থাকে
   - কোন একটি রীতিতে লিখলে ভালো হবে তা নির্ধারণ করুন (লেখার ধরন অনুযায়ী)
   - মিশ্রিত শব্দগুলো শনাক্ত করে সঠিক রীতিতে পরিবর্তনের সাজেশন দিন

4. **বিরাম চিহ্ন সমস্যা** খুঁজে বের করুন:
   - যেখানে বিরাম চিহ্ন প্রয়োজন কিন্তু নেই
   - যেখানে ভুল বিরাম চিহ্ন ব্যবহার হয়েছে

5. **শ্রুতিমধুরতা (Euphony)** উন্নতি:
   - একই শব্দের পুনরাবৃত্তি এড়াতে সমার্থক শব্দ
   - শব্দ চয়ন যা বাক্যকে আরো সুন্দর করবে

Response format (শুধুমাত্র valid JSON object return করুন):
{
  "spellingErrors": [
    {
      "wrong": "ভুল শব্দ",
      "suggestions": ["সঠিক বানান ১", "সঠিক বানান ২"],
      "position": index
    }
  ],
  "toneImprovements": [
    {
      "current": "বর্তমান শব্দ",
      "suggestions": ["ভাব অনুযায়ী শব্দ ১"],
      "reason": "কেন পরিবর্তন করা উচিত"
    }
  ],
  "languageStyleMixing": {
    "detected": true/false,
    "recommendedStyle": "সাধু রীতি" অথবা "চলিত রীতি",
    "reason": "কেন এই রীতি প্রস্তাবিত",
    "corrections": [
      {
        "current": "বর্তমান শব্দ",
        "suggestion": "প্রস্তাবিত রীতি অনুযায়ী",
        "type": "সাধু→চলিত"
      }
    ]
  },
  "punctuationIssues": [
    {
      "issue": "সমস্যার বর্ণনা",
      "currentSentence": "বর্তমান বাক্য",
      "correctedSentence": "সংশোধিত বাক্য",
      "explanation": "কেন এই পরিবর্তন প্রয়োজন"
    }
  ],
  "euphonyImprovements": [
    {
      "current": "বর্তমান শব্দ",
      "suggestions": ["মধুর বিকল্প ১"],
      "reason": "কেন এই পরিবর্তন লেখাকে আরো সুন্দর করবে"
    }
  ]
}

যদি কোন ভুল/পরামর্শ না থাকে, তাহলে খালি array বা false দিন।`
              }]
            }]
          })
        }
      );

      if (!response.ok) throw new Error('API request failed');

      const data = await response.json();
      const result = data.candidates[0].content.parts[0].text;
      const jsonMatch = result.match(/\{[\s\S]*\}/);

      if (jsonMatch) {
        const analysisData = JSON.parse(jsonMatch[0]);
        
        const spellingErrors = analysisData.spellingErrors || [];
        spellingErrors.forEach((error: SpellingError) => { // ✅ FIXED: removed unused 'index'
          if (error.position === undefined) {
            error.position = text.indexOf(error.wrong);
          }
        });
        spellingErrors.sort((a: SpellingError, b: SpellingError) => 
          (a.position || 0) - (b.position || 0)
        );

        setCorrections(spellingErrors);
        setToneImprovements(analysisData.toneImprovements || []);
        setLanguageStyleMixing(analysisData.languageStyleMixing || null);
        setPunctuationIssues(analysisData.punctuationIssues || []);
        setEuphonyImprovements(analysisData.euphonyImprovements || []);

        // Highlight errors in Word
        await clearHighlights();
        for (const error of spellingErrors) {
          await highlightTextInWord(error.wrong, '#fee2e2');
        }
        for (const improvement of (analysisData.toneImprovements || [])) {
          await highlightTextInWord(improvement.current, '#dbeafe');
        }
        if (analysisData.languageStyleMixing?.corrections) {
          for (const correction of analysisData.languageStyleMixing.corrections) {
            await highlightTextInWord(correction.current, '#e9d5ff');
          }
        }
        for (const improvement of (analysisData.euphonyImprovements || [])) {
          await highlightTextInWord(improvement.current, '#fce7f3');
        }

        updateStats(text, spellingErrors);
      }

      // Perform content analysis
      await analyzeContent(text);
      
    } catch (error) {
      console.error('Error:', error);
      showMessage('একটি ত্রুটি হয়েছে। অনুগ্রহ করে API Key যাচাই করুন এবং আবার চেষ্টা করুন।', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const analyzeContent = async (text: string) => {
    try {
      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{
              parts: [{
                text: `আপনি একজন বাংলা সাহিত্য ও লেখনী বিশেষজ্ঞ। নিচের বাংলা লেখাটি বিশ্লেষণ করুন এবং তথ্য দিন।

টেক্সট: "${text}"

নির্দেশনা:
1. এই লেখাটি কি ধরনের (যেমন: চিঠি, আবেদন, প্রবন্ধ, গল্প, কবিতা ইত্যাদি)
2. এই ধরনের লেখায় সাধারণত কি কি উপাদান থাকা উচিত যা এই লেখায় নেই
3. লেখাটি আরও ভালো করতে কি কি পরামর্শ দেওয়া যায়

Response format (শুধুমাত্র valid JSON object return করুন):
{
  "contentType": "লেখার ধরন",
  "description": "এই ধরনের লেখার সংক্ষিপ্ত বর্ণনা",
  "missingElements": ["অনুপস্থিত উপাদান ১"],
  "suggestions": ["উন্নতির পরামর্শ ১"]
}`
              }]
            }]
          })
        }
      );

      if (response.ok) {
        const data = await response.json();
        const result = data.candidates[0].content.parts[0].text;
        const jsonMatch = result.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          setContentAnalysis(JSON.parse(jsonMatch[0]));
        }
      }
    } catch (error) {
      console.error('Content analysis error:', error);
    }
  };

  const updateStats = (text: string, errors: SpellingError[]) => {
    const words = text.trim().split(/\s+/).filter(w => w.length > 0);
    const totalWords = words.length;
    const errorCount = errors.length;
    const accuracy = totalWords > 0 ? Math.round(((totalWords - errorCount) / totalWords) * 100) : 100;
    setStats({ totalWords, errorCount, accuracy });
  };

  const handleReplace = async (wrongWord: string, correctWord: string) => {
    await replaceTextInWord(wrongWord, correctWord);
    
    setCorrections(prev => prev.filter(c => c.wrong !== wrongWord));
    setToneImprovements(prev => prev.filter(t => t.current !== wrongWord));
    setEuphonyImprovements(prev => prev.filter(e => e.current !== wrongWord));
    
    if (languageStyleMixing?.corrections) {
      const filtered = languageStyleMixing.corrections.filter(c => c.current !== wrongWord);
      if (filtered.length === 0) {
        setLanguageStyleMixing(null);
      } else {
        setLanguageStyleMixing({ ...languageStyleMixing, corrections: filtered });
      }
    }

    const newText = currentText.replace(new RegExp(wrongWord, 'g'), correctWord);
    setCurrentText(newText);
    updateStats(newText, corrections.filter(c => c.wrong !== wrongWord));
    
    showMessage(`"${wrongWord}" কে "${correctWord}" দিয়ে প্রতিস্থাপন করা হয়েছে`, 'success');
  };

  const handleHoverWord = async (word: string) => {
    await scrollToTextInWord(word);
  };

  // ----------------------------------------------------------------------
  // RENDER METHOD (Fixed Layout for Footer)
  // ----------------------------------------------------------------------
  return (
    // Root Container: Full Viewport Height, No Outer Scroll
    <div style={{ 
      fontFamily: "'Noto Sans Bengali', sans-serif", 
      background: 'linear-gradient(to bottom right, #eff6ff, #e0e7ff)', 
      height: '100vh', 
      display: 'flex', 
      flexDirection: 'column',
      overflow: 'hidden'
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Bengali:wght@400;500;600;700&display=swap');
        
        * { box-sizing: border-box; margin: 0; padding: 0; }
        
        body { font-family: 'Noto Sans Bengali', sans-serif; }
        
        .loader {
          border: 3px solid #f3f4f6;
          border-top: 3px solid #3b82f6;
          border-radius: 50%;
          width: 24px;
          height: 24px;
          animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        .btn {
          padding: 12px 24px;
          border-radius: 8px;
          border: none;
          cursor: pointer;
          font-weight: 600;
          transition: all 0.2s;
        }
        
        .btn-primary {
          background: linear-gradient(to right, #4f46e5, #7c3aed);
          color: white;
        }
        
        .btn-primary:hover {
          background: linear-gradient(to right, #4338ca, #6d28d9);
        }
        
        .modal {
          position: fixed;
          inset: 0;
          background: rgba(0,0,0,0.5);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 50;
          padding: 16px;
        }
        
        .modal-content {
          background: white;
          border-radius: 16px;
          max-width: 600px;
          width: 100%;
          max-height: 90vh;
          overflow-y: auto;
        }
        
        .suggestion-card {
          border: 1px solid #e5e7eb;
          border-radius: 8px;
          padding: 16px;
          margin-bottom: 12px;
          transition: all 0.2s;
        }
        
        .suggestion-card:hover {
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        .suggestion-btn {
          width: 100%;
          text-align: left;
          padding: 8px 12px;
          border-radius: 6px;
          border: 1px solid;
          cursor: pointer;
          margin-top: 4px;
          transition: all 0.2s;
          font-weight: 500;
        }
      `}</style>

      {/* Settings Modal */}
      {showSettings && (
        <div className="modal" onClick={() => setShowSettings(false)}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div style={{ background: 'linear-gradient(to right, #4f46e5, #7c3aed)', color: 'white', padding: '24px', borderRadius: '16px 16px 0 0' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h2 style={{ fontSize: '24px', fontWeight: 'bold' }}>⚙️ সেটিংস</h2>
                <button onClick={() => setShowSettings(false)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', borderRadius: '50%', padding: '8px', cursor: 'pointer', color: 'white' }}>✕</button>
              </div>
            </div>
            <div style={{ padding: '24px' }}>
              <div style={{ marginBottom: '20px' }}>
                <label style={{ display: 'block', fontSize: '14px', fontWeight: '600', marginBottom: '8px' }}>🔑 Google Gemini API Key</label>
                <input
                  type="password"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="আপনার API Key এখানে দিন"
                  style={{ width: '100%', padding: '12px', border: '1px solid #d1d5db', borderRadius: '8px' }}
                />
              </div>
              <div style={{ marginBottom: '20px' }}>
                <label style={{ display: 'block', fontSize: '14px', fontWeight: '600', marginBottom: '8px' }}>🤖 AI Model সিলেক্ট করুন</label>
                <select
                  value={selectedModel}
                  onChange={(e) => setSelectedModel(e.target.value)}
                  style={{ width: '100%', padding: '12px', border: '1px solid #d1d5db', borderRadius: '8px' }}
                >
                  <option value="gemini-2.5-flash">Gemini 2.5 Flash (সর্বশেষ ও সেরা)</option>
                  <option value="gemini-2.0-flash-exp">Gemini 2.0 Flash (নতুন ও দ্রুততম)</option>
                  <option value="gemini-1.5-pro">Gemini 1.5 Pro (সেরা কোয়ালিটি)</option>
                  <option value="gemini-1.5-flash">Gemini 1.5 Flash (দ্রুত)</option>
                  <option value="gemini-pro">Gemini Pro (স্ট্যান্ডার্ড)</option>
                </select>
              </div>
              <div style={{ display: 'flex', gap: '12px' }}>
                <button onClick={saveSettings} className="btn btn-primary" style={{ flex: 1 }}>✓ সংরক্ষণ করুন</button>
                <button onClick={() => setShowSettings(false)} style={{ padding: '12px 24px', background: '#e5e7eb', borderRadius: '8px', border: 'none', cursor: 'pointer', fontWeight: '600' }}>বাতিল</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Instructions Modal */}
      {showInstructions && (
        <div className="modal" onClick={() => setShowInstructions(false)}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div style={{ background: 'linear-gradient(to right, #0d9488, #06b6d4)', color: 'white', padding: '24px', borderRadius: '16px 16px 0 0' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h2 style={{ fontSize: '24px', fontWeight: 'bold' }}>🎯 কীভাবে ব্যবহার করবেন</h2>
                <button onClick={() => setShowInstructions(false)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', borderRadius: '50%', padding: '8px', cursor: 'pointer', color: 'white' }}>✕</button>
              </div>
            </div>
            <div style={{ padding: '24px' }}>
              <ol style={{ paddingLeft: '20px', lineHeight: '1.8' }}>
                <li>⚙️ সেটিংস আইকনে ক্লিক করুন</li>
                <li>আপনার Google Gemini API Key দিন</li>
                <li>Word document এ বাংলা টেক্সট লিখুন</li>
                <li>"বানান পরীক্ষা করুন" বাটনে ক্লিক করুন</li>
                <li>সাজেশন দেখুন এবং ক্লিক করে প্রতিস্থাপন করুন</li>
              </ol>
            </div>
          </div>
        </div>
      )}

      {/* Main App Wrapper */}
      <div style={{ 
        background: 'white', 
        borderRadius: '0', 
        boxShadow: '0 4px 6px rgba(0,0,0,0.1)', 
        flex: 1, 
        display: 'flex', 
        flexDirection: 'column',
        height: '100%',
        overflow: 'hidden' // Important for inner layout
      }}>
        
        {/* Header - Fixed Height */}
        <div style={{ background: 'linear-gradient(to right, #4f46e5, #7c3aed)', color: 'white', padding: '20px', flexShrink: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <button onClick={() => setShowInstructions(true)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', borderRadius: '50%', padding: '8px', cursor: 'pointer', color: 'white' }}>❓</button>
            <div style={{ textAlign: 'center', flex: 1 }}>
              <h1 style={{ fontSize: '20px', fontWeight: 'bold', marginBottom: '4px' }}>🌟 ভাষা মিত্র</h1>
              <p style={{ fontSize: '12px', opacity: 0.9 }}>বাংলা বানান ও ব্যাকরণ শুদ্ধ করুন</p>
            </div>
            <button onClick={() => setShowSettings(true)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', borderRadius: '50%', padding: '8px', cursor: 'pointer', color: 'white' }}>⚙️</button>
          </div>
        </div>

        {/* Scrollable Content Area */}
        <div style={{ 
          padding: '20px', 
          flex: 1, 
          overflowY: 'auto', // This enables scrolling ONLY in the middle
          overflowX: 'hidden'
        }}>
          {/* Check Button */}
          <button 
            onClick={checkSpelling}
            disabled={isLoading}
            className="btn btn-primary"
            style={{ width: '100%', marginBottom: '16px', fontSize: '16px' }}
          >
            {isLoading ? 'পরীক্ষা করা হচ্ছে...' : 'বানান পরীক্ষা করুন'}
          </button>

          {/* Loading Indicator */}
          {isLoading && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', marginBottom: '16px', color: '#4f46e5' }}>
              <div className="loader"></div>
              <span>পরীক্ষা করা হচ্ছে...</span>
            </div>
          )}

          {/* Message */}
          {message && (
            <div style={{ 
              padding: '12px', 
              marginBottom: '16px', 
              borderRadius: '8px', 
              background: message.type === 'success' ? '#d1fae5' : '#fee2e2',
              color: message.type === 'success' ? '#065f46' : '#991b1b',
              border: `1px solid ${message.type === 'success' ? '#6ee7b7' : '#fca5a5'}`
            }}>
              {message.text}
            </div>
          )}

          {/* Statistics */}
          {stats.totalWords > 0 && (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '12px', marginBottom: '20px' }}>
              <div style={{ background: 'white', borderRadius: '8px', padding: '16px', textAlign: 'center', border: '1px solid #e5e7eb' }}>
                <p style={{ fontSize: '24px', fontWeight: 'bold', color: '#4f46e5' }}>{stats.totalWords}</p>
                <p style={{ fontSize: '12px', color: '#6b7280' }}>মোট শব্দ</p>
              </div>
              <div style={{ background: 'white', borderRadius: '8px', padding: '16px', textAlign: 'center', border: '1px solid #e5e7eb' }}>
                <p style={{ fontSize: '24px', fontWeight: 'bold', color: '#dc2626' }}>{stats.errorCount}</p>
                <p style={{ fontSize: '12px', color: '#6b7280' }}>ভুল বানান</p>
              </div>
              <div style={{ background: 'white', borderRadius: '8px', padding: '16px', textAlign: 'center', border: '1px solid #e5e7eb' }}>
                <p style={{ fontSize: '24px', fontWeight: 'bold', color: '#16a34a' }}>{stats.accuracy}%</p>
                <p style={{ fontSize: '12px', color: '#6b7280' }}>শুদ্ধতা</p>
              </div>
            </div>
          )}

          {/* Content Analysis */}
          {contentAnalysis && (
            <>
              <div style={{ border: '2px solid #6ee7b7', background: 'linear-gradient(to bottom right, #d1fae5, #a7f3d0)', borderRadius: '8px', padding: '16px', marginBottom: '16px' }}>
                <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#065f46', marginBottom: '12px' }}>📋 লেখার ধরন</h3>
                <p style={{ fontSize: '16px', fontWeight: 'bold', color: '#047857', marginBottom: '4px' }}>{contentAnalysis.contentType}</p>
                {contentAnalysis.description && <p style={{ fontSize: '12px', color: '#374151' }}>{contentAnalysis.description}</p>}
              </div>

              {contentAnalysis.missingElements && contentAnalysis.missingElements.length > 0 && (
                <div style={{ border: '2px solid #fcd34d', background: 'linear-gradient(to bottom right, #fef3c7, #fde68a)', borderRadius: '8px', padding: '16px', marginBottom: '16px' }}>
                  <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#78350f', marginBottom: '12px' }}>⚠️ যা যোগ করা উচিত</h3>
                  <ul>
                    {contentAnalysis.missingElements.map((element, i) => (
                      <li key={i} style={{ marginBottom: '8px', color: '#374151' }}>• {element}</li>
                    ))}
                  </ul>
                </div>
              )}

              {contentAnalysis.suggestions && contentAnalysis.suggestions.length > 0 && (
                <div style={{ border: '2px solid #5eead4', background: 'linear-gradient(to bottom right, #ccfbf1, #99f6e4)', borderRadius: '8px', padding: '16px', marginBottom: '16px' }}>
                  <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#115e59', marginBottom: '12px' }}>✨ উন্নতির পরামর্শ</h3>
                  <ul>
                    {contentAnalysis.suggestions.map((suggestion, i) => (
                      <li key={i} style={{ marginBottom: '8px', color: '#374151', display: 'flex', gap: '8px' }}>
                        <span style={{ flexShrink: 0, width: '20px', height: '20px', background: '#0d9488', color: 'white', borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '10px', fontWeight: 'bold' }}>{i + 1}</span>
                        <span>{suggestion}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </>
          )}

          {/* Spelling Errors */}
          {corrections.length > 0 && (
            <>
              <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#374151', marginBottom: '12px' }}>📝 বানান ভুল</h3>
              {corrections.map((correction, i) => (
                <div 
                  key={i} 
                  className="suggestion-card" 
                  style={{ borderColor: '#fecaca', background: 'rgba(254, 202, 202, 0.3)' }}
                  onMouseEnter={() => handleHoverWord(correction.wrong)}
                >
                  <div style={{ fontSize: '14px', fontWeight: '600', color: '#dc2626', marginBottom: '8px' }}>❌ {correction.wrong}</div>
                  <div style={{ fontSize: '12px', color: '#6b7280', marginBottom: '8px' }}>সঠিক বানান:</div>
                  {correction.suggestions.map((suggestion, j) => (
                    <button
                      key={j}
                      onClick={() => handleReplace(correction.wrong, suggestion)}
                      className="suggestion-btn"
                      style={{ background: '#d1fae5', borderColor: '#6ee7b7', color: '#065f46' }}
                    >
                      ✓ {suggestion}
                    </button>
                  ))}
                </div>
              ))}
            </>
          )}

          {/* Language Style Mixing */}
          {languageStyleMixing && languageStyleMixing.detected && (
            <>
              <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#374151', marginTop: '20px', marginBottom: '12px' }}>🔄 সাধু-চলিত ভাষা মিশ্রণ</h3>
              <div style={{ border: '1px solid #ddd6fe', background: 'rgba(216, 180, 254, 0.5)', borderRadius: '8px', padding: '16px', marginBottom: '12px' }}>
                <div style={{ fontSize: '14px', fontWeight: '600', color: '#6b21a8' }}>প্রস্তাবিত রীতি: {languageStyleMixing.recommendedStyle}</div>
                <div style={{ fontSize: '12px', color: '#6b7280', fontStyle: 'italic', marginTop: '4px' }}>{languageStyleMixing.reason}</div>
              </div>
              {languageStyleMixing.corrections?.map((correction, i) => (
                <div 
                  key={i} 
                  className="suggestion-card" 
                  style={{ borderColor: '#ddd6fe', background: 'rgba(216, 180, 254, 0.3)' }}
                  onMouseEnter={() => handleHoverWord(correction.current)}
                >
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginBottom: '8px' }}>
                    <span style={{ fontSize: '14px', fontWeight: '600', color: '#7c3aed' }}>🔄 {correction.current}</span>
                    <span style={{ fontSize: '10px', background: '#e9d5ff', color: '#6b21a8', padding: '2px 8px', borderRadius: '4px' }}>{correction.type}</span>
                  </div>
                  <button
                    onClick={() => handleReplace(correction.current, correction.suggestion)}
                    className="suggestion-btn"
                    style={{ background: '#e9d5ff', borderColor: '#c084fc', color: '#6b21a8' }}
                  >
                    ➜ {correction.suggestion}
                  </button>
                </div>
              ))}
            </>
          )}

          {/* Tone Improvements */}
          {toneImprovements.length > 0 && (
            <>
              <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#374151', marginTop: '20px', marginBottom: '12px' }}>💬 লেখার ভাব অনুযায়ী পরিবর্তন</h3>
              {toneImprovements.map((improvement, i) => (
                <div 
                  key={i} 
                  className="suggestion-card" 
                  style={{ borderColor: '#bfdbfe', background: 'rgba(191, 219, 254, 0.3)' }}
                  onMouseEnter={() => handleHoverWord(improvement.current)}
                >
                  <div style={{ fontSize: '14px', fontWeight: '600', color: '#2563eb', marginBottom: '4px' }}>💡 {improvement.current}</div>
                  <div style={{ fontSize: '11px', color: '#6b7280', fontStyle: 'italic', marginBottom: '8px' }}>{improvement.reason}</div>
                  {improvement.suggestions.map((suggestion, j) => (
                    <button
                      key={j}
                      onClick={() => handleReplace(improvement.current, suggestion)}
                      className="suggestion-btn"
                      style={{ background: '#dbeafe', borderColor: '#93c5fd', color: '#1e40af' }}
                    >
                      ✨ {suggestion}
                    </button>
                  ))}
                </div>
              ))}
            </>
          )}

          {/* Punctuation Issues */}
          {punctuationIssues.length > 0 && (
            <>
              <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#374151', marginTop: '20px', marginBottom: '12px' }}>🔤 বিরাম চিহ্ন সমস্যা</h3>
              {punctuationIssues.map((issue, i) => (
                <div key={i} className="suggestion-card" style={{ borderColor: '#fed7aa', background: 'rgba(254, 215, 170, 0.3)' }}>
                  <div style={{ fontSize: '14px', fontWeight: '600', color: '#ea580c', marginBottom: '4px' }}>⚠️ {issue.issue}</div>
                  <div style={{ fontSize: '11px', color: '#6b7280', fontStyle: 'italic', marginBottom: '8px' }}>{issue.explanation}</div>
                  <div style={{ background: '#fee2e2', border: '1px solid #fca5a5', borderRadius: '6px', padding: '8px', marginBottom: '8px' }}>
                    <div style={{ fontSize: '11px', color: '#6b7280', marginBottom: '4px' }}>বর্তমান:</div>
                    <div style={{ fontSize: '13px', color: '#374151' }}>{issue.currentSentence}</div>
                  </div>
                  <button
                    onClick={() => handleReplace(issue.currentSentence, issue.correctedSentence)}
                    className="suggestion-btn"
                    style={{ background: '#d1fae5', borderColor: '#6ee7b7', color: '#065f46' }}
                  >
                    <div style={{ fontSize: '11px', marginBottom: '4px' }}>✓ সংশোধিত:</div>
                    <div style={{ fontSize: '13px', fontWeight: '500' }}>{issue.correctedSentence}</div>
                  </button>
                </div>
              ))}
            </>
          )}

          {/* Euphony Improvements */}
          {euphonyImprovements.length > 0 && (
            <>
              <h3 style={{ fontSize: '14px', fontWeight: 'bold', color: '#374151', marginTop: '20px', marginBottom: '12px' }}>🎵 শ্রুতিমধুরতা উন্নতি</h3>
              {euphonyImprovements.map((improvement, i) => (
                <div 
                  key={i} 
                  className="suggestion-card" 
                  style={{ borderColor: '#fbcfe8', background: 'rgba(251, 207, 232, 0.3)' }}
                  onMouseEnter={() => handleHoverWord(improvement.current)}
                >
                  <div style={{ fontSize: '14px', fontWeight: '600', color: '#db2777', marginBottom: '4px' }}>🎵 {improvement.current}</div>
                  <div style={{ fontSize: '11px', color: '#6b7280', fontStyle: 'italic', marginBottom: '8px' }}>{improvement.reason}</div>
                  {improvement.suggestions.map((suggestion, j) => (
                    <button
                      key={j}
                      onClick={() => handleReplace(improvement.current, suggestion)}
                      className="suggestion-btn"
                      style={{ background: '#fce7f3', borderColor: '#f9a8d4', color: '#9f1239' }}
                    >
                      ♪ {suggestion}
                    </button>
                  ))}
                </div>
              ))}
            </>
          )}

          {/* No issues found message */}
          {!isLoading && corrections.length === 0 && toneImprovements.length === 0 && 
           !languageStyleMixing?.detected && punctuationIssues.length === 0 && 
           euphonyImprovements.length === 0 && !contentAnalysis && (
            <p style={{ textAlign: 'center', color: '#9ca3af', marginTop: '80px' }}>সাজেশন এখানে দেখা যাবে...</p>
          )}
        </div>

        {/* Developer Footer - FIXED BOTTOM */}
        <div style={{ 
          background: 'linear-gradient(to right, #f3f4f6, #e5e7eb)', 
          padding: '16px', 
          textAlign: 'center', 
          borderTop: '2px solid #d1d5db', 
          flexShrink: 0 
        }}>
          <p style={{ fontSize: '12px', color: '#6b7280', marginBottom: '4px', fontWeight: '600' }}>
            Developed by: হিমাদ্রি বিশ্বাস
          </p>
          <p style={{ fontSize: '11px', color: '#9ca3af' }}>
            📞 +880 9696 196566
          </p>
        </div>

      </div>
    </div>
  );
}

// Initialize Office and React
Office.onReady(() => {
  const root = ReactDOM.createRoot(document.getElementById('root')!);
  root.render(<App />);
});
