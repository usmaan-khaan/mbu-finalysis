import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';
import './tw-built.css';

// This static build has no backend, so it can never safely hold an Anthropic
// API key (anything shipped in client-side JS is visible to every visitor).
// Setting this flag makes the two LLM-backed features (the "Generate AI
// insights" narrative summary and the "Ask a question" chat on the Insights
// tab) fail with a clear, friendly message instead of firing a doomed,
// keyless request at api.anthropic.com from the browser. Every other
// feature — including the fully offline rule-based pattern detection,
// correlations, outliers, and threshold alerts on that same tab — is
// unaffected. See the README's "AI features" section if you want to wire
// these up for real via your own backend proxy.
window.__ARTIFACT_BUILD__ = true;

// Saved setups / saved views persist through this async get/set/delete
// shim on top of the browser's own localStorage — each visitor's saves
// stay local to their own browser, nothing is sent anywhere.
const LS_PREFIX = 'mbu_finalysis:';
window.storage = {
  async get(key) {
    try {
      const raw = localStorage.getItem(LS_PREFIX + key);
      if (raw === null) return null;
      return { value: raw };
    } catch (e) {
      return null;
    }
  },
  async set(key, value) {
    try {
      localStorage.setItem(LS_PREFIX + key, value);
    } catch (e) {
      throw new Error('Storage not available in this environment.');
    }
    return true;
  },
  async delete(key) {
    try {
      localStorage.removeItem(LS_PREFIX + key);
    } catch (e) {
      /* nothing to delete */
    }
    return true;
  },
};

createRoot(document.getElementById('root')).render(<App />);
