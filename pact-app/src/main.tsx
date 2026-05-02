import React from 'react'
import ReactDOM from 'react-dom/client'
import { App } from './App.tsx'
import './index.css'

const rootElement = document.getElementById('root');

const renderCrashGuard = (error: any) => {
  if (!rootElement) return;
  const message = error instanceof Error ? error.message : String(error);
  const stack = error instanceof Error ? error.stack : 'No stack trace available';

  rootElement.innerHTML = `
    <div style="padding: 40px; font-family: sans-serif; color: #721c24; background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 12px; margin: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.1)">
      <h2 style="margin-top: 0; color: #a94442;">🚨 PACT Diagnostic: Application Error</h2>
      <p>The system encountered a runtime crash. Please copy the error details below:</p>
      <hr style="opacity: 0.1; margin: 20px 0" />
      <div style="background: rgba(0,0,0,0.05); padding: 15px; border-radius: 4px; overflow-x: auto; font-size: 13px; line-height: 1.5;">
        <strong>Error:</strong> ${message}<br/><br/>
        <strong>Details:</strong><br/>
        <pre style="margin: 0; white-space: pre-wrap;">${stack}</pre>
      </div>
      <div style="margin-top: 24px; display: flex; gap: 12px;">
        <button onclick="localStorage.clear(); location.reload();" style="padding: 12px 20px; background: #721c24; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 600;">
          Wipe Cache & Restart
        </button>
        <button onclick="location.reload();" style="padding: 12px 20px; background: #f5c6cb; color: #721c24; border: 1px solid #721c24; border-radius: 6px; cursor: pointer; font-weight: 600;">
          Simple Refresh
        </button>
      </div>
    </div>
  `;
};

// Global Listeners for "White Screen" detection
window.onerror = (message, _source, _lineno, _colno, error) => {
  renderCrashGuard(error || message);
  return false;
};

window.onunhandledrejection = (event) => {
  renderCrashGuard(event.reason);
};

try {
  if (!rootElement) throw new Error("Root element not found");

  ReactDOM.createRoot(rootElement).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  );
} catch (error) {
  renderCrashGuard(error);
}
