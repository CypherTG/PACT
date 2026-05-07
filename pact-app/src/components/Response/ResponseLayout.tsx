import React from 'react';
import { Outlet } from 'react-router-dom';
import './ResponseLayout.css';

/**
 * ResponseLayout — Minimal, employee-facing layout (no sidebar navigation)
 * Used for Accept/Appeal response pages that employees access via email links.
 */
export const ResponseLayout: React.FC = () => {
  return (
    <div className="response-layout">
      {/* Minimal Header */}
      <header className="response-header">
        <div className="response-logo">
          <img
            src="./kcc-logo.png"
            alt="KCC"
            onError={(e) => {
              (e.target as HTMLImageElement).style.display = 'none';
              const fallback = document.getElementById('response-logo-fallback');
              if (fallback) fallback.style.display = 'flex';
            }}
          />
          <div id="response-logo-fallback" className="response-logo-fallback" style={{ display: 'none' }}>
            P
          </div>
          <h1>PACT</h1>
        </div>
        <span className="response-header-badge">Compliance Governance Platform</span>
      </header>

      {/* Content */}
      <div className="response-content">
        <Outlet />
      </div>

      {/* Footer */}
      <footer className="response-footer">
        PACT Compliance Governance Platform © {new Date().getFullYear()} Konstructum. All rights reserved.
      </footer>
    </div>
  );
};
