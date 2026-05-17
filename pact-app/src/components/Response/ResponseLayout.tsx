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
            src="/kcc-logo.png"
            alt="KCC"
          />
          <h1>PACT</h1>
        </div>
        <span className="response-header-badge">Konstructum Compliance Governance</span>
      </header>

      {/* Content */}
      <main className="response-content">
        <Outlet />
      </main>

      {/* Footer */}
      <footer className="response-footer">
        <p>This is a secure, automated notification system for Konstructum Group.</p>
        <p>PACT Compliance Governance Platform © {new Date().getFullYear()}. All rights reserved.</p>
      </footer>
    </div>
  );
};
