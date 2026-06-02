import React, { useEffect } from 'react';
import { Outlet } from 'react-router-dom';
import './ResponseLayout.css';

const RESPONSE_PAGE_CLASS = 'pact-response-active';

/**
 * ResponseLayout — Minimal, employee-facing layout (no sidebar navigation)
 * Used for Accept/Appeal response pages that employees access via email links.
 */
export const ResponseLayout: React.FC = () => {
  useEffect(() => {
    document.documentElement.classList.add(RESPONSE_PAGE_CLASS);
    return () => document.documentElement.classList.remove(RESPONSE_PAGE_CLASS);
  }, []);

  return (
    <div className="response-layout" role="main">
      <div className="response-shell">
        <header className="response-header">
          <div className="response-logo">
            <img src="/kcc-logo.png" alt="KCC" />
            <h1>PACT</h1>
          </div>
          <span className="response-header-badge">Konstructum Compliance Governance</span>
        </header>

        <div className="response-content">
          <Outlet />
        </div>

        <footer className="response-footer">
          <p>This is a secure, automated notification system for Konstructum Group.</p>
          <p>PACT Compliance Governance Platform © {new Date().getFullYear()}. All rights reserved.</p>
        </footer>
      </div>
    </div>
  );
};
