import React from 'react';
import './ResponseLayout.css';

interface ResponseLayoutProps {
  children?: React.ReactNode;
}

/**
 * ResponseLayout — Minimal, employee-facing layout (no sidebar navigation)
 * Used for Accept/Appeal response pages that employees access via email links.
 */
export const ResponseLayout: React.FC<ResponseLayoutProps> = ({ children }) => {
  return (
    <div className="response-layout">
      {/* Minimal Header */}
      <header className="response-header">
        <div className="response-logo">
          <div className="logo-icon" style={{ width: '36px', height: '36px', overflow: 'hidden', background: 'transparent', boxShadow: 'none' }}>
            <img 
              src={require('../../assets/kcc-logo.png')} 
              alt="KCC" 
              style={{ width: '100%', height: '100%', objectFit: 'contain' }}
              onError={(e) => {
                (e.target as HTMLImageElement).style.display = 'none';
                const parent = (e.target as HTMLImageElement).parentElement;
                if (parent) {
                  parent.innerText = 'P';
                  parent.style.background = 'linear-gradient(135deg, var(--primary), #ff6b35)';
                  parent.style.display = 'flex';
                  parent.style.alignItems = 'center';
                  parent.style.justifyContent = 'center';
                  parent.style.color = 'white';
                  parent.style.fontWeight = 'bold';
                  parent.style.borderRadius = '8px';
                }
              }}
            />
          </div>
          <h1>PACT</h1>
        </div>
        <span className="response-header-badge">Compliance Governance Platform</span>
      </header>

      {/* Content */}
      <div className="response-content">
        {children}
      </div>

      {/* Footer */}
      <footer className="response-footer">
        PACT Compliance Governance Platform © {new Date().getFullYear()} Konstructum. All rights reserved.
      </footer>
    </div>
  );
};
