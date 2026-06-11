import * as React from 'react';
import { App } from '../pact/App';
import { SharePointService, sharePointService } from '../pact/services/SharePointService';
import { IPactAppProps } from './IPactAppProps';
import '../pact/styles';

export interface IAppContext {
  service: SharePointService;
  userDisplayName: string;
}

export const AppContext = React.createContext<IAppContext | null>(null);

interface ErrorBoundaryState {
  hasError: boolean;
  message: string;
  stack?: string;
}

class ErrorBoundary extends React.Component<React.PropsWithChildren<{}>, ErrorBoundaryState> {
  public constructor(props: React.PropsWithChildren<{}>) {
    super(props);
    this.state = { hasError: false, message: '', stack: undefined };
  }

  public static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return {
      hasError: true,
      message: error.message,
      stack: error.stack,
    };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('PACT app render error', error, info.componentStack);
  }

  public render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div className="glass-panel" style={{ margin: '2rem', padding: '1.5rem', borderLeft: '4px solid #ef4444' }}>
          <h2 style={{ marginTop: 0, color: '#ef4444' }}>PACT failed to render</h2>
          <div style={{ marginBottom: '0.75rem', fontWeight: 600 }}>Runtime error</div>
          <pre style={{ whiteSpace: 'pre-wrap', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>
            {this.state.message}
            {this.state.stack ? `\n\n${this.state.stack}` : ''}
          </pre>
        </div>
      );
    }

    return this.props.children;
  }
}

interface PactAppState {
  isOpen: boolean;
}

export default class PactApp extends React.Component<IPactAppProps, PactAppState> {
  constructor(props: IPactAppProps) {
    super(props);
    this.state = {
      isOpen: false
    };
    // Initialize the singleton service with the SPFx context
    SharePointService.init(props.context);
  }

  public async componentDidMount(): Promise<void> {
    if (sharePointService) {
      await sharePointService.initialize();
    }
  }

  public componentDidUpdate(prevProps: IPactAppProps, prevState: PactAppState): void {
    if (this.state.isOpen !== prevState.isOpen) {
      if (typeof document !== 'undefined') {
        if (this.state.isOpen) {
          document.body.style.overflow = 'hidden';
        } else {
          document.body.style.overflow = '';
        }
      }
    }
  }

  public componentWillUnmount(): void {
    if (typeof document !== 'undefined') {
      document.body.style.overflow = '';
    }
  }

  private isCaseResponseRoute(): boolean {
    if (typeof window === 'undefined') return false;
    const hash = window.location.hash || '';
    const path = window.location.pathname || '';
    return hash.includes('/case-response/') || path.includes('/case-response/');
  }


  public render(): React.ReactElement<IPactAppProps> {
    const { userDisplayName } = this.props;
    const isResponse = this.isCaseResponseRoute();

    // 1. Employee Accept/Appeal link: Render as a clean, standalone fullscreen page (covering all SharePoint chrome)
    if (isResponse) {
      return (
        <AppContext.Provider value={{ service: sharePointService, userDisplayName }}>
          <ErrorBoundary>
            <div style={{
              position: 'fixed',
              top: 0,
              left: 0,
              width: '100vw',
              height: '100vh',
              zIndex: 999999,
              background: '#f4f7fb',
              overflow: 'auto'
            }}>
              <App />
            </div>
          </ErrorBoundary>
        </AppContext.Provider>
      );
    }

    // Check if opened via the "open" hash
    const isFullApp = typeof window !== 'undefined' && (
      window.location.hash === '#/open' ||
      (window.location.hash.startsWith('#/') && window.location.hash !== '#/')
    );

    // 2. Full PACT Portal view: Render as fullscreen overlay with a Close/Exit header
    if (isFullApp || this.state.isOpen) {
      return (
        <AppContext.Provider value={{ service: sharePointService, userDisplayName }}>
          <ErrorBoundary>
            <div style={{
              position: 'fixed',
              top: 0,
              left: 0,
              width: '100vw',
              height: '100vh',
              zIndex: 999999,
              background: '#0B0F19',
              display: 'flex',
              flexDirection: 'column',
              overflow: 'hidden'
            }}>
              <div style={{
                background: '#111827',
                borderBottom: '1px solid #1f2937',
                padding: '12px 24px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                color: '#ffffff',
                fontFamily: 'Segoe UI, sans-serif',
                boxSizing: 'border-box',
                height: '56px',
                flexShrink: 0
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <img src={require('../pact/assets/kcc-logo.png')} alt="KCC" style={{ height: '28px', width: 'auto' }} />
                  <span style={{ fontWeight: 600, fontSize: '15px', letterSpacing: '0.5px' }}>PACT Portal Preview</span>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    this.setState({ isOpen: false });
                    if (typeof window !== 'undefined') {
                      window.location.hash = '';
                    }
                  }}
                  style={{
                    background: '#C0272D',
                    color: '#ffffff',
                    border: 'none',
                    borderRadius: '4px',
                    padding: '8px 16px',
                    fontSize: '13px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    transition: 'background 0.2s',
                    boxShadow: 'none'
                  }}
                  onMouseOver={(e) => (e.currentTarget.style.background = '#A02025')}
                  onMouseOut={(e) => (e.currentTarget.style.background = '#C0272D')}
                >
                  Close Preview
                </button>
              </div>
              <div style={{ flex: 1, overflow: 'auto' }}>
                <App />
              </div>
            </div>
          </ErrorBoundary>
        </AppContext.Provider>
      );
    }

    return (
      <AppContext.Provider value={{
        service: sharePointService,
        userDisplayName
      }}>
        <ErrorBoundary>
          {/* Custom Styles for Startup Card */}
          <style>{`
            .pact-startup-container {
              display: flex;
              justify-content: center;
              align-items: center;
              padding: 3rem 1.5rem;
              font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
              background: transparent;
            }
            .pact-startup-card {
              background: #FFFFFF;
              border: 1px solid #E0E0E0 !important;
              border-radius: 8px;
              padding: 3rem 3rem;
              max-width: 580px;
              width: 100%;
              box-shadow: none !important;
              text-align: center;
              box-sizing: border-box;
              transition: all 0.3s ease;
            }
            .pact-startup-card:hover {
              box-shadow: none !important;
              transform: none !important;
            }
            .pact-startup-logo {
              width: 64px;
              height: 64px;
              margin: 0 auto 1.5rem;
              background: transparent;
              border-radius: 0;
              display: flex;
              align-items: center;
              justify-content: center;
              font-size: 48px;
              font-family: serif;
              font-weight: bold;
              color: #1A1A1A;
              box-shadow: none;
              overflow: hidden;
            }
            .pact-startup-title {
              font-size: 20px;
              font-weight: 700;
              color: #1A1A1A;
              margin: 0 0 0.5rem;
            }
            .pact-startup-subtitle {
              font-size: 11px;
              color: #C0272D;
              font-weight: 600;
              margin: 0 0 1.25rem;
              text-transform: uppercase;
              letter-spacing: 2px;
            }
            .pact-startup-desc {
              font-size: 13px;
              color: #666666;
              line-height: 1.6;
              margin-bottom: 2rem;
            }
            .pact-startup-btn {
              display: inline-flex;
              align-items: center;
              justify-content: center;
              width: 100%;
              padding: 0.85rem 1.5rem;
              background: #C0272D;
              color: white;
              font-size: 14px;
              font-weight: 600;
              border: none;
              border-radius: 6px;
              cursor: pointer;
              box-shadow: none;
              transition: all 0.2s ease;
              gap: 8px;
            }
            .pact-startup-btn:hover {
              transform: translateY(-1px);
              background: #A02025;
              box-shadow: none;
            }
          `}</style>

          <div className="pact-startup-container">
            <div className="pact-startup-card">
              <div className="pact-startup-logo" style={{ overflow: 'hidden', padding: '0', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'transparent' }}>
                <img src={require('../pact/assets/kcc-logo.png')} alt="KCC" style={{ width: '100%', height: '100%', objectFit: 'contain' }} />
              </div>
              <h2 className="pact-startup-title">PACT Portal</h2>
              <div className="pact-startup-subtitle">Konstructum</div>
              <p className="pact-startup-desc">
                Policy Agreement &amp; Compliance Tracking Platform. Securely manage compliance cases, disciplinary metrics, and mail logs.
              </p>
              <button 
                type="button" 
                className="pact-startup-btn" 
                onClick={() => {
                  this.setState({ isOpen: true });
                }}
              >
                Open PACT Portal
              </button>
            </div>
          </div>
        </ErrorBoundary>
      </AppContext.Provider>
    );
  }
}

