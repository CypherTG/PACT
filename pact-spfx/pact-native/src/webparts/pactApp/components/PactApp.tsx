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

    // If it's an employee case response route, render it immediately
    if (isResponse) {
      return (
        <AppContext.Provider value={{ service: sharePointService, userDisplayName }}>
          <ErrorBoundary>
            <App />
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
          {/* Custom Styles for Startup Card & Overlay */}
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
              background: linear-gradient(135deg, #fffefe, #fff5f5);
              border: none !important;
              border-radius: 16px;
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
              background: linear-gradient(135deg, #dc2626, #ff6b35);
              border-radius: 16px;
              display: flex;
              align-items: center;
              justify-content: center;
              font-size: 2rem;
              font-weight: bold;
              color: white;
              box-shadow: 0 8px 20px rgba(220, 38, 38, 0.25);
              overflow: hidden;
            }
            .pact-startup-title {
              font-size: 1.6rem;
              font-weight: 700;
              color: #111827;
              margin: 0 0 0.5rem;
            }
            .pact-startup-subtitle {
              font-size: 0.95rem;
              color: #4b5563;
              font-weight: 600;
              margin: 0 0 1.25rem;
              text-transform: uppercase;
              letter-spacing: 0.05em;
            }
            .pact-startup-desc {
              font-size: 0.875rem;
              color: #6b7280;
              line-height: 1.6;
              margin-bottom: 2rem;
            }
            .pact-startup-btn {
              display: inline-flex;
              align-items: center;
              justify-content: center;
              width: 100%;
              padding: 0.85rem 1.5rem;
              background: linear-gradient(135deg, #dc2626, #d2334e);
              color: white;
              font-size: 1rem;
              font-weight: 600;
              border: none;
              border-radius: 10px;
              cursor: pointer;
              box-shadow: 0 4px 14px rgba(220, 38, 38, 0.3);
              transition: all 0.2s ease;
              gap: 8px;
            }
            .pact-startup-btn:hover {
              transform: translateY(-1px);
              box-shadow: 0 6px 20px rgba(220, 38, 38, 0.5);
              background: linear-gradient(135deg, #b91c1c, #dc2626);
            }
            
            /* Full page overlay layout */
            .pact-fullscreen-wrapper {
              position: fixed !important;
              top: 0 !important;
              left: 0 !important;
              width: 100vw !important;
              height: 100vh !important;
              z-index: 999999 !important;
              background: #f9fafb !important;
              overflow: auto !important;
              display: flex;
              flex-direction: column;
            }
            .pact-fullscreen-banner {
              display: flex;
              justify-content: space-between;
              align-items: center;
              padding: 10px 24px;
              background: #111827;
              color: #f3f4f6;
              font-size: 0.85rem;
              font-weight: 500;
              border-bottom: 1px solid rgba(255,255,255,0.05);
            }
            .pact-fullscreen-btn-close {
              background: rgba(255,255,255,0.1);
              color: white;
              border: 1px solid rgba(255,255,255,0.15);
              padding: 6px 16px;
              border-radius: 6px;
              font-size: 0.78rem;
              cursor: pointer;
              font-weight: 600;
              transition: all 0.2s ease;
            }
            .pact-fullscreen-btn-close:hover {
              background: rgba(255,255,255,0.2);
              transform: translateY(-1px);
            }
            .pact-fullscreen-content {
              flex: 1;
              position: relative;
            }
          `}</style>

          {this.state.isOpen ? (
            <div className="pact-fullscreen-wrapper">
              <div className="pact-fullscreen-banner">
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#10b981', boxShadow: '0 0 8px #10b981' }}></span>
                  <span>PACT Client Portal is active and fully functional on all screen sizes.</span>
                </div>
                <button 
                  type="button" 
                  className="pact-fullscreen-btn-close" 
                  onClick={() => this.setState({ isOpen: false })}
                >
                  Close & Minimize Portal
                </button>
              </div>
              <div className="pact-fullscreen-content">
                <App />
              </div>
            </div>
          ) : (
            <div className="pact-startup-container">
              <div className="pact-startup-card">
                <div className="pact-startup-logo">
                  <img 
                    src={require('../assets/kcc-logo.png')} 
                    alt="KCC" 
                    style={{ width: '100%', height: '100%', objectFit: 'contain' }}
                    onError={(e) => {
                      (e.target as HTMLImageElement).style.display = 'none';
                      const parent = (e.target as HTMLImageElement).parentElement;
                      if (parent) {
                        parent.innerText = 'P';
                      }
                    }}
                  />
                </div>
                <h2 className="pact-startup-title">PACT Portal</h2>
                <div className="pact-startup-subtitle">Konstructum</div>
                <p className="pact-startup-desc">
                  Policy Agreement & Compliance Tracking Platform. Securely manage compliance cases, disciplinary metrics, and mail logs.
                </p>
                <button 
                  type="button" 
                  className="pact-startup-btn" 
                  onClick={() => this.setState({ isOpen: true })}
                >
                  Open PACT Portal
                </button>
              </div>
            </div>
          )}
        </ErrorBoundary>
      </AppContext.Provider>
    );
  }
}
