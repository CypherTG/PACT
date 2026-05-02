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

export default class PactApp extends React.Component<IPactAppProps> {
  constructor(props: IPactAppProps) {
    super(props);
    // Initialize the singleton service with the SPFx context
    SharePointService.init(props.context);
  }

  public async componentDidMount(): Promise<void> {
    if (sharePointService) {
      await sharePointService.initialize();
    }
  }

  public render(): React.ReactElement<IPactAppProps> {
    const { userDisplayName } = this.props;

    return (
      <AppContext.Provider value={{
        service: sharePointService,
        userDisplayName
      }}>
        <ErrorBoundary>
          <App />
        </ErrorBoundary>
      </AppContext.Provider>
    );
  }
}
