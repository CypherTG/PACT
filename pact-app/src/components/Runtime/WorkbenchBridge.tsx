import React from 'react';
import { Copy, RefreshCw, ShieldCheck, Monitor, PanelRightOpen, PanelRightClose } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';

type Snapshot = {
  mode: string;
  runtimeLabel: string;
  route: string;
  lastRefreshed: string;
  totals: {
    cases: number;
    staff: number;
    appeals: number;
    mails: number;
    escalations: number;
    activeCases: number;
    repeatOffenders: number;
  };
};

const params = new URLSearchParams(window.location.search);
const qaEnabled = params.get('qa') === '1' || params.get('pactQa') === '1';

const isEmbedded = (() => {
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
})();

export const WorkbenchBridge: React.FC = () => {
  const [visible, setVisible] = React.useState(qaEnabled || isEmbedded);
  const [snapshot, setSnapshot] = React.useState<Snapshot>({
    mode: sharePointService.isStandalone() ? 'demo' : 'sharepoint',
    runtimeLabel: sharePointService.getRuntimeLabel(),
    route: window.location.hash || window.location.pathname || '/',
    lastRefreshed: new Date().toISOString(),
    totals: {
      cases: 0,
      staff: 0,
      appeals: 0,
      mails: 0,
      escalations: 0,
      activeCases: 0,
      repeatOffenders: 0
    }
  });
  const snapshotRef = React.useRef(snapshot);

  const refreshSnapshot = React.useCallback(async () => {
    const [cases, staff, appeals, mails, escalations, stats] = await Promise.all([
      sharePointService.getCases(),
      sharePointService.getStaffDirectory(),
      sharePointService.getAppeals(),
      sharePointService.getMailHistory(),
      sharePointService.getEscalationLog(),
      sharePointService.getDashboardStats()
    ]);

    const nextSnapshot: Snapshot = {
      mode: sharePointService.isStandalone() ? 'demo' : 'sharepoint',
      runtimeLabel: sharePointService.getRuntimeLabel(),
      route: window.location.hash || window.location.pathname || '/',
      lastRefreshed: new Date().toISOString(),
      totals: {
        cases: cases.length,
        staff: staff.length,
        appeals: appeals.length,
        mails: mails.length,
        escalations: escalations.length,
        activeCases: stats.totalActiveCases,
        repeatOffenders: stats.repeatOffenders
      }
    };

    snapshotRef.current = nextSnapshot;
    setSnapshot(nextSnapshot);

    window.parent?.postMessage(
      {
        type: 'PACT_RUNTIME_UPDATE',
        payload: nextSnapshot
      },
      '*'
    );
  }, []);

  React.useEffect(() => {
    let timer: number | undefined;
    const start = async () => {
      await sharePointService.initialize();
      await refreshSnapshot();
      timer = window.setInterval(() => {
        void refreshSnapshot();
      }, 15000);
    };
    void start();

    const handleMessage = (event: MessageEvent) => {
      const message = event.data;
      if (!message || typeof message !== 'object') return;
      if (message.type === 'PACT_QA_REFRESH') {
        void refreshSnapshot();
      }
      if (message.type === 'PACT_QA_PANEL') {
        setVisible(Boolean(message.visible));
      }
    };

    const handleKeydown = (event: KeyboardEvent) => {
      if (event.ctrlKey && event.altKey && event.key.toLowerCase() === 'q') {
        setVisible(value => !value);
      }
    };

    window.addEventListener('message', handleMessage);
    window.addEventListener('keydown', handleKeydown);
    return () => {
      if (timer) window.clearInterval(timer);
      window.removeEventListener('message', handleMessage);
      window.removeEventListener('keydown', handleKeydown);
    };
  }, [refreshSnapshot]);

  React.useEffect(() => {
    document.title = `PACT | ${snapshot.runtimeLabel} | ${snapshot.route}`;
  }, [snapshot.runtimeLabel, snapshot.route]);

  React.useEffect(() => {
    window.PACT_QA = {
      refresh: refreshSnapshot,
      getSnapshot: async () => snapshotRef.current,
      setPanelVisible: setVisible,
      getMode: () => snapshotRef.current.mode
    };

    return () => {
      delete window.PACT_QA;
    };
  }, [refreshSnapshot]);

  const copySnapshot = async () => {
    try {
      await navigator.clipboard.writeText(JSON.stringify(snapshotRef.current, null, 2));
    } catch {
      console.warn('Clipboard access was blocked.');
    }
  };

  if (!visible) {
    return (
      <button
        type="button"
        className="qa-fab glass-panel"
        onClick={() => setVisible(true)}
        data-testid="qa-fab"
      >
        <PanelRightOpen size={16} />
        QA
      </button>
    );
  }

  return (
    <aside className="qa-panel glass-panel" data-testid="qa-panel">
      <div className="qa-panel-header">
        <div>
          <div className="qa-kicker">
            <Monitor size={14} />
            SharePoint QA
          </div>
          <h3>Workbench telemetry</h3>
        </div>
        <button type="button" className="qa-icon-button" onClick={() => setVisible(false)} aria-label="Close QA panel">
          <PanelRightClose size={16} />
        </button>
      </div>

      <div className="qa-grid">
        <div className="qa-stat">
          <span>Mode</span>
          <strong>{snapshot.runtimeLabel}</strong>
        </div>
        <div className="qa-stat">
          <span>Route</span>
          <strong>{snapshot.route}</strong>
        </div>
        <div className="qa-stat">
          <span>Cases</span>
          <strong>{snapshot.totals.cases}</strong>
        </div>
        <div className="qa-stat">
          <span>Active</span>
          <strong>{snapshot.totals.activeCases}</strong>
        </div>
        <div className="qa-stat">
          <span>Staff</span>
          <strong>{snapshot.totals.staff}</strong>
        </div>
        <div className="qa-stat">
          <span>Escalations</span>
          <strong>{snapshot.totals.escalations}</strong>
        </div>
      </div>

      <div className="qa-actions">
        <button type="button" className="btn btn-secondary" onClick={() => void refreshSnapshot()}>
          <RefreshCw size={14} />
          Refresh
        </button>
        <button type="button" className="btn btn-secondary" onClick={() => void copySnapshot()}>
          <Copy size={14} />
          Copy JSON
        </button>
      </div>

      <div className="qa-footer">
        <ShieldCheck size={14} />
        Last refreshed {new Date(snapshot.lastRefreshed).toLocaleTimeString()}
      </div>
    </aside>
  );
};
