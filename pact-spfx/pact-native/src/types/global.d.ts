export {};

declare global {
  interface Window {
    PACT_QA?: {
      refresh: () => Promise<void>;
      getSnapshot: () => Promise<Record<string, unknown>>;
      setPanelVisible: (visible: boolean) => void;
      getMode: () => string;
    };
  }
}
