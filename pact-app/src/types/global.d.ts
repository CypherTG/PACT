export {};

declare global {
  interface Window {
    PACT_QA?: any;
    pact_mock_email?: any;
  }
}
