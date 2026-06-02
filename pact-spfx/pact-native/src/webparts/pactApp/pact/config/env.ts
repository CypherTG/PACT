/**
 * Safe read of Vite env vars. SPFx webpack may not define import.meta.env.
 */
export function readViteEnv(key: string): string {
  try {
    const env = (globalThis as any).__PACT_ENV__;
    if (env && typeof env[key] === 'string') {
      return env[key];
    }
  } catch {
    /* SPFx / non-Vite bundles */
  }

  return '';
}
