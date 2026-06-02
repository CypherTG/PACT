/**
 * Safe read of Vite env vars. SPFx webpack may not define import.meta.env.
 */
export function readViteEnv(key: string): string {
  try {
    const env = (import.meta as any).env;
    if (env && typeof env[key] === 'string') {
      return env[key];
    }
  } catch {
    /* non-Vite bundles */
  }

  return '';
}
