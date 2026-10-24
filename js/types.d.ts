// Declarations for the JSDoc type check (npm run typecheck); not loaded by the application.

// UMD wrappers of the modules: AMD, CommonJS (tests) or browser globals
declare const define: ((deps: string[], factory: (...args: any[]) => any) => void) & { amd?: unknown };
declare const module: { exports: any } | undefined;
declare function require(id: string): any;
declare const Buffer: any;

// Errors carry a machine-readable code (translated by the UI) and message parameters
interface Error {
  code?: string;
  params?: Record<string, unknown>;
}
