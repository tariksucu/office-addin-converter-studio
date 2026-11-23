// Types of tools/manifest.mjs, for the tests.
export interface AddinConfig {
  readonly baseUrl: string;
  readonly providerName: string;
  readonly supportUrl: string;
  readonly outlook: { readonly id: string; readonly name: string; readonly description: string };
  readonly office: { readonly id: string; readonly name: string; readonly description: string };
}
export function esc(value: unknown): string;
export function manifestVersion(version: string): string;
export function normalBase(base: string): string;
export function outlookManifest(config: AddinConfig, version: string, baseUrl?: string): string;
export function officeManifest(config: AddinConfig, version: string, baseUrl?: string): string;
