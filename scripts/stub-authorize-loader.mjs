import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const stubUrl = pathToFileURL(
  join(dirname(fileURLToPath(import.meta.url)), 'stub-authorize.js')
).href;

export async function resolve(specifier, context, nextResolve) {
  const parent = context.parentURL || '';
  if (specifier === './auth.js' && parent.endsWith('/dist/server.js')) {
    return { shortCircuit: true, url: stubUrl };
  }
  return nextResolve(specifier, context);
}
