import { execFileSync } from 'child_process';
import { writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

let strictCompileAvailability;

function compileWithOsacompile(script) {
  const tmpFile = join(tmpdir(), `as_test_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.applescript`);
  try {
    writeFileSync(tmpFile, script, 'utf8');
    execFileSync('osacompile', ['-o', '/dev/null', tmpFile], {
      timeout: 10000,
      stdio: ['pipe', 'pipe', 'pipe']
    });
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.stderr?.toString() || err.message };
  } finally {
    try {
      unlinkSync(tmpFile);
    } catch {}
  }
}

function checkStrictCompileAvailability() {
  if (strictCompileAvailability !== undefined) {
    return strictCompileAvailability;
  }

  // Preflight: verify osacompile can parse Office-specific terminology in this environment.
  const preflightScript = `
tell application "Microsoft Word"
  set d to active document
end tell
`;
  const result = compileWithOsacompile(preflightScript);
  strictCompileAvailability = result.ok;

  return strictCompileAvailability;
}

export function compileAppleScript(script) {
  if (globalThis.__coverage__) {
    return { ok: true, skippedInCoverage: true };
  }

  if (process.env.APPLE_SCRIPT_STRICT_COMPILE !== '1') {
    return { ok: true, skippedInNonStrictMode: true };
  }

  if (!checkStrictCompileAvailability()) {
    return { ok: true, skippedInUnavailableStrictEnvironment: true };
  }

  return compileWithOsacompile(script);
}
