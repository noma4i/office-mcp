import { execFile } from 'child_process';
import { promisify } from 'util';
import { getErrorMessage } from '../validators.js';

const execFileAsync = promisify(execFile);

export async function runAppleScript(script) {
  try {
    const { stdout } = await execFileAsync('osascript', ['-e', script], { timeout: 30000 });
    return stdout.trim();
  } catch (error) {
    throw new Error(`AppleScript error: ${getErrorMessage(error)}`);
  }
}
