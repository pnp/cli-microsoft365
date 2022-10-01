import * as os from 'os';
import * as fs from 'fs';
import { execSync } from 'child_process';
import { cache } from './cache';

function getProcessNameOnMacOs(pid: number): string | undefined {
  const stdout = execSync(`ps -o comm= ${pid}`, { encoding: 'utf8' });
  return stdout.trim();
}

function getProcessNameOnLinux(pid: number): string | undefined {
  if (!fs.existsSync(`/proc/${pid}/stat`)) {
    return undefined;
  }

  const stat: string = fs.readFileSync(`/proc/${pid}/stat`, 'utf8');
  const start: number = stat.indexOf('(');
  const procName = stat.substring(start + 1, stat.indexOf(')') - start);
  return procName;
}

function getProcessNameOnWindows(pid: number): string | undefined {
  const stdout = execSync(`wmic PROCESS where ProcessId=${pid} get Caption | find /V "Caption"`, { encoding: 'utf8' });
  return stdout.trim();
}

export const pid = {
  getProcessName(pid: number): string | undefined {
    let processName: string | undefined = cache.getValue(pid.toString());
    if (processName) {
      return processName;
    }

    let getPidName: ((pid: number) => string | undefined) | undefined = undefined;

    const platform: string = os.platform();
    if (platform.indexOf('win') === 0) {
      getPidName = getProcessNameOnWindows;
    }
    if (platform === 'darwin') {
      getPidName = getProcessNameOnMacOs;
    }
    if (platform === 'linux') {
      getPidName = getProcessNameOnLinux;
    }

    if (getPidName) {
      try {
        processName = getPidName(pid);
        if (processName) {
          cache.setValue(pid.toString(), processName);
        }
        return processName;
      }
      catch {
        return undefined;
      }
    }

    return undefined;
  }
};