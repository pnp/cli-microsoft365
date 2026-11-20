import child_process from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { cache } from './cache.js';

function getProcessNameOnMacOs(pid: number): string | undefined {
  const res = child_process.spawnSync('ps', ['-o', 'comm=', pid.toString()], { encoding: 'utf8' });
  if (res.error || res.stderr) {
    return undefined;
  }
  else {
    return res.stdout.trim();
  }
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
  const tasklist = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'tasklist.exe');
  const res = child_process.spawnSync(tasklist, ['/FI', `PID eq ${pid}`, '/FO', 'CSV', '/NH'], { encoding: 'utf8' });
  if (res.error || res.status !== 0) {
    return undefined;
  }

  // "pwsh.exe","12345","Console","1","98,765 K"
  const match = res.stdout.match(/^"([^"]+)","(\d+)"/m);
  return match && match[2] === pid.toString() ? match[1] : undefined;
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
      processName = getPidName(pid);
      if (processName) {
        cache.setValue(pid.toString(), processName);
      }
      return processName;
    }

    return undefined;
  },
  isPowerShell(): boolean {
    const processName: string | undefined = pid.getProcessName(process.ppid) || '';
    return processName.indexOf('powershell') > -1 ||
      processName.indexOf('pwsh') > -1;
  }
};