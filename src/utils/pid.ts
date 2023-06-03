import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import { cache } from './cache';

function getProcessNameOnMacOs(pid: number): string | undefined {
  const res = spawnSync('ps', ['-o', 'comm=', pid.toString()], { encoding: 'utf8' });
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
  const findProcess = spawnSync('wmic', ['PROCESS', 'where', `ProcessId=${pid}`, 'get', 'Caption'], { encoding: 'utf8' });
  if (findProcess.error || findProcess.stderr) {
    return undefined;
  }
  else {
    const getPid = spawnSync('find', ['/V', '"Caption"'], {
      encoding: 'utf8',
      input: findProcess.stdout,
      // must include or passing quoted "Caption" will fail
      windowsVerbatimArguments: true
    });
    if (getPid.error || getPid.stderr) {
      return undefined;
    }
    else {
      return getPid.stdout.trim();
    }
  }
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