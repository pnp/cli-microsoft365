import child_process from 'child_process';
import fs from 'fs';
import { homedir } from 'os';
import util from 'util';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force?: boolean;
}

interface Win32Process {
  ProcessId: number;
}

class TeamsCacheRemoveCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CACHE_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Teams client cache';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async () => {
        if (process.env.CLIMICROSOFT365_ENV === 'docker') {
          return 'Because you\'re running CLI for Microsoft 365 in a Docker container, we can\'t clear the cache on your host. Instead run this command on your host using "npx ..."';
        }

        if (process.platform !== 'win32' && process.platform !== 'darwin') {
          return `${process.platform} platform is unsupported for this command`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        await this.clearTeamsCache(logger);
      }
      else {
        await logger.logToStderr('This command will execute the following steps.');
        await logger.logToStderr('- Stop the Microsoft Teams client.');
        await logger.logToStderr('- Clear the Microsoft Teams cached files.');

        const result = await Cli.promptForConfirmation(`Are you sure you want to clear your Microsoft Teams cache?`);

        if (result) {
          await this.clearTeamsCache(logger);
        }
      }
    }
    catch (err: any) {
      this.handleError(err);
    }
  }

  private async clearTeamsCache(logger: Logger): Promise<void> {
    const filePath = await this.getTeamsCacheFolderPath(logger);
    const folderExists = await this.checkIfCacheFolderExists(filePath, logger);

    if (folderExists) {
      await this.killRunningProcess(logger);
      await this.removeCacheFiles(filePath, logger);
      await logger.logToStderr('Teams cache cleared!');
    }
    else {
      await logger.logToStderr('Cache folder does not exist. Nothing to remove.');
    }

  }

  private async getTeamsCacheFolderPath(logger: Logger): Promise<string> {
    const platform = process.platform;

    if (this.verbose) {
      await logger.logToStderr(`Getting path of Teams cache folder for platform ${platform}...`);
    }

    let filePath = '';

    switch (platform) {
      case 'win32':
        filePath = `${process.env.APPDATA}\\Microsoft\\Teams`;
        break;
      case 'darwin':
        filePath = `${homedir}/Library/Application Support/Microsoft/Teams`;
        break;
    }
    return filePath;
  }

  private async checkIfCacheFolderExists(filePath: string, logger: Logger): Promise<boolean> {
    if (this.verbose) {
      await logger.logToStderr(`Checking if Teams cache folder exists at ${filePath}...`);
    }

    return fs.existsSync(filePath);
  }

  private async killRunningProcess(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Stopping Teams client...');
    }

    const platform = process.platform;
    let cmd = '';

    switch (platform) {
      case 'win32':
        cmd = 'wmic process where caption="Teams.exe" get ProcessId';
        break;
      case 'darwin':
        cmd = `ps ax | grep MacOS/Teams -m 1 | grep -v grep | awk '{ print $1 }'`;
        break;
    }

    if (this.debug) {
      await logger.logToStderr(cmd);
    }

    const cmdOutput = await this.exec(cmd);

    if (platform === 'darwin') {
      process.kill(parseInt(cmdOutput.stdout));
    }
    else if (platform === 'win32') {
      const processJson: Win32Process[] = formatting.parseCsvToJson(cmdOutput.stdout);
      processJson.filter(proc => proc.ProcessId).map((proc: Win32Process) => {
        process.kill(proc.ProcessId);
      });
    }
    if (this.verbose) {
      await logger.logToStderr('Teams client closed');
    }
  }

  private async removeCacheFiles(filePath: string, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Removing Teams cache files...');
    }

    const platform = process.platform;
    let cmd = '';

    switch (platform) {
      case 'win32':
        cmd = `rmdir /s /q "${filePath}"`;
        break;
      case 'darwin':
        cmd = `rm -r "${filePath}"`;
        break;
    }

    if (this.debug) {
      await logger.logToStderr(cmd);
    }

    await this.exec(cmd);
  }

  private exec = util.promisify(child_process.exec);
}

export default new TeamsCacheRemoveCommand();