import child_process from 'child_process';
import fs from 'fs';
import { homedir } from 'os';
import util from 'util';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  client?: string;
  force?: boolean;
}

interface Win32Process {
  PID: number;
}

class TeamsCacheRemoveCommand extends AnonymousCommand {
  private static readonly allowedClients: string[] = ['new', 'classic'];

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
        client: args.options.client,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-c, --client',
        autocomplete: TeamsCacheRemoveCommand.allowedClients
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.client && !TeamsCacheRemoveCommand.allowedClients.includes(args.options.client.toLowerCase())) {
          return `'${args.options.client}' is not a valid value for option 'client'. Allowed values are ${TeamsCacheRemoveCommand.allowedClients.join(', ')}`;
        }

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
        await this.clearTeamsCache(args.options.client?.toLowerCase() || 'new', logger);
      }
      else {
        await logger.logToStderr('This command will execute the following steps.');
        await logger.logToStderr('- Stop the Microsoft Teams client.');
        await logger.logToStderr('- Clear the Microsoft Teams cached files.');

        const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear your Microsoft Teams cache?` });

        if (result) {
          await this.clearTeamsCache(args.options.client?.toLowerCase() || 'new', logger);
        }
      }
    }
    catch (err: any) {
      this.handleError(err);
    }
  }

  private async clearTeamsCache(client: string, logger: Logger): Promise<void> {
    const filePaths = await this.getTeamsCacheFolderPaths(client, logger);

    let folderExists = true;
    for (const filePath of filePaths) {
      const exists = await this.checkIfCacheFolderExists(filePath, logger);
      if (!exists) {
        folderExists = false;
      }
    }

    if (folderExists) {
      await this.killRunningProcess(client, logger);
      await this.removeCacheFiles(filePaths, logger);
      await logger.logToStderr('Teams cache cleared!');
    }
    else {
      await logger.logToStderr('Cache folder does not exist. Nothing to remove.');
    }

  }

  private async getTeamsCacheFolderPaths(client: string, logger: Logger): Promise<string[]> {
    const platform = process.platform;

    if (this.verbose) {
      await logger.logToStderr(`Getting path of Teams cache folder for platform ${platform}...`);
    }
    const filePaths: string[] = [];

    switch (platform) {
      case 'win32':
        if (client === 'classic') {
          filePaths.push(`${process.env.APPDATA}\\Microsoft\\Teams`);
        }
        else {
          filePaths.push(`${process.env.LOCALAPPDATA}\\Packages\\MSTeams_8wekyb3d8bbwe\\LocalCache\\Microsoft\\MSTeams`);
        }
        break;
      case 'darwin':
        if (client === 'classic') {
          filePaths.push(`${homedir}/Library/Application Support/Microsoft/Teams`);
        }
        else {
          filePaths.push(`${homedir}/Library/Group Containers/UBF8T346G9.com.microsoft.teams`);
          filePaths.push(`${homedir}/Library/Containers/com.microsoft.teams2`);
        }
        break;
    }
    return filePaths;
  }

  private async checkIfCacheFolderExists(filePath: string, logger: Logger): Promise<boolean> {
    if (this.verbose) {
      await logger.logToStderr(`Checking if Teams cache folder exists at ${filePath}...`);
    }

    return fs.existsSync(filePath);
  }

  private async killRunningProcess(client: string, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Stopping Teams client...');
    }

    const platform = process.platform;
    let cmd = '';

    switch (platform) {
      case 'win32':
        if (client === 'classic') {
          cmd = 'tasklist /FI "IMAGENAME eq Teams.exe" /FO csv';
        }
        else {
          cmd = 'tasklist /FI "IMAGENAME eq ms-teams.exe" /FO csv';
        }
        break;
      case 'darwin':
        if (client === 'classic') {
          cmd = `ps ax | grep MacOS/Teams -m 1 | grep -v grep | awk '{ print $1 }'`;
        }
        else {
          cmd = `ps ax | grep MacOS/MSTeams -m 1 | grep -v grep | awk '{ print $1 }'`;
        }

        break;
    }

    if (this.debug) {
      await logger.logToStderr(cmd);
    }

    const cmdOutput = await this.exec(cmd);

    if (platform === 'darwin' && cmdOutput.stdout) {
      process.kill(parseInt(cmdOutput.stdout));
    }
    else if (platform === 'win32') {
      const processJson: Win32Process[] = formatting.parseCsvToJson(cmdOutput.stdout);
      for (const proc of processJson) {
        process.kill(proc.PID);
      }
    }
    if (this.verbose) {
      await logger.logToStderr('Teams client closed');
    }
  }

  private async removeCacheFiles(filePaths: string[], logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Removing Teams cache files...');
    }

    const platform = process.platform;
    const baseCmd = platform === 'win32' ? 'rmdir /s /q ' : 'rm -r ';

    for (const filePath of filePaths) {
      const cmd = `${baseCmd}"${filePath}"`;

      if (this.debug) {
        await logger.logToStderr(cmd);
      }

      try {
        await this.exec(cmd);
      }
      catch (err: any) {
        if (err?.stderr?.includes('Operation not permitted')) {
          await logger.log('Deleting the folder failed. Please have a look at the following URL to delete the folders manually: https://answers.microsoft.com/en-us/msteams/forum/all/clearing-cache-on-microsoft-teams/35876f6b-eb1a-4b77-bed1-02ce3277091f');
        }
        else {
          throw err;
        }
      }
    }
  }

  private exec = util.promisify(child_process.exec);
}

export default new TeamsCacheRemoveCommand();