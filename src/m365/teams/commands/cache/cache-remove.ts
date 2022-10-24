import * as util from 'util';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import AnonymousCommand from '../../../base/AnonymousCommand';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends AnonymousCommand {
  confirm?: boolean;
}

class TeamsCacheRemoveCommand extends GraphCommand {
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
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--confirm'
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
      if (args.options.confirm) {
        await this.clearTeamsCache(logger);
      }
      else {
        logger.logToStderr('This command will execute the following steps.');
        logger.logToStderr('- Stop the Microsoft Teams client.');
        logger.logToStderr('- Clear the Microsoft Teams cached files.');

        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to clear your Microsoft Teams cache?`
        });

        if (result.continue) {
          await this.clearTeamsCache(logger);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async clearTeamsCache(logger: Logger): Promise<void> {
    try {
      const folderExists = await this.checkIfCacheFolderExists(logger);
      if (folderExists) {
        await this.killRunningProcess(logger);
        await this.removeCacheFiles(logger);
        logger.logToStderr('Teams cache cleared!');
      }
      else {
        logger.logToStderr('Cache folder does not exist. Nothing to remove.');
      }
    }
    catch (e: any) {
      throw e.message as string;
    }
  }

  private async checkIfCacheFolderExists(logger: Logger): Promise<boolean> {
    if (this.verbose) {
      logger.logToStderr('Checking if cache folder exists');
    }

    const platform = process.platform;
    let cmd = '';
    const echoMessage = 'echo Folder does not exist';

    switch (platform) {
      case 'win32':
        cmd = `IF NOT EXIST %userprofile%\\appdata\\roaming\\microsoft\\teams ${echoMessage}`;
        break;
      case 'darwin':
        cmd = `if [ ! -d  ~/Library/Application\\ Support/Microsoft/Teams ]
        then
        ${echoMessage}
        fi`;
        break;
    }

    if (this.debug) {
      logger.logToStderr(cmd);
    }

    try {
      const cmdOutput = await this.exec(cmd);

      if (cmdOutput.stdout !== '' && cmdOutput.stdout.startsWith('Folder does not exist')) {
        if (this.verbose) {
          logger.logToStderr(`Teams cache folder exists for ${platform}. Continuing the deletion`);
        }
        return false;
      }
      else {
        return true;
      }
    }
    catch (e: any) {
      throw new Error(e.message);
    }
  }

  private async killRunningProcess(logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Stop Teams client');
    }

    const platform = process.platform;
    let cmd = '';

    switch (platform) {
      case 'win32':
        cmd = 'taskkill /IM "Teams.exe" /F';
        break;
      case 'darwin':
        cmd = `ps ax | grep MacOS/Teams -m 1 | grep -v grep | awk '{ print $1 }'`;
        break;
    }

    if (this.debug) {
      logger.logToStderr(cmd);
    }

    try {
      const cmdOutput = await this.exec(cmd);

      if (cmdOutput.stdout !== '' && platform === 'darwin') {
        process.kill(cmdOutput.stdout);
      }

      if (this.verbose) {
        logger.logToStderr('Teams client closed');
      }
    }
    catch (e: any) {
      const errorMessage = e.message as string;

      if (errorMessage.includes('ERROR: The process "Teams.exe" not found.')) {
        if (this.verbose) {
          logger.logToStderr('Teams client isn\'t running');
        }
      }
      else {
        throw new Error(errorMessage);
      }
    }
  }

  private async removeCacheFiles(logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Clear Teams cached files');
    }

    const platform = process.platform;
    let cmd = '';

    switch (platform) {
      case 'win32':
        cmd = 'cd %userprofile% && rmdir /s /q AppData\\Roaming\\Microsoft\\Teams';
        break;
      case 'darwin':
        cmd = 'rm -r ~/Library/Application\\ Support/Microsoft/Teams';
        break;
    }

    if (this.debug) {
      logger.logToStderr(cmd);
    }

    try {
      await this.exec(cmd);
    }
    catch (e: any) {
      throw new Error(e.message as string);
    }
  }

  private exec = util.promisify(require('child_process').exec);
}

module.exports = new TeamsCacheRemoveCommand();