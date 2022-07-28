import * as util from 'util';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (args.options.confirm) {
      this.clearTeamsCache(logger, cb);
    }
    else {
      logger.logToStderr('This command will execute the following steps.');
      logger.logToStderr('- Stop the Microsoft Teams client.');
      logger.logToStderr('- Clear the Microsoft Teams cached files.');

      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear your Microsoft Teams cache?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          this.clearTeamsCache(logger, cb);
        }
      });
    }
  }

  private async clearTeamsCache(logger: Logger, cb: (err?: any) => void): Promise<void> {
    try {
      await this.killRunningProcess(logger);
      await this.removeCacheFiles(logger);
    }
    catch (e: any) {
      cb(e.message as string);
      return;
    }

    logger.logToStderr('Teams cache cleared!');
    cb();
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): boolean | string {
    if (process.env.CLIMICROSOFT365_ENV === 'docker') {
      return 'Because you\'re running CLI for Microsoft 365 in a Docker container, we can\'t clear the cache on your host. Instead run this command on your host using "npx ..."';
    }

    if (process.platform !== 'win32' && process.platform !== 'darwin') {
      return `${process.platform} platform is unsupported for this command`;
    }

    return true;
  }
}

module.exports = new TeamsCacheRemoveCommand();