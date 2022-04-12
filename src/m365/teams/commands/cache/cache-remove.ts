import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import util = require('util');
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (process.env.CLIMICROSOFT365_ENV === 'docker') {
      logger.log('Because you\'re running CLI for Microsoft 365 in a Docker container, we can\'t clear the cache on your host. Instead run this command on your host using "npx ..."');
      cb();
      return;
    }
    
    if (process.platform !== 'win32' && process.platform !== 'darwin') {
      logger.log(`${process.platform} platform is unsupported for this command`);
      cb();
      return;
    }

    if (args.options.confirm) {
      this.clearTeamsCache(logger, cb);
    }
    else {
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

  private async clearTeamsCache(logger: Logger, cb: () => void): Promise<void> {
    await this.killRunningProcess();
    await this.removeCachingFiles();

    logger.log('Teams cache cleared!');
    cb();
  }

  private async killRunningProcess(): Promise<void> {
    const platform = process.platform;
    let cmd = '';

    switch (platform) {
      case 'win32': cmd = `taskkill /IM "Teams.exe" /F`; break;
      case 'darwin': cmd = `kill -9 \`pidof Teams.exe\``; break;
    }

    await this.exec(cmd);
  }

  private async removeCachingFiles(): Promise<void> {
    const platform = process.platform;
    let cmd = '';

    switch (platform) {
      case 'win32': cmd = `cd %userprofile% && rmdir /s /q AppData\\Roaming\\Microsoft\\Teams`; break;
      case 'darwin': cmd = `rm -r ~/Library/Application\ Support/Microsoft/Teams`; break;
    }

    await this.exec(cmd);
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
}

module.exports = new TeamsCacheRemoveCommand();