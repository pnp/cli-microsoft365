import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  loginName?: string;
  confirm: boolean;
}

class SpoUserRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_REMOVE;
  }

  public get description(): string {
    return 'Removes user from specific web';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.loginName = (!(!args.options.loginName)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeUser = (): void => {
      if (this.verbose) {
        logger.log(`Removing user from  subsite ${args.options.webUrl} ...`);
      }

      let requestUrl: string = '';

      if (args.options.id) {
        requestUrl = `${encodeURI(args.options.webUrl)}/_api/web/siteusers/removebyid(${args.options.id})`;
      }

      if (args.options.loginName) {
        requestUrl = `${encodeURI(args.options.webUrl)}/_api/web/siteusers/removeByLoginName('${encodeURIComponent(args.options.loginName as string)}')`;
      }

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .post(requestOptions)
        .then((): void => {
          if (this.verbose) {
            logger.log(chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }

    if (args.options.confirm) {
      removeUser();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove specified user from the site ${args.options.webUrl}`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeUser();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the web to remove user'
      },
      {
        option: '-i, --id [id]',
        description: 'ID of the user to remove from web'
      },
      {
        option: '--loginName [loginName]',
        description: 'Login name of the site user to remove'
      },
      {
        option: '--confirm',
        description: 'Do not prompt for confirmation before removing user from web'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.id && !args.options.loginName) {
      return 'Required option id or loginName missing, one is required';
    }

    if (args.options.id && args.options.loginName) {
      return 'Use either id or loginName, but not both';
    }

    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoUserRemoveCommand();