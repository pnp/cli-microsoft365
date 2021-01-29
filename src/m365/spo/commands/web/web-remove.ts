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
  confirm?: boolean;
}

class SpoWebAddCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_REMOVE;
  }

  public get description(): string {
    return 'Delete specified subsite';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeWeb = (): void => {
      const requestOptions: any = {
        url: `${encodeURI(args.options.webUrl)}/_api/web`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'X-HTTP-Method': 'DELETE'
        },
        responseType: 'json'
      };

      if (this.verbose) {
        logger.logToStderr(`Deleting subsite ${args.options.webUrl} ...`);
      }

      request
        .post(requestOptions)
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }

    if (args.options.confirm) {
      removeWeb();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the subsite ${args.options.webUrl}`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeWeb();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoWebAddCommand();