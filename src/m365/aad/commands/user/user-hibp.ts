import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
  apiKey: string;
  domain?: string
}

class AadUserHibpCommand extends AnonymousCommand {
  public get name(): string {
    return commands.USER_HIBP;
  }

  public get description(): string {
    return 'Allows you to retrieve all accounts that have been pwned with the specified username';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.domain = args.options.domain;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `https://haveibeenpwned.com/api/v3/breachedaccount/${encodeURIComponent(args.options.userName)}${(args.options.domain ? `?domain=${encodeURIComponent(args.options.domain)}` : '')}`,
      headers: {
        'accept': 'application/json',
        'hibp-api-key': args.options.apiKey,
        'x-anonymous': true
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      })
      .catch((err: any): void => {
        if ((err && err.response !== undefined && err.response.status === 404) && (this.debug || this.verbose)) {
          logger.log('No pwnage found');
          cb();
          return;
        }
        return this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --userName <userName>'
      },
      {
        option: '--apiKey, <apiKey>'
      },
      {
        option: '--domain, [domain]'
      }
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidUserPrincipalName(args.options.userName)) {
      return 'Specify valid userName';
    }

    return true;
  }
}

module.exports = new AadUserHibpCommand();
