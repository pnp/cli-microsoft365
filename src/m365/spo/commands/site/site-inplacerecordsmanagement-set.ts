import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  enabled: string;
}

class SpoSiteInPlaceRecordsManagementSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_INPLACERECORDSMANAGEMENT_SET;
  }

  public get description(): string {
    return 'Activates or deactivates in-place records management for a site collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.enabled = args.options.enabled;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const enabled: boolean = args.options.enabled.toLocaleLowerCase() === 'true';

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/site/features/${enabled ? 'add' : 'remove'}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        featureId: 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9',
        force: true
      },
      responseType: 'json'
    };

    if (this.verbose) {
      logger.logToStderr(`${enabled ? 'Activating' : 'Deactivating'} in-place records management for site ${args.options.siteUrl}`);
    }

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--enabled <enabled>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidBoolean(args.options.enabled)) {
      return 'Invalid "enabled" option value. Specify "true" or "false"';
    }

    return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
  }
}

module.exports = new SpoSiteInPlaceRecordsManagementSetCommand();