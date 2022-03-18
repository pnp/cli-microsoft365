import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Feature } from './Feature';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  scope?: string;
}

class SpoFeatureListCommand extends SpoCommand {
  public get name(): string {
    return commands.FEATURE_LIST;
  }

  public get description(): string {
    return 'Lists Features activated in the specified site or site collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'Web';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope : 'Web';
    const requestOptions: any = {
      url: `${args.options.url}/_api/${scope}/Features?$select=DisplayName,DefinitionId`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: Feature[] }>(requestOptions)
      .then((features: { value: Feature[] }): void => {
        if (features.value && features.value.length > 0) {
          logger.log(features.value);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No activated Features found');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.scope) {
      if (args.options.scope !== 'Site' &&
        args.options.scope !== 'Web') {
        return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
      }
    }

    return validation.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoFeatureListCommand();