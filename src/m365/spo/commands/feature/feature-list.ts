import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { Feature } from './Feature';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  scope?: string;
}

class SpoFeatureListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.FEATURE_LIST}`;
  }

  public get description(): string {
    return 'Lists Features activated in the specified site or site collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'Web';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope : 'Web';
    const requestOptions: any = {
      url: `${args.options.url}/_api/${scope}/Features?$select=DisplayName,DefinitionId`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<{ value: Feature[] }>(requestOptions)
      .then((features: { value: Feature[] }): void => {
        if (features.value && features.value.length > 0) {
          cmd.log(features.value);
        }
        else {
          if (this.verbose) {
            cmd.log('No activated Features found');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site (collection) to retrieve the activated Features from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the Features to retrieve. Allowed values Site|Web. Default Web',
        autocomplete: ['Site', 'Web']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.scope) {
        if (args.options.scope !== 'Site' &&
          args.options.scope !== 'Web') {
          return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
        }
      }

      return SpoCommand.isValidSharePointUrl(args.options.url);
    };
  }
}

module.exports = new SpoFeatureListCommand();