import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { profileCardPropertyNames } from './profileCardProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  displayName?: string;
}

class TenantPeopleProfileCardPropertyAddCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_ADD;
  }

  public get description(): string {
    return 'Adds an additional attribute to the profile card properties';
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
        name: args.options.name,
        displayName: typeof args.options.displayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>',
        autocomplete: profileCardPropertyNames
      },
      {
        option: '-d, --displayName [displayName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const propertyName = args.options.name.toLowerCase();

        if (profileCardPropertyNames.every(n => n.toLowerCase() !== propertyName)) {
          return `${args.options.name} is not a valid value for name. Allowed values are ${profileCardPropertyNames.join(', ')}`;
        }

        if (propertyName.startsWith("customattribute") && args.options.displayName === undefined) {
          return `The option 'displayName' is required when adding customAttributes as profile card properties`;
        }

        if (!propertyName.startsWith("customattribute") && args.options.displayName !== undefined) {
          return `The option 'displayName' can only be used when adding customAttributes as profile card properties`;
        }

        const excludeOptions: string[] = ['name', 'displayName', 'debug', 'verbose', 'output'];
        const unknownOptions = Object.keys(args.options).filter(key => !excludeOptions.includes(key));

        if (!propertyName.startsWith('customattribute') && unknownOptions.length > 0) {
          return `Unknown options like ${unknownOptions.join(', ')} are only supported with customAttributes`;
        }

        if (propertyName.startsWith('customattribute')) {
          const wronglyFormattedOptions = unknownOptions.filter(key => !key.toLowerCase().startsWith('displayname-'));
          if (wronglyFormattedOptions.length > 0) {
            return `Wrong option format detected for the following option(s): ${wronglyFormattedOptions.join(', ')}'. When adding localizations for customAttributes, use the format displayName-<languageTag>.`;
          }
        }

        return true;
      }
    );
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding '${args.options.name}' as a profile card property...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/admin/people/profileCardProperties`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        directoryPropertyName: args.options.name,
        annotations: this.getAnnotations(args.options)
      }
    };

    try {
      const response: any = await request.post(requestOptions);

      if (args.options.output !== 'json') {
        const annotation = response.annotations[0];

        if (annotation) {
          response.displayName = annotation.displayName;
          annotation.localizations.forEach((l: { languageTag: string, displayName: string }) => {
            response[`displayName-${l.languageTag}`] = l.displayName;
          });
        }

        delete response.annotations;
      }

      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getAnnotations(options: Options): { displayName: string, localizations?: { languageTag: string, displayName: string }[] }[] {
    if (!options.displayName) {
      return [];
    }

    return [
      {
        displayName: options.displayName!,
        localizations: this.getLocalizations(options)
      }
    ];
  }

  private getLocalizations(options: Options): { languageTag: string, displayName: string }[] {
    const excludeOptions: string[] = ['name', 'displayName', 'debug', 'verbose', 'output'];
    const unknownOptions = Object.keys(options).filter(key => !excludeOptions.includes(key));

    if (unknownOptions.length === 0) {
      return [];
    }

    const localizations: { languageTag: string, displayName: string }[] = [];

    unknownOptions.forEach(key => {
      localizations.push({
        languageTag: key.replace('displayName-', ''),
        displayName: options[key]
      });
    });

    return localizations;
  }
}

export default new TenantPeopleProfileCardPropertyAddCommand();