import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Localization, ProfileCardProperty, profileCardPropertyNames as allProfileCardPropertyNames } from './profileCardProperties.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  displayName?: string;
}

class TenantPeopleProfileCardPropertySetCommand extends GraphCommand {
  private readonly profileCardPropertyNames = allProfileCardPropertyNames.filter(p => p.toLowerCase().startsWith('customattribute'));

  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_SET;
  }

  public get description(): string {
    return 'Updates a custom attribute to the profile card property';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      // Add unknown options to telemetry
      const unknownOptions = Object.keys(this.getUnknownOptions(args.options));
      const unknownOptionsObj = unknownOptions.reduce((obj, key) => ({ ...obj, [key]: true }), {});

      Object.assign(this.telemetryProperties, {
        displayName: typeof args.options.displayName !== 'undefined',
        ...unknownOptionsObj
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>',
        autocomplete: this.profileCardPropertyNames
      },
      {
        option: '-d, --displayName <displayName>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!this.profileCardPropertyNames.some(p => p.toLowerCase() === args.options.name.toLowerCase())) {
          return `'${args.options.name}' is not a valid value for option name. Allowed values are: ${this.profileCardPropertyNames.join(', ')}.`;
        }

        // Unknown options are allowed only if they start with 'displayName-'
        const unknownOptionKeys = Object.keys(this.getUnknownOptions(args.options));
        const invalidOptionKey = unknownOptionKeys.find(o => !o.startsWith('displayName-'));
        if (invalidOptionKey) {
          return `Invalid option: '${invalidOptionKey}'`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('name', 'displayName');
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Updating profile card property '${args.options.name}'...`);
      }

      // Get the right casing for the profile card property name
      const profileCardProperty = this.profileCardPropertyNames.find(p => p.toLowerCase() === args.options.name.toLowerCase());

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties/${profileCardProperty}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          annotations: [
            {
              displayName: args.options.displayName,
              localizations: this.getLocalizations(args.options)
            }
          ]
        }
      };

      const result = await request.patch<ProfileCardProperty>(requestOptions);
      let output: any = result;

      // Transform the output to make it more readable
      if (args.options.output && args.options.output !== 'json' && result.annotations.length > 0) {
        output = result.annotations[0].localizations.reduce((acc, curr) => ({
          ...acc,
          ['displayName ' + curr.languageTag]: curr.displayName
        }), {
          ...result,
          displayName: result.annotations[0].displayName
        });

        delete output.annotations;
      }

      await logger.log(output);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Transform option to localization object.
   * @example Transform "--displayName-en-US 'Cost center'" to { languageTag: 'en-US', displayName: 'Cost center' }
   */
  private getLocalizations(options: Options): Localization[] {
    const unknownOptions = this.getUnknownOptions(options);

    const result = Object.keys(unknownOptions).map(o => ({
      languageTag: o.substring(o.indexOf('-') + 1),
      displayName: unknownOptions[o]
    } as Localization));

    return result;
  }
}

export default new TenantPeopleProfileCardPropertySetCommand();