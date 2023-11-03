import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { profileCardPropertyNames, ProfileCardProperty } from './profileCardProperties.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class TenantPeopleProfileCardPropertyGetCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_GET;
  }

  public get description(): string {
    return 'Retrieves information about a specific profile card property';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>',
        autocomplete: profileCardPropertyNames
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!profileCardPropertyNames.some(p => p.toLowerCase() === args.options.name.toLowerCase())) {
          return `'${args.options.name}' is not a valid value for option name. Allowed values are: ${profileCardPropertyNames.join(', ')}.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving information about profile card property '${args.options.name}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties/${args.options.name}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const result = await request.get<ProfileCardProperty>(requestOptions);
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
      if (err.response?.status === 404) {
        this.handleError(`Profile card property '${args.options.name}' does not exist.`);
      }

      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantPeopleProfileCardPropertyGetCommand();