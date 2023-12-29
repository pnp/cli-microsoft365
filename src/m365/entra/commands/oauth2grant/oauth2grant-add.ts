import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  clientId: string;
  resourceId: string;
  scope: string;
}

class AadOAuth2GrantAddCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_ADD;
  }

  public get description(): string {
    return 'Grant the specified service principal OAuth2 permissions to the specified resource';
  }

  public alias(): string[] | undefined {
    return [aadCommands.OAUTH2GRANT_ADD];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --clientId <clientId>'
      },
      {
        option: '-r, --resourceId <resourceId>'
      },
      {
        option: '-s, --scope <scope>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.clientId)) {
          return `${args.options.clientId} is not a valid GUID`;
        }

        if (!validation.isValidGuid(args.options.resourceId)) {
          return `${args.options.resourceId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Granting the service principal specified permissions...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/oauth2PermissionGrants`,
        headers: {
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          "clientId": args.options.clientId,
          "consentType": "AllPrincipals",
          "principalId": null,
          "resourceId": args.options.resourceId,
          "scope": args.options.scope
        }
      };

      await request.post<void>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadOAuth2GrantAddCommand();