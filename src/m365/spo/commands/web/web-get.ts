import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { WebProperties } from './WebProperties.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  url: string;
  withGroups?: boolean;
  withPermissions?: boolean;
}

class SpoWebGetCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_GET;
  }

  public get description(): string {
    return 'Retrieve information about the specified site';
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
        withGroups: !!args.options.withGroups,
        withPermissions: !!args.options.withPermissions
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '--withGroups'
      },
      {
        option: '--withPermissions'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let url: string = `${args.options.url}/_api/web`;
    if (args.options.withGroups) {
      url += '?$expand=AssociatedMemberGroup,AssociatedOwnerGroup,AssociatedVisitorGroup';
    }
    const requestOptions: any = {
      url,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const webProperties: WebProperties = await request.get<WebProperties>(requestOptions);
      if (args.options.withPermissions) {
        requestOptions.url = `${args.options.url}/_api/web/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
        const response = await request.get<{ value: any[] }>(requestOptions);
        response.value.forEach((r: any) => {
          r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
        });
        webProperties.RoleAssignments = response.value;
      }
      await logger.log(webProperties);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebGetCommand();