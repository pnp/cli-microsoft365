import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import { BasePermissions } from '../../base-permissions';
import commands from '../../commands';
import { RoleAssignment, RoleDefinition } from '../roledefinition/RoleDefinition';
import { RoleType } from '../roledefinition/RoleType';
import { WebProperties } from './WebProperties';

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
        const roleAssignments = this.setFriendlyPermissions(response.value);
        webProperties.RoleAssignments = roleAssignments;
      }
      logger.log(webProperties);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private setFriendlyPermissions(response: any[]): RoleAssignment[] {
    response.forEach((r: any) => {
      r.RoleDefinitionBindings.forEach((r: RoleDefinition) => {
        const permissions: BasePermissions = new BasePermissions();
        permissions.high = r.BasePermissions.High as number;
        permissions.low = r.BasePermissions.Low as number;
        r.BasePermissionsValue = permissions.parse();
        r.RoleTypeKindValue = RoleType[r.RoleTypeKind];
      });
    });

    return response;
  }
}

module.exports = new SpoWebGetCommand();