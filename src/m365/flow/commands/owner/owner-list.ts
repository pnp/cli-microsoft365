import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface FlowPermissionResponse {
  name: string;
  id: string;
  type: string;
  properties: FlowPermissionProperties;
}

interface FlowPermissionProperties {
  roleName: string;
  permissionType: string;
  principal: FlowPermissionPrincipal;
}

interface FlowPermissionPrincipal {
  id: string;
  type: string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  environmentName: string;
  asAdmin?: boolean;
}

class FlowOwnerListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.OWNER_LIST;
  }

  public get description(): string {
    return 'Lists all owners of a Power Automate flow';
  }

  public defaultProperties(): string[] | undefined {
    return ['roleName', 'id', 'type'];
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
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Listing owners for flow ${args.options.name} in environment ${args.options.environmentName}`);
      }

      const response = await odata.getAllItems<FlowPermissionResponse>(`${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}/permissions?api-version=2016-11-01`);
      if (!args.options.output || !Cli.shouldTrimOutput(args.options.output)) {
        logger.log(response);
      }
      else {
        //converted to text friendly output
        logger.log(response.map(res => {
          return {
            roleName: res.properties.roleName,
            id: res.properties.principal.id,
            type: res.properties.principal.type
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new FlowOwnerListCommand();