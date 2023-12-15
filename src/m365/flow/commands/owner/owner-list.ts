import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';

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
  flowName: string;
  environmentName: string;
  asAdmin?: boolean;
}

class FlowOwnerListCommand extends PowerAutomateCommand {
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
        option: '--flowName <flowName>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.flowName)) {
          return `${args.options.flowName} is not a valid GUID.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Listing owners for flow ${args.options.flowName} in environment ${args.options.environmentName}`);
      }

      const response = await odata.getAllItems<FlowPermissionResponse>(`${this.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/permissions?api-version=2016-11-01`);
      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(response);
      }
      else {
        //converted to text friendly output
        await logger.log(response.map(res => ({
          roleName: res.properties.roleName,
          id: res.properties.principal.id,
          type: res.properties.principal.type
        })));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowOwnerListCommand();