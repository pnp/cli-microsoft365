import GlobalOptions from '../../../../GlobalOptions.js';
import { Organization } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { cli } from '../../../../cli/cli.js';

interface MultitenantOrganizationMember {
  tenantId?: string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force?: boolean;
}

class EntraMultitenantRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.MULTITENANT_REMOVE;
  }

  public get description(): string {
    return 'Removes a multitenant organization';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --force'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    const removeMultitenantOrg = async (): Promise<void> => {
      try {
        const tenantId = await this.getCurrentTenantId();

        let tenantsId = await this.getAllTenantsIds();
        const tenantsCount = tenantsId.length;

        if (tenantsCount > 0) {
          const tasks = tenantsId
            .filter(x => x !== tenantId)
            .map(t => this.removeTenant(logger, t));
          await Promise.all(tasks);

          do {
            if (this.verbose) {
              await logger.logToStderr(`Waiting 30 seconds...`);
            }

            await new Promise(resolve => setTimeout(resolve, 30000));

            // from current behavior, removing tenant can take a few seconds
            // current tenant must be removed once all previous ones were removed
            if (this.verbose) {
              await logger.logToStderr(`Checking all tenants were removed...`);
            }

            tenantsId = await this.getAllTenantsIds();

            if (this.verbose) {
              await logger.logToStderr(`Number of removed tenants: ${tenantsCount - tenantsId.length}`);
            }
          }
          while (tenantsId.length !== 1);

          // current tenant must be removed as the last one
          await this.removeTenant(logger, tenantId);
          await logger.logToStderr('Your Multi-Tenant organization is being removed; this can take up to 2 hours.')
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeMultitenantOrg();
    }
    else {
      const result = await cli.promptForConfirmation({ message: 'Are you sure you want to remove multitenant organization?' });

      if (result) {
        await removeMultitenantOrg();
      }
    }
  }

  private async getAllTenantsIds(): Promise<string[]> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization/tenants?$select=tenantId`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const tenants = await odata.getAllItems<MultitenantOrganizationMember>(requestOptions);
    return tenants.map(x => x.tenantId!);
  }

  private async getCurrentTenantId(): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/organization?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    const organizations = await odata.getAllItems<Organization>(requestOptions);
    return organizations[0].id!;
  }

  private async removeTenant(logger: Logger, tenantId: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing tenant: ${tenantId}`);
    }
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization/tenants/${tenantId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    await request.delete(requestOptions);
  }
}

export default new EntraMultitenantRemoveCommand();