import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
import { AxiosRequestConfig } from 'axios';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}

class TeamsTeamAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_APP_LIST;
  }

  public get description(): string {
    return 'List apps installed in the specified team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'distributionMethod'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const group = await aadGroup.getGroupByDisplayName(args.options.name!);
    if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
      throw `The specified team does not exist in the Microsoft Teams`;
    }

    return group.id!;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId: string = await this.getTeamId(args);
      const requestOptions: AxiosRequestConfig = {
        url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/installedApps?$expand=teamsApp,teamsAppDefinition`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      const res = await odata.get<>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsTeamAppListCommand();