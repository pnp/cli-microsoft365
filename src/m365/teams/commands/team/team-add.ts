import { Group, TeamsAsyncOperation } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { setTimeout } from "timers/promises";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description?: string;
  name?: string;
  template?: string;
  wait?: boolean;
}

class TeamsTeamAddCommand extends GraphCommand {
  private pollingInterval: number = 30_000;

  public get name(): string {
    return commands.TEAM_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        template: typeof args.options.template !== 'undefined',
        wait: !!args.options.wait
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--template [template]'
      },
      {
        option: '--wait'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['name'],
        runsWhen: (args) => {
          return !args.options.template;
        }
      },
      {
        options: ['description'],
        runsWhen: (args) => {
          return !args.options.template;
        }
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let requestBody: any;
    if (args.options.template) {
      if (this.verbose) {
        logger.logToStderr(`Using template...`);
      }
      requestBody = JSON.parse(args.options.template);

      if (args.options.name) {
        if (this.verbose) {
          logger.logToStderr(`Using '${args.options.name}' as name...`);
        }
        requestBody.displayName = args.options.name;
      }

      if (args.options.description) {
        if (this.verbose) {
          logger.logToStderr(`Using '${args.options.description}' as description...`);
        }
        requestBody.description = args.options.description;
      }
    }
    else {
      if (this.verbose) {
        logger.logToStderr(`Creating team with name ${args.options.name} and description ${args.options.description}`);
      }

      requestBody = {
        'template@odata.bind': `https://graph.microsoft.com/v1.0/teamsTemplates('standard')`,
        displayName: args.options.name,
        description: args.options.description
      };
    }

    const requestOptionsPost: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'stream'
    };

    try {
      const res = await request.post<any>(requestOptionsPost);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0${res.headers.location}`,
        headers: {
          accept: 'application/json;odata.metadata=minimal'
        },
        responseType: 'json'
      };

      const teamsAsyncOperation: TeamsAsyncOperation = await request.get<TeamsAsyncOperation>(requestOptions);

      if (!args.options.wait) {
        logger.log(teamsAsyncOperation);
      }
      else {
        await this.waitUntilTeamFinishedProvisioning(teamsAsyncOperation, requestOptions, logger);
        const aadGroup = await this.getAadGroup(teamsAsyncOperation.targetResourceId!, logger);
        logger.log(aadGroup);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async waitUntilTeamFinishedProvisioning(teamsAsyncOperation: TeamsAsyncOperation, requestOptions: CliRequestOptions, logger: Logger): Promise<void> {
    if (teamsAsyncOperation.status === 'succeeded') {
      if (this.verbose) {
        logger.logToStderr('Team provisioned succesfully. Returning...');
      }
      return;
    }
    else if (teamsAsyncOperation.error) {
      throw teamsAsyncOperation.error;
    }
    else {
      if (this.verbose) {
        logger.logToStderr(`Team still provisioning. Retrying in ${this.pollingInterval / 1000} seconds...`);
      }
      await setTimeout(this.pollingInterval);
      teamsAsyncOperation = await request.get<TeamsAsyncOperation>(requestOptions);
      await this.waitUntilTeamFinishedProvisioning(teamsAsyncOperation, requestOptions, logger);
    }
  }

  private async getAadGroup(id: string, logger: Logger): Promise<Group> {
    let group: Group;

    try {
      group = await aadGroup.getGroupById(id);
    }
    catch {
      if (this.verbose) {
        logger.logToStderr(`Error occured on retrieving the aad group. Retrying in ${this.pollingInterval / 1000} seconds.`);
      }
      await setTimeout(this.pollingInterval);
      return await this.getAadGroup(id, logger);
    }

    return group!;
  }
}

module.exports = new TeamsTeamAddCommand();