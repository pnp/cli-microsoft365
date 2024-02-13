import { Group, TeamsAsyncOperation } from '@microsoft/microsoft-graph-types';
import { setTimeout } from 'timers/promises';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description?: string;
  name?: string;
  template?: string;
  wait?: boolean;
  ownerUserNames?: string;
  ownerIds?: string;
  ownerEmails?: string;
  memberUserNames?: string;
  memberIds?: string;
  memberEmails?: string;
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
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        template: typeof args.options.template !== 'undefined',
        wait: !!args.options.wait,
        ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
        ownerIds: typeof args.options.ownerIds !== 'undefined',
        ownerEmails: typeof args.options.ownerEmails !== 'undefined',
        memberUserNames: typeof args.options.memberUserNames !== 'undefined',
        memberIds: typeof args.options.memberIds !== 'undefined',
        memberEmails: typeof args.options.memberEmails !== 'undefined'
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
      },
      {
        option: '--ownerUserNames [ownerUserNames]'
      },
      {
        option: '--ownerIds [ownerIds]'
      },
      {
        option: '--ownerEmails [ownerEmails]'
      },
      {
        option: '--memberUserNames [memberUserNames]'
      },
      {
        option: '--memberIds [memberIds]'
      },
      {
        option: '--memberEmails [memberEmails]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.ownerUserNames) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.ownerUserNames.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `Owner username '${isValidUserPrincipalNameArray}' is invalid for option 'ownerUserNames'.`;
          }
        }

        if (args.options.ownerEmails) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.ownerEmails.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `Owner email '${isValidUserPrincipalNameArray}' is invalid for option 'ownerEmails'.`;
          }
        }

        if (args.options.ownerIds && !validation.isValidGuidArray(args.options.ownerIds.split(','))) {
          return `The option 'ownerIds' contains one or more invalid GUIDs.`;
        }

        if (args.options.memberUserNames) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.memberUserNames.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `Member username '${isValidUserPrincipalNameArray}' is invalid for option 'memberUserNames'.`;
          }
        }

        if (args.options.memberEmails) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.memberEmails.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `Member email '${isValidUserPrincipalNameArray}' is invalid for option 'memberEmails'.`;
          }
        }

        if (args.options.memberIds && !validation.isValidGuidArray(args.options.memberIds.split(','))) {
          return `The option 'memberIds' contains one or more invalid GUIDs`;
        }

        return true;
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
      },
      {
        options: ['ownerUserNames', 'ownerIds', 'ownerEmails'],
        runsWhen: (args) => {
          return args.options.ownerUserNames || args.options.ownerIds || args.options.ownerEmails;
        }
      },
      {
        options: ['memberUserNames', 'memberIds', 'memberEmails'],
        runsWhen: (args) => {
          return args.options.memberUserNames || args.options.memberIds || args.options.memberEmails;
        }
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);

    if (isAppOnlyAccessToken && !args.options.ownerUserNames && !args.options.ownerIds && !args.options.ownerEmails) {
      this.handleError(`Specify at least 'ownerUserNames', 'ownerIds' or 'ownerEmails' when using application permissions.`);
    }

    let requestBody: any;
    if (args.options.template) {
      if (this.verbose) {
        await logger.logToStderr(`Using template...`);
      }
      requestBody = JSON.parse(args.options.template);

      if (args.options.name) {
        if (this.verbose) {
          await logger.logToStderr(`Using '${args.options.name}' as name...`);
        }
        requestBody.displayName = args.options.name;
      }

      if (args.options.description) {
        if (this.verbose) {
          await logger.logToStderr(`Using '${args.options.description}' as description...`);
        }
        requestBody.description = args.options.description;
      }
    }
    else {
      if (this.verbose) {
        await logger.logToStderr(`Creating team with name ${args.options.name} and description ${args.options.description}`);
      }

      requestBody = {
        'template@odata.bind': `https://graph.microsoft.com/v1.0/teamsTemplates('standard')`,
        displayName: args.options.name,
        description: args.options.description
      };
    }

    let members: TeamMember[] = [];

    if (args.options.ownerEmails || args.options.ownerIds || args.options.ownerUserNames) {
      members = await this.retrieveMembersToAdd(members, 'owner', args.options.ownerEmails, args.options.ownerIds, args.options.ownerUserNames);
    }

    if (args.options.memberEmails || args.options.memberIds || args.options.memberUserNames) {
      members = await this.retrieveMembersToAdd(members, 'member', args.options.memberEmails, args.options.memberIds, args.options.memberUserNames);
    }

    if (members.length > 0 && members.filter(y => y.roles.includes('owner')).length > 0 && accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      const groupOwner = members.filter(y => y.roles.includes('owner')).slice(0, 1);
      members = members.filter(y => y !== groupOwner[0]);
      requestBody.members = groupOwner;
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

      if (!args.options.wait && members.length === 0) {
        await logger.log(teamsAsyncOperation);
      }
      else {
        await this.waitUntilTeamFinishedProvisioning(teamsAsyncOperation, requestOptions, logger);
        const entraGroup = await this.getEntraGroup(teamsAsyncOperation.targetResourceId!, logger);
        if (members.length > 0) {
          if (this.verbose) {
            await logger.logToStderr('Adding members to the team...');
          }
          await this.addMembers(members, entraGroup.id!);
        }
        await logger.log(entraGroup);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addMembers(members: TeamMember[], groupId: string): Promise<void> {
    for (const member of members) {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${groupId}/members`,
        headers: {
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: member
      };
      await request.post(requestOptions);
    }
  }

  private async waitUntilTeamFinishedProvisioning(teamsAsyncOperation: TeamsAsyncOperation, requestOptions: CliRequestOptions, logger: Logger): Promise<void> {
    if (teamsAsyncOperation.status === 'succeeded') {
      if (this.verbose) {
        await logger.logToStderr('Team provisioned succesfully. Returning...');
      }
      return;
    }
    else if (teamsAsyncOperation.error) {
      throw teamsAsyncOperation.error;
    }
    else {
      if (this.verbose) {
        await logger.logToStderr(`Team still provisioning. Retrying in ${this.pollingInterval / 1000} seconds...`);
      }
      await setTimeout(this.pollingInterval);
      teamsAsyncOperation = await request.get<TeamsAsyncOperation>(requestOptions);
      await this.waitUntilTeamFinishedProvisioning(teamsAsyncOperation, requestOptions, logger);
    }
  }

  private async getEntraGroup(id: string, logger: Logger): Promise<Group> {
    let group: Group;

    try {
      group = await entraGroup.getGroupById(id);
    }
    catch {
      if (this.verbose) {
        await logger.logToStderr(`Error occurred on retrieving the aad group. Retrying in ${this.pollingInterval / 1000} seconds.`);
      }
      await setTimeout(this.pollingInterval);
      return await this.getEntraGroup(id, logger);
    }

    return group!;
  }

  private async retrieveMembersToAdd(members: TeamMember[], role: string, emails?: string, ids?: string, userNames?: string): Promise<any[]> {
    let itemsToProcess: string[] = [];

    if (emails) {
      for (const email of emails.split(',')) {
        const userId = await entraUser.getUserIdByEmail(email);
        itemsToProcess.push(userId);
      }
    }
    else if (ids) {
      itemsToProcess = ids.split(',');
    }
    else if (userNames) {
      for (const un of userNames.split(',')) {
        const userId = await entraUser.getUserIdByUpn(un);
        itemsToProcess.push(userId);
      }
    }

    itemsToProcess.map((item: string) => {
      if (!members.some((y: TeamMember) => y['user@odata.bind'] === `https://graph.microsoft.com/v1.0/users('${item}')`)) {
        members.push({
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${item}')`,
          roles: [role]
        });
      }
      else {
        const returnObject = members.find((y: TeamMember) => y['user@odata.bind'] === `https://graph.microsoft.com/v1.0/users('${item}')`);
        returnObject?.roles.push(role);
      }
    });

    return members;
  }
}

export interface TeamMember {
  "@odata.type": string
  roles: string[]
  "user@odata.bind": string
}

export default new TeamsTeamAddCommand();