import { Group, TeamsAsyncOperation } from '@microsoft/microsoft-graph-types';
import { setTimeout } from 'timers/promises';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().optional().alias('n'),
  description: z.string().optional().alias('d'),
  template: z.string().optional(),
  wait: z.boolean().optional(),
  ownerUserNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for option 'ownerUserNames': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    })
    .optional(),
  ownerIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for option 'ownerIds': ${validation.isValidGuidArray(e.input as string)}.`
    })
    .optional(),
  ownerEmails: z.string()
    .refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, {
      error: e => `The following user principal names are invalid for option 'ownerEmails': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    })
    .optional(),
  memberUserNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for option 'memberUserNames': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    })
    .optional(),
  memberIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for option 'memberIds': ${validation.isValidGuidArray(e.input as string)}.`
    })
    .optional(),
  memberEmails: z.string()
    .refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, {
      error: e => `The following user principal names are invalid for option 'memberEmails': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTeamAddCommand extends GraphCommand {
  private pollingInterval: number = 30_000;

  public get name(): string {
    return commands.TEAM_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Teams team';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !!opts.template || opts.name !== undefined, {
        message: `Specify 'name' when not using a template.`,
        path: ['name'],
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => !!opts.template || opts.description !== undefined, {
        message: `Specify 'description' when not using a template.`,
        path: ['description'],
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => [opts.ownerUserNames, opts.ownerIds, opts.ownerEmails].filter(o => o !== undefined).length <= 1, {
        message: `Specify only one of the following options: 'ownerUserNames', 'ownerIds' or 'ownerEmails'.`,
        params: {
          customCode: 'optionSet',
          options: ['ownerUserNames', 'ownerIds', 'ownerEmails']
        }
      })
      .refine(opts => [opts.memberUserNames, opts.memberIds, opts.memberEmails].filter(o => o !== undefined).length <= 1, {
        message: `Specify only one of the following options: 'memberUserNames', 'memberIds' or 'memberEmails'.`,
        params: {
          customCode: 'optionSet',
          options: ['memberUserNames', 'memberIds', 'memberEmails']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)!;

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
      await this.retrieveMembersToAdd(members, 'owner', args.options.ownerEmails, args.options.ownerIds, args.options.ownerUserNames);
    }

    if (args.options.memberEmails || args.options.memberIds || args.options.memberUserNames) {
      await this.retrieveMembersToAdd(members, 'member', args.options.memberEmails, args.options.memberIds, args.options.memberUserNames);
    }

    // We filter out the first owner here and add it to the request body when we are using application only permissions. This is required or the Graph API will throw an error.
    if (members.length > 0 && members.filter(y => y.roles.includes('owner')).length > 0 && isAppOnlyAccessToken) {
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

  private async retrieveMembersToAdd(members: TeamMember[], role: string, emails?: string, ids?: string, userNames?: string): Promise<void> {
    let itemsToProcess: string[] = [];

    if (emails) {
      itemsToProcess = await entraUser.getUserIdsByEmails(formatting.splitAndTrim(emails));
    }
    else if (ids) {
      itemsToProcess = formatting.splitAndTrim(ids);
    }
    else if (userNames) {
      itemsToProcess = await entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
    }

    itemsToProcess.map((item: string) => {
      const member = members.find((y: TeamMember) => y['user@odata.bind'] === `https://graph.microsoft.com/v1.0/users('${item}')`);

      if (!member) {
        members.push({
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${item}')`,
          roles: [role]
        });
      }
      else {
        member.roles.push(role);
      }
    });
  }
}

export interface TeamMember {
  "@odata.type": string
  roles: string[]
  "user@odata.bind": string
}

export default new TeamsTeamAddCommand();