import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  principalId: z.number().optional(),
  upn: z.string().optional(),
  groupName: z.string().optional(),
  entraGroupId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID for option entraGroupId.`
  }).optional(),
  entraGroupName: z.string().optional(),
  roleDefinitionId: z.number().optional(),
  roleDefinitionName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to web';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.principalId, options.upn, options.groupName, options.entraGroupId, options.entraGroupName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'principalId', 'upn', 'groupName', 'entraGroupId', or 'entraGroupName'.`,
        params: {
          customCode: 'optionSet',
          options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName']
        }
      })
      .refine(options => [options.roleDefinitionId, options.roleDefinitionName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'roleDefinitionId' or 'roleDefinitionName'.`,
        params: {
          customCode: 'optionSet',
          options: ['roleDefinitionId', 'roleDefinitionName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding role assignment to web ${args.options.webUrl}...`);
    }

    try {
      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);

      if (args.options.upn) {
        const principalId = await this.getUserPrincipalId(args.options, logger);
        await this.addRoleAssignment(args.options.webUrl, principalId, roleDefinitionId, logger);
      }
      else if (args.options.groupName) {
        const principalId = await this.getGroupPrincipalId(args.options, logger);
        await this.addRoleAssignment(args.options.webUrl, principalId, roleDefinitionId, logger);
      }
      else if (args.options.entraGroupId || args.options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        const group = args.options.entraGroupId
          ? await entraGroup.getGroupById(args.options.entraGroupId)
          : await entraGroup.getGroupByDisplayName(args.options.entraGroupName!);

        const siteUser = await spo.ensureEntraGroup(args.options.webUrl, group);
        await this.addRoleAssignment(args.options.webUrl, siteUser.Id, roleDefinitionId, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleAssignment(webUrl: string, principalId: number, roleDefinitionId: number, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Adding role assignment...');
    }

    const requestOptions: any = {
      url: `${webUrl}/_api/web/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async getRoleDefinitionId(options: Options, logger: Logger): Promise<number> {
    if (!options.roleDefinitionName) {
      return options.roleDefinitionId as number;
    }

    const roledefinition = await spo.getRoleDefinitionByName(options.webUrl, options.roleDefinitionName, logger, this.verbose);

    return roledefinition.Id;
  }

  private async getGroupPrincipalId(options: Options, logger: Logger): Promise<number> {
    const group = await spo.getGroupByName(options.webUrl, options.groupName!, logger, this.verbose);
    return group.Id;
  }

  private async getUserPrincipalId(options: Options, logger: Logger): Promise<number> {
    const user = await spo.getUserByEmail(options.webUrl, options.upn!, logger, this.verbose);
    return user.Id;
  }
}

export default new SpoWebRoleAssignmentAddCommand();