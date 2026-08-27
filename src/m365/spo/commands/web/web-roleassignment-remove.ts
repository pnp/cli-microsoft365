import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
import { cli } from '../../../../cli/cli.js';

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
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from web permissions';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(options => [options.principalId, options.upn, options.groupName, options.entraGroupId, options.entraGroupName].filter(x => x !== undefined).length === 1, {
      error: `Specify either 'principalId', 'upn', 'groupName', 'entraGroupId', or 'entraGroupName'.`,
      params: {
        customCode: 'optionSet',
        options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRoleAssignment(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from web ${args.options.webUrl}?` });

      if (result) {
        await this.removeRoleAssignment(logger, args.options);
      }
    }
  }

  private async removeRoleAssignment(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing role assignment from web ${options.webUrl}...`);
    }

    try {
      if (options.upn) {
        const principalId = await this.getUserPrincipalId(options, logger);
        await this.removeRoleAssignmentWithOptions(options.webUrl, principalId, logger);
      }
      else if (options.groupName) {
        const principalId = await this.getGroupPrincipalId(options, logger);
        await this.removeRoleAssignmentWithOptions(options.webUrl, principalId, logger);
      }
      else if (options.entraGroupId || options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        const group = options.entraGroupId
          ? await entraGroup.getGroupById(options.entraGroupId)
          : await entraGroup.getGroupByDisplayName(options.entraGroupName!);

        const siteUser = await spo.ensureEntraGroup(options.webUrl, group);
        await this.removeRoleAssignmentWithOptions(options.webUrl, siteUser.Id, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeRoleAssignmentWithOptions(webUrl: string, principalId: number, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Removing role assignment...');
    }

    const requestOptions: any = {
      url: `${webUrl}/_api/web/roleassignments/removeroleassignment(principalid='${principalId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
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

export default new SpoWebRoleAssignmentRemoveCommand();