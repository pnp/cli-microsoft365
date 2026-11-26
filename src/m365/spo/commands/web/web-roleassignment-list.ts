import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import request from '../../../../request.js';
import { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(value => validation.isValidSharePointUrl(value) === true, {
      error: issue => `'${issue.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebRoleAssignmentListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists all role assignments for a specific web.';
  }

  public defaultProperties(): string[] | undefined {
    return ['PrincipalId', 'Member', 'RoleDefinitionBindings'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving role assignments for web ${args.options.webUrl}...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/RoleAssignments?$expand=Member,RoleDefinitionBindings`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const response = await request.get<{ value: any[] }>(requestOptions);
      response.value.forEach((r: any) => {
        r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
      });

      if (args.options.output && args.options.output !== 'json' && !args.options.query) {
        await logger.log(response.value.map(assignment => ({
          PrincipalId: assignment.PrincipalId,
          Member: assignment.Member.Title,
          RoleDefinitionBindings: assignment.RoleDefinitionBindings.map((role: any) => role.Name).join(', ')
        })));
      }
      else {
        await logger.log(response.value);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebRoleAssignmentListCommand();
