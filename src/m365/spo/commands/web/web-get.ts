import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { WebProperties } from './WebProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  withGroups: z.boolean().optional(),
  withPermissions: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebGetCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_GET;
  }

  public get description(): string {
    return 'Retrieve information about the specified site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let url: string = `${args.options.url}/_api/web`;
    if (args.options.withGroups) {
      url += '?$expand=AssociatedMemberGroup,AssociatedOwnerGroup,AssociatedVisitorGroup';
    }
    const requestOptions: any = {
      url,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const webProperties: WebProperties = await request.get<WebProperties>(requestOptions);
      if (args.options.withPermissions) {
        requestOptions.url = `${args.options.url}/_api/web/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
        const response = await request.get<{ value: any[] }>(requestOptions);
        response.value.forEach((r: any) => {
          r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
        });
        webProperties.RoleAssignments = response.value;
      }
      await logger.log(webProperties);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebGetCommand();