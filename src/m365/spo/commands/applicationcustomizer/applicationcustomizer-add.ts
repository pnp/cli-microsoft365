import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CustomAction } from '../customaction/customaction.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  title: z.string().alias('t'),
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    error: e => `${e.input} is not a valid SharePoint URL`
  }).alias('u'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not a valid GUID`
  }).alias('i'),
  description: z.string().optional(),
  clientSideComponentProperties: z.string().refine(val => {
    try {
      JSON.parse(val);
      return true;
    }
    catch {
      return false;
    }
  }, {
    error: 'Specified clientSideComponentProperties is not a valid JSON string.'
  }).optional(),
  hostProperties: z.string().refine(val => {
    try {
      JSON.parse(val);
      return true;
    }
    catch {
      return false;
    }
  }, {
    error: 'Specified hostProperties is not a valid JSON string.'
  }).optional(),
  scope: z.enum(['Site', 'Web']).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoApplicationCustomizerAddCommand extends SpoCommand {
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_ADD;
  }

  public get description(): string {
    return 'Adds an application customizer to a site.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding application customizer with title '${args.options.title}' and clientSideComponentId '${args.options.clientSideComponentId}' to the site`);
    }

    const requestBody: any = {
      Title: args.options.title,
      Name: args.options.title,
      Description: args.options.description,
      Location: 'ClientSideExtension.ApplicationCustomizer',
      ClientSideComponentId: args.options.clientSideComponentId,
      HostProperties: args.options.hostProperties || ''
    };

    if (args.options.clientSideComponentProperties) {
      requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
    }

    const scope = args.options.scope || 'Site';

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/${scope}/UserCustomActions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    await request.post<CustomAction>(requestOptions);
  }
}

export default new SpoApplicationCustomizerAddCommand();