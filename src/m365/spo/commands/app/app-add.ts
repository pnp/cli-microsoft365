import fs from 'fs';
import path from 'path';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  filePath: z.string()
    .refine(filePath => fs.existsSync(path.resolve(filePath)), {
      error: e => `File '${path.resolve(e.input as string)}' not found`,
      abort: true
    })
    .refine(filePath => !fs.lstatSync(path.resolve(filePath)).isDirectory(), {
      error: e => `Path '${path.resolve(e.input as string)}' points to a directory`
    })
    .alias('p'),
  overwrite: z.boolean().optional(),
  appCatalogScope: zod.coercedEnum({ tenant: 'tenant', sitecollection: 'sitecollection' }).optional().alias('s'),
  appCatalogUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .optional()
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppAddCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_ADD;
  }

  public get description(): string {
    return 'Adds an app to the specified SharePoint Online app catalog';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.appCatalogScope !== 'sitecollection' || opts.appCatalogUrl !== undefined, {
        error: 'You must specify appCatalogUrl when appCatalogScope is sitecollection',
        path: ['appCatalogUrl'],
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
    const overwrite: boolean = args.options.overwrite || false;

    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

      const fullPath: string = path.resolve(args.options.filePath);
      if (this.verbose) {
        await logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
      }

      const fileName: string = path.basename(fullPath);
      const requestOptions: CliRequestOptions = {
        url: `${appCatalogUrl}/_api/web/${scope}appcatalog/Add(overwrite=${(overwrite.toString().toLowerCase())}, url='${fileName}')`,
        headers: {
          accept: 'application/json;odata=nometadata',
          binaryStringRequestBody: 'true'
        },
        data: fs.readFileSync(fullPath)
      };

      const res = await request.post<string>(requestOptions);

      const json: { UniqueId: string; } = JSON.parse(res);
      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(json);
      }
      else {
        await logger.log(json.UniqueId);
      }
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }
}

export default new SpoAppAddCommand();