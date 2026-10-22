import { z } from 'zod';
import Command, { globalOptionsZod } from '../../../../Command.js';
import type GlobalOptions from '../../../../GlobalOptions.js';
import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
import paAppListCommand from '../app/app-list.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional()
    .alias('n'),
  displayName: z.string().optional().alias('d'),
  environmentName: z.string().optional().alias('e'),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppGetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Power App';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.name, opts.displayName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'name' or 'displayName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['name', 'displayName']
        }
      })
      .refine(opts => !opts.asAdmin || opts.environmentName, {
        message: 'When specifying the asAdmin option, the environment option is required as well.'
      })
      .refine(opts => !opts.environmentName || opts.asAdmin, {
        message: 'When specifying the environment option, the asAdmin option is required as well.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.name) {
        let endpoint = `${this.resource}/providers/Microsoft.PowerApps`;
        if (args.options.asAdmin) {
          endpoint += `/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName!)}`;
        }
        endpoint += `/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`;

        const requestOptions: CliRequestOptions = {
          url: endpoint,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        if (this.verbose) {
          await logger.logToStderr(`Retrieving information about Microsoft Power App with name '${args.options.name}'...`);
        }

        const res = await request.get<any>(requestOptions);
        await logger.log(this.setProperties(res));
      }
      else {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving information about Microsoft Power App with displayName '${args.options.displayName}'...`);
        }

        const getAppsOutput = await this.getApps(args, logger);
        if (getAppsOutput.stdout && JSON.parse(getAppsOutput.stdout).length > 0) {
          const allApps: any[] = JSON.parse(getAppsOutput.stdout);
          const app = allApps.find((a: any) => {
            return a.properties.displayName.toLowerCase() === `${args.options.displayName}`.toLowerCase();
          });

          if (app) {
            await logger.log(this.setProperties(app));
          }
          else {
            throw `No app found with displayName '${args.options.displayName}'.`;
          }
        }
        else {
          throw 'No apps found.';
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getApps(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all apps...`);
    }

    const options: GlobalOptions = {
      output: 'json',
      debug: this.debug,
      verbose: this.verbose,
      environmentName: args.options.environmentName,
      asAdmin: args.options.asAdmin
    };

    return await cli.executeCommandWithOutput(paAppListCommand as Command, { options: { ...options, _: [] } });
  }

  private setProperties(app: any): any {
    app.displayName = app.properties.displayName;
    app.description = app.properties.description || '';
    app.appVersion = app.properties.appVersion;
    app.owner = app.properties.owner.email || '';
    return app;
  }
}

export default new PaAppGetCommand();