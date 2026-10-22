import fs from 'fs';
import path from 'path';
import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import request, { CliRequestOptions } from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import PowerPlatformCommand from '../../base/PowerPlatformCommand.js';
import commands from '../commands.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  name: z.uuid().alias('n'),
  packageDisplayName: z.string().optional().alias('d'),
  packageDescription: z.string().optional(),
  packageCreatedBy: z.string().optional().alias('c'),
  packageSourceEnvironment: z.string().optional().alias('s'),
  format: z.string().transform(v => v.toLowerCase()).pipe(z.enum(['json', 'zip'], {
    error: 'Option format must be json or zip. Default is zip'
  })).optional().alias('f'),
  path: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowExportCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.EXPORT;
  }

  public get description(): string {
    return 'Exports the specified Microsoft Flow as a file';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => options.format !== 'json' || !options.packageCreatedBy, {
        error: 'packageCreatedBy cannot be specified with output of json',
        path: ['packageCreatedBy']
      })
      .refine(options => options.format !== 'json' || !options.packageDescription, {
        error: 'packageDescription cannot be specified with output of json',
        path: ['packageDescription']
      })
      .refine(options => options.format !== 'json' || !options.packageDisplayName, {
        error: 'packageDisplayName cannot be specified with output of json',
        path: ['packageDisplayName']
      })
      .refine(options => options.format !== 'json' || !options.packageSourceEnvironment, {
        error: 'packageSourceEnvironment cannot be specified with output of json',
        path: ['packageSourceEnvironment']
      })
      .refine(options => {
        if (options.path) {
          return fs.existsSync(path.dirname(options.path));
        }

        return true;
      }, {
        error: 'Specified path where to save the file does not exist',
        path: ['path']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const formatArgument = args.options.format?.toLowerCase() || '';

    if (this.verbose) {
      await logger.logToStderr(`Retrieving package resources for Microsoft Flow ${args.options.name}...`);
    }

    try {
      let res: any;
      if (formatArgument === 'json') {
        if (this.verbose) {
          await logger.logToStderr('format = json, skipping listing package resources step.');
        }
      }
      else {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/listPackageResources?api-version=2016-11-01`,
          headers: {
            accept: 'application/json'
          },
          data: {
            "baseResourceIds": [
              `/providers/Microsoft.Flow/flows/${args.options.name}`
            ]
          },
          responseType: 'json'
        };

        res = await request.post<any>(requestOptions);
      }

      if (typeof res !== 'undefined' && res.errors && res.errors.length && res.errors.length > 0) {
        throw res.errors[0].message;
      }

      if (this.verbose) {
        await logger.logToStderr(`Initiating package export for Microsoft Flow ${args.options.name}...`);
      }

      let requestOptions: CliRequestOptions = {
        url: formatArgument === 'json' ?
          `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`
          : `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/exportPackage?api-version=2016-11-01`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      if (formatArgument !== 'json') {
        // adds suggestedCreationType property to all resources
        // see https://github.com/pnp/cli-microsoft365/issues/1845
        Object.keys(res.resources).forEach((key) => {
          if (res.resources[key].type === 'Microsoft.Flow/flows') {
            res.resources[key].suggestedCreationType = 'Update';
          }
          else {
            res.resources[key].suggestedCreationType = 'Existing';
          }
        });

        requestOptions.data = {
          includedResourceIds: [
            `/providers/Microsoft.Flow/flows/${args.options.name}`
          ],
          details: {
            displayName: args.options.packageDisplayName,
            description: args.options.packageDescription,
            creator: args.options.packageCreatedBy,
            sourceEnvironment: args.options.packageSourceEnvironment
          },
          resources: res.resources
        };
      }

      res = formatArgument === 'json' ? await request.get(requestOptions) : await request.post(requestOptions);

      if (this.verbose) {
        await logger.logToStderr(`Getting file for Microsoft Flow ${args.options.name}...`);
      }

      const downloadFileUrl: string = formatArgument === 'json' ? '' : res.packageLink.value;
      const filenameRegEx: RegExp = /([^/]+\.zip)/i;
      let filenameFromApi = formatArgument === 'json' ? `${res.properties.displayName}.json` : (filenameRegEx.exec(downloadFileUrl) || ['output.zip'])[0];
      // Replace all illegal characters from the file name
      const illegalCharsRegEx = /[\\/:*?"<>|]/g;
      filenameFromApi = filenameFromApi.replace(illegalCharsRegEx, '_');

      if (this.verbose) {
        await logger.logToStderr(`Filename from PowerApps API: ${filenameFromApi}.`);
        await logger.logToStderr('');
      }

      requestOptions = {
        url: formatArgument === 'json' ?
          `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}/exportToARMTemplate?api-version=2016-11-01`
          : downloadFileUrl,
        // Set responseType to arraybuffer, otherwise binary data will be encoded
        // to utf8 and binary data is corrupt
        responseType: 'arraybuffer',
        headers: formatArgument === 'json' ?
          {
            accept: 'application/json'
          } : {
            'x-anonymous': true
          }
      };

      const file = formatArgument === 'json' ?
        await request.post<string>(requestOptions)
        : await request.get<string>(requestOptions);

      const path = args.options.path ? args.options.path : `./${filenameFromApi}`;

      fs.writeFileSync(path, file, 'binary');
      if (!args.options.path || this.verbose) {
        if (this.verbose) {
          await logger.logToStderr(`File saved to path '${path}'.`);
        }
        else {
          await logger.log(path);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowExportCommand();