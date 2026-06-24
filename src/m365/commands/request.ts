import { AxiosRequestConfig, AxiosResponse, RawAxiosRequestHeaders } from 'axios';
import fs from 'fs';
import path from 'path';
import { z } from 'zod';
import auth from '../../Auth.js';
import Command, { globalOptionsZod } from '../../Command.js';
import { Logger } from '../../cli/Logger.js';
import request from '../../request.js';
import commands from './commands.js';

const allowedMethods = ['get', 'post', 'put', 'patch', 'delete', 'head', 'options'] as const;

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  url: z.string().alias('u'),
  method: z.enum(allowedMethods).default('get').alias('m'),
  resource: z.string().optional().alias('r'),
  body: z.string().optional().alias('b'),
  filePath: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class RequestCommand extends Command {
  public get name(): string {
    return commands.REQUEST;
  }

  public get description(): string {
    return 'Executes the specified web request using CLI for Microsoft 365';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.body || (opts.method !== 'get'), {
        error: 'Specify a different method when using body'
      })
      .refine(opts => !opts.body || opts['content-type'], {
        error: 'Specify the content-type when using body'
      })
      .refine(opts => !opts.filePath || fs.existsSync(path.dirname(opts.filePath)), {
        error: 'The location specified in the filePath does not exist'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.debug) {
      await logger.logToStderr(`Preparing request...`);
    }

    try {
      const url = this.resolveUrlTokens(args.options.url);
      const method = args.options.method.toUpperCase();
      const headers: RawAxiosRequestHeaders = {};

      this.addUnknownOptionsToPayloadZod(headers, args.options);

      if (!headers.accept) {
        headers.accept = 'application/json';
      }

      if (args.options.resource) {
        headers['x-resource'] = args.options.resource;
      }

      const config: AxiosRequestConfig<string> = {
        url: url,
        headers,
        method,
        data: args.options.body
      };

      if (headers.accept.toString().startsWith('application/json')) {
        config.responseType = 'json';
      }

      if (args.options.filePath) {
        config.responseType = 'stream';
      }

      if (this.verbose) {
        await logger.logToStderr(`Executing request...`);
      }


      if (args.options.filePath) {
        const file: AxiosResponse = await request.execute<AxiosResponse>(config);
        const filePath: string = await new Promise((resolve, reject) => {
          const writer = fs.createWriteStream(args.options.filePath as string);

          file.data.pipe(writer);

          writer.on('error', err => {
            reject(err);
          });
          writer.on('close', () => {
            resolve(args.options.filePath as string);
          });
        });

        if (this.verbose) {
          await logger.logToStderr(`File saved to path ${filePath}`);
        }
      }
      else {
        const res = await request.execute<string>(config);
        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleError(err);
    }
  }

  private resolveUrlTokens(url: string): string {
    if (url.startsWith('@graphbeta')) {
      return url.replace('@graphbeta', 'https://graph.microsoft.com/beta');
    }
    if (url.startsWith('@graph')) {
      return url.replace('@graph', 'https://graph.microsoft.com/v1.0');
    }
    if (url.startsWith('@spo')) {
      if (auth.connection.spoUrl) {
        return url.replace('@spo', auth.connection.spoUrl);
      }

      throw `SharePoint root site URL is unknown. Please set your SharePoint URL using command 'spo set'.`;
    }

    return url;
  }
}

export default new RequestCommand();