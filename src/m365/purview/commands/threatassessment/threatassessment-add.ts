import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import auth from '../../../../Auth.js';
import fs from 'fs';
import path from 'path';

const allowedTypes = ['file', 'url'] as const;
const allowedExpectedAssessments = ['block', 'unblock'] as const;
const allowedCategories = ['spam', 'phishing', 'malware'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: z.enum(allowedTypes).alias('t'),
  expectedAssessment: z.enum(allowedExpectedAssessments).alias('e'),
  category: z.enum(allowedCategories).alias('c'),
  path: z.string().refine(val => fs.existsSync(val), {
    error: 'Specified file does not exist.'
  }).optional().alias('p'),
  url: z.string().optional().alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewThreatAssessmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.THREATASSESSMENT_ADD;
  }

  public get description(): string {
    return 'Create a threat assessment';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => {
        if (opts.type === 'file') {
          return opts.path !== undefined;
        }
        return true;
      }, {
        error: `'path' is required when type is 'file'.`,
        path: ['path'],
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => {
        if (opts.type === 'url') {
          return opts.url !== undefined;
        }
        return true;
      }, {
        error: `'url' is required when type is 'url'.`,
        path: ['url'],
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {

      if (accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)) {
        throw 'This command currently does not support app only permissions.';
      }

      if (this.verbose) {
        await logger.logToStderr(`Adding threat assessment of type ${args.options.type} with expected assessment ${args.options.expectedAssessment} and category ${args.options.category}`);
      }

      const requestBody: any = {
        expectedAssessment: args.options.expectedAssessment,
        category: args.options.category,
        url: args.options.url,
        contentData: args.options.path && fs.readFileSync(args.options.path).toString('base64'),
        fileName: args.options.path && path.basename(args.options.path)
      };

      switch (args.options.type) {
        case 'file':
          requestBody['@odata.type'] = '#microsoft.graph.fileAssessmentRequest';
          break;
        case 'url':
          requestBody['@odata.type'] = '#microsoft.graph.urlAssessmentRequest';
          break;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/informationProtection/threatAssessmentRequests`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: requestBody,
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }
}

export default new PurviewThreatAssessmentAddCommand();