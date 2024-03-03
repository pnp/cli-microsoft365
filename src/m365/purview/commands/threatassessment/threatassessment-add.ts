import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import auth from '../../../../Auth.js';
import fs from 'fs';
import path from 'path';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
  expectedAssessment: string;
  category: string;
  recipientEmail?: string;
  path?: string;
  url?: string;
  messageUri?: string;
}

class PurviewThreatAssessmentAddCommand extends GraphCommand {
  private readonly allowedTypes = ['file', 'url'];
  private readonly allowedExpectedAssessments = ['block', 'unblock'];
  private readonly allowedCategories = ['spam', 'phishing', 'malware'];

  public get name(): string {
    return commands.THREATASSESSMENT_ADD;
  }

  public get description(): string {
    return 'Create a threat assessment';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        path: typeof args.options.path !== 'undefined',
        url: typeof args.options.url !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type <type>',
        autocomplete: this.allowedTypes
      },
      {
        option: '-e, --expectedAssessment <expectedAssessment>',
        autocomplete: this.allowedExpectedAssessments
      },
      {
        option: '-c, --category <category>',
        autocomplete: this.allowedCategories
      },
      {
        option: '-p, --path [path]'
      },
      {
        option: '-u, --url [url]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!this.allowedTypes.some(type => type === args.options.type)) {
          return `${args.options.type} is not an allowed type. Allowed types are ${this.allowedTypes.join('|')}`;
        }

        if (!this.allowedExpectedAssessments.some(expectedAssessment => expectedAssessment === args.options.expectedAssessment)) {
          return `${args.options.expectedAssessment} is not an allowed expected assessment. Allowed expected assessments are ${this.allowedExpectedAssessments.join('|')}`;
        }

        if (!this.allowedCategories.some(category => category === args.options.category)) {
          return `${args.options.category} is not an allowed category. Allowed categories are ${this.allowedCategories.join('|')}`;
        }

        if (args.options.path && !fs.existsSync(args.options.path)) {
          return `File '${args.options.path}' not found. Please provide a valid path to the file.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['path'],
        runsWhen: (args) => {
          return args.options.type === 'file';
        }
      },
      {
        options: ['url'],
        runsWhen: (args) => {
          return args.options.type === 'url';
        }
      }
    );
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