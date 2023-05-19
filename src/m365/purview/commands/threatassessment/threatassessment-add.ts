import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import auth from '../../../../Auth';
import * as fs from 'fs';

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
  private static readonly allowedTypes = ['mail', 'file', 'emailFile', 'url'];
  private static readonly allowedExpectedAssessments = ['block', 'unblock'];
  private static readonly allowedCategories = ['spam', 'phishing', 'malware'];

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
        recipientEmail: typeof args.options.recipientEmail !== 'undefined',
        path: typeof args.options.path !== 'undefined',
        url: typeof args.options.url !== 'undefined',
        messageUri: typeof args.options.messageUri !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type <type>',
        autocomplete: PurviewThreatAssessmentAddCommand.allowedTypes
      },
      {
        option: '-e, --expectedAssessment <expectedAssessment>',
        autocomplete: PurviewThreatAssessmentAddCommand.allowedExpectedAssessments
      },
      {
        option: '-c, --category <category>',
        autocomplete: PurviewThreatAssessmentAddCommand.allowedCategories
      },
      {
        option: '-r, --recipientEmail [recipientEmail]'
      },
      {
        option: '-p, --path [path]'
      },
      {
        option: '-u, --url [url]'
      },
      {
        option: '-m, --messageUri [messageUri]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!PurviewThreatAssessmentAddCommand.allowedTypes.some(type => type === args.options.type)) {
          return `${args.options.type} is not an allowed type. Allowed types are ${PurviewThreatAssessmentAddCommand.allowedTypes.join('|')}`;
        }

        if (!PurviewThreatAssessmentAddCommand.allowedExpectedAssessments.some(expectedAssessment => expectedAssessment === args.options.expectedAssessment)) {
          return `${args.options.expectedAssessment} is not an allowed expected assessment. Allowed expected assessments are ${PurviewThreatAssessmentAddCommand.allowedExpectedAssessments.join('|')}`;
        }

        if (!PurviewThreatAssessmentAddCommand.allowedCategories.some(category => category === args.options.category)) {
          return `${args.options.category} is not an allowed category. Allowed categories are ${PurviewThreatAssessmentAddCommand.allowedCategories.join('|')}`;
        }

        if (args.options.path && !fs.existsSync(args.options.path)) {

          return `File at path ${args.options.path} does not exist. Please specify a path to an existing file`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['recipientEmail'],
        runsWhen: (args) => {
          return args.options.type === 'mail' || args.options.type === 'emailFile';
        }
      },
      {
        options: ['path'],
        runsWhen: (args) => {
          return args.options.type === 'file' || args.options.type === 'emailFile';
        }
      },
      {
        options: ['url'],
        runsWhen: (args) => {
          return args.options.type === 'url';
        }
      },
      {
        options: ['messageUri'],
        runsWhen: (args) => {
          return args.options.type === 'mail';
        }
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
        throw 'This command currently does not support app only permissions.';
      }

      if (this.verbose) {
        logger.logToStderr(`Adding threat assessment of type ${args.options.type} with expected assessment ${args.options.expectedAssessment} and category ${args.options.category}`);
      }

      const requestBody: any = {
        expectedAssessment: args.options.expectedAssessment,
        category: args.options.category,
        recipientEmail: args.options.recipientEmail,
        url: args.options.url,
        messageUri: args.options.messageUri,
        contentData: this.encodeFileFromPath(args.options.path)
      };

      switch (args.options.type) {
        case 'mail':
          requestBody['@odata.type'] = '#microsoft.graph.mailAssessmentRequest';
          break;
        case 'emailFile':
          requestBody['@odata.type'] = '#microsoft.graph.emailFileAssessmentRequest';
          break;
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
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }

  private encodeFileFromPath(path: string | undefined): string | undefined {
    if (!path) {
      return undefined;
    }
    return fs.readFileSync(path, 'base64');
  }
}

module.exports = new PurviewThreatAssessmentAddCommand();