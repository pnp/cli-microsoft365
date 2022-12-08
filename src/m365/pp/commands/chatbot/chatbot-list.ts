import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  asAdmin?: boolean;
}

class PpChatbotListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.CHATBOT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Platform chatbots in the specified Power Platform environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'botid', 'publishedOn', 'createdOn', 'botModifiedOn'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of chatbots for environment '${args.options.environment}'.`);
    }

    const fetchXml: string = `
      <fetch mapping='logical' version='1.0' >
        <entity name='bot'>
          <attribute name='accesscontrolpolicy' alias='accessControlPolicy' />,
          <attribute name='applicationmanifestinformation' alias='applicationManifestInformation' />,
          <attribute name='authenticationmode' alias='authenticationMode' />,
          <attribute name='authenticationtrigger' alias='authenticationTrigger' />,
          <attribute name='authorizedsecuritygroupids' alias='authorizedSecurityGroupIds' />,
          <attribute name='componentidunique' alias='componentIdUnique' />,
          <attribute name='componentstate' alias='componentState' />,
          <attribute name='configuration' alias='configuration' />,
          <attribute name='createdon' alias='createdOn' />,
          <attribute name='importsequencenumber' alias='importSequenceNumber' />,
          <attribute name='ismanaged' alias='isManaged' />,
          <attribute name='language' alias='language' />,
          <attribute name='modifiedon' alias='botModifiedOn' />,
          <attribute name='overriddencreatedon' alias='overriddenCreatedOn' />,
          <attribute name='overwritetime' alias='overwriteTime' />,
          <attribute name='iconbase64' alias='iconBase64' />,
          <attribute name='publishedon' alias='publishedOn' />,
          <attribute name='schemaname' alias='schemaName' />,
          <attribute name='solutionid' alias='solutionId' />,
          <attribute name='statecode' alias='stateCode' />,
          <attribute name='statuscode' alias='statusCode' />,
          <attribute name='timezoneruleversionnumber' alias='timezoneRuleVersionNumber' />,
          <attribute name='utcconversiontimezonecode' alias='utcConversionTimezoneCode' />,
          <attribute name='versionnumber' alias='versionNumber' />,
          <attribute name='name' alias='displayName' />,
          <attribute name='botid' alias='cdsBotId' />,
          <attribute name='ownerid' alias='ownerId' />,
          <attribute name='synchronizationstatus' alias='synchronizationStatus' />
          <link-entity name='systemuser' to='ownerid' from='systemuserid' link-type='inner' >
            <attribute name='fullname' alias='owner' />
          </link-entity>
          <link-entity name='systemuser' to='modifiedby' from='systemuserid' link-type='inner' >
            <attribute name='fullname' alias='botModifiedBy' />
          </link-entity>
        </entity>
      </fetch>
    `;

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const items = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.1/bots?fetchXml=${fetchXml}`);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpChatbotListCommand();