import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  asAdmin?: boolean;
}

class PpCopilotListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.COPILOT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Platform copilots in the specified Power Platform environment';
  }

  public alias(): string[] | undefined {
    return [commands.CHATBOT_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'botid', 'publishedOn', 'createdOn', 'botModifiedOn'];
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
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of copilots for environment '${args.options.environmentName}'.`);
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
          <attribute name='name' alias='name' />,
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
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const items = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.1/bots?fetchXml=${fetchXml}`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpCopilotListCommand();