import { Logger } from '../../../cli';
import commands from '../commands';
import auth, { AuthType } from '../../../Auth';
import * as os from 'os';
import Command from '../../../Command';
import Utils from '../../../Utils';
const packageJSON = require('../../../../package.json');

interface CliDiagnosticInfo {
  OS: {
    Platform: string;
    Version: string;
    Release: string;
  };
  AuthMode: string;
  CliAadAppId: string;
  CliAadAppTenant: string;
  CliEnvironment: string;
  NodeVersion: string;
  CliVersion: string;
  Roles:string[];
  Scopes:string[];
  Shell: string;
}

class CliDoctorCommand extends Command {
  public get name(): string {
    return commands.DOCTOR;
  }

  public get description(): string {
    return 'Retrieves diagnostic information about the current environment';
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {

    const roles:string[] = [];
    const scopes:string[] = [];
    
    Object.keys(auth.service.accessTokens).forEach((tokenKey) => {
      const accessToken:string = auth.service.accessTokens[tokenKey].accessToken;

      Utils.getRolesFromAccessToken(accessToken).forEach(role => roles.push(role));
      Utils.getScopesFromAccessToken(accessToken).forEach(scope => scopes.push(scope));
    });

    const diagnosticInfo: CliDiagnosticInfo = {
      Shell: process.env.SHELL || "",
      OS: {
        Platform: os.platform(),
        Version: os.version(),
        Release: os.release()
      },
      CliVersion: packageJSON.version,
      NodeVersion: process.version,
      CliAadAppId: auth.service.appId,
      CliAadAppTenant: Utils.isValidGuid(auth.service.tenant) ? "single" : auth.service.tenant,
      AuthMode: AuthType[auth.service.authType],
      CliEnvironment: process.env.CLIMICROSOFT365_ENV ? process.env.CLIMICROSOFT365_ENV : "",
      Roles:roles,
      Scopes:scopes
    };

    logger.log(diagnosticInfo);
    cb();
  }
}

module.exports = new CliDoctorCommand();