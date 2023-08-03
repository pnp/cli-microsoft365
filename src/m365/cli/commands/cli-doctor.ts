import os from 'os';
import auth, { AuthType } from '../../../Auth.js';
import { Cli } from '../../../cli/Cli.js';
import { Logger } from '../../../cli/Logger.js';
import Command from '../../../Command.js';
import { app } from '../../../utils/app.js';
import { validation } from '../../../utils/validation.js';
import commands from '../commands.js';

interface CliDiagnosticInfo {
  os: {
    platform: string;
    version: string;
    release: string;
  };
  authMode: string;
  cliAadAppId: string;
  cliAadAppTenant: string;
  cliEnvironment: string;
  nodeVersion: string;
  cliVersion: string;
  cliConfig: any;
  roles: string[];
  scopes: string[];
}

class CliDoctorCommand extends Command {
  public get name(): string {
    return commands.DOCTOR;
  }

  public get description(): string {
    return 'Retrieves diagnostic information about the current environment';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const roles: string[] = [];
    const scopes: string[] = [];

    Object.keys(auth.service.accessTokens).forEach(resource => {
      const accessToken: string = auth.service.accessTokens[resource].accessToken;

      this.getRolesFromAccessToken(accessToken).forEach(role => roles.push(role));
      this.getScopesFromAccessToken(accessToken).forEach(scope => scopes.push(scope));
    });

    const diagnosticInfo: CliDiagnosticInfo = {
      os: {
        platform: os.platform(),
        version: os.version(),
        release: os.release()
      },
      cliVersion: app.packageJson().version,
      nodeVersion: process.version,
      cliAadAppId: auth.service.appId,
      cliAadAppTenant: validation.isValidGuid(auth.service.tenant) ? 'single' : auth.service.tenant,
      authMode: AuthType[auth.service.authType],
      cliEnvironment: process.env.CLIMICROSOFT365_ENV ? process.env.CLIMICROSOFT365_ENV : '',
      cliConfig: Cli.getInstance().config.all,
      roles: roles,
      scopes: scopes
    };

    await logger.log(diagnosticInfo);
  }

  private getRolesFromAccessToken(accessToken: string): string[] {
    let roles: string[] = [];

    if (!accessToken || accessToken.length === 0) {
      return roles;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return roles;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    const token: { roles: string[] } = JSON.parse(tokenString);
    if (token.roles !== undefined) {
      roles = token.roles;
    }

    return roles;
  }

  private getScopesFromAccessToken(accessToken: string): string[] {
    let scopes: string[] = [];

    if (!accessToken || accessToken.length === 0) {
      return scopes;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return scopes;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();

    const token: { scp: string } = JSON.parse(tokenString);
    if (token.scp?.length > 0) {
      scopes = token.scp.split(' ');
    }

    return scopes;
  }
}

export default new CliDoctorCommand();