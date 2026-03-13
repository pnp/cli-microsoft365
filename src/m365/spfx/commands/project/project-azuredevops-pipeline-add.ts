import fs from 'fs';
import path from 'path';
import yaml from 'yaml';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { BaseProjectCommand } from './base-project-command.js';
import { validation } from '../../../../utils/validation.js';
import { pipeline } from './DeployWorkflow.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import { AzureDevOpsPipeline, AzureDevOpsPipelineStep } from './project-azuredevops-pipeline-model.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { versions } from '../SpfxCompatibilityMatrix.js';
import { spfx } from '../../../../utils/spfx.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
  branchName?: string;
  loginMethod?: string;
  scope?: string;
  skipFeatureDeployment?: boolean;
  siteUrl?: string;
}

class SpfxProjectAzureDevOpsPipelineAddCommand extends BaseProjectCommand {
  private static loginMethod: string[] = ['application', 'user'];
  private static scope: string[] = ['tenant', 'sitecollection'];
  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;

  public get name(): string {
    return commands.PROJECT_AZUREDEVOPS_PIPELINE_ADD;
  }

  public get description(): string {
    return 'Adds a Azure DevOps Pipeline for a SharePoint Framework project.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined',
        branchName: typeof args.options.branchName !== 'undefined',
        loginMethod: typeof args.options.loginMethod !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        skipFeatureDeployment: !!args.options.skipFeatureDeployment
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '-b, --branchName [branchName]'
      },
      {
        option: '-l, --loginMethod [loginMethod]',
        autocomplete: SpfxProjectAzureDevOpsPipelineAddCommand.loginMethod
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: SpfxProjectAzureDevOpsPipelineAddCommand.scope
      },
      {
        option: '-u, --siteUrl [siteUrl]'
      },
      {
        option: '--skipFeatureDeployment'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.scope && args.options.scope === 'sitecollection') {
          if (!args.options.siteUrl) {
            return `siteUrl option has to be defined when scope set to ${args.options.scope}`;
          }

          const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        if (args.options.loginMethod && SpfxProjectAzureDevOpsPipelineAddCommand.loginMethod.indexOf(args.options.loginMethod) < 0) {
          return `${args.options.loginMethod} is not a valid login method. Allowed values are ${SpfxProjectAzureDevOpsPipelineAddCommand.loginMethod.join(', ')}`;
        }

        if (args.options.scope && SpfxProjectAzureDevOpsPipelineAddCommand.scope.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Allowed values are ${SpfxProjectAzureDevOpsPipelineAddCommand.scope.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      throw new CommandError(`Couldn't find project root folder`, SpfxProjectAzureDevOpsPipelineAddCommand.ERROR_NO_PROJECT_ROOT_FOLDER);
    }

    const solutionPackageJsonFile: string = path.join(this.projectRootPath, 'package.json');
    const packageJson: string = fs.readFileSync(solutionPackageJsonFile, 'utf-8');
    const solutionName = JSON.parse(packageJson).name;

    if (this.debug) {
      await logger.logToStderr(`Adding Azure DevOps pipeline in the current SPFx project`);
    }

    try {
      this.updatePipeline(solutionName, pipeline, args.options);
      this.savePipeline(pipeline);
    }
    catch (error: any) {
      this.handleError(error);
    }
  }

  private savePipeline(pipeline: AzureDevOpsPipeline): void {
    const azureDevOpsPath: string = path.join(this.projectRootPath as string, '.azuredevops');
    fsUtil.ensureDirectory(azureDevOpsPath);

    const pipelinesPath: string = path.join(azureDevOpsPath, 'pipelines');
    fsUtil.ensureDirectory(pipelinesPath);

    const pipelineFile: string = path.join(pipelinesPath, 'deploy-spfx-solution.yml');
    fs.writeFileSync(path.resolve(pipelineFile), yaml.stringify(pipeline), 'utf-8');
  }

  private updatePipeline(solutionName: string, pipeline: AzureDevOpsPipeline, options: GlobalOptions): void {
    if (options.name) {
      pipeline.name = options.name;
    }
    else {
      delete pipeline.name;
    }

    if (options.branchName) {
      pipeline.trigger.branches.include[0] = options.branchName;
    }

    const version = this.getProjectVersion();

    if (!version) {
      throw 'Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on the version property in the .yo-rc.json file.';
    }

    const versionRequirements = versions[version];

    if (!versionRequirements) {
      throw `Could not find Node version for version '${version}' of SharePoint Framework.`;
    }

    const nodeVersion: string = spfx.getHighestNodeVersion(versionRequirements.node.range);

    this.assignPipelineVariables(pipeline, 'NodeVersion', nodeVersion);

    const script = this.getScriptAction(pipeline);
    if (script.script) {
      if (options.loginMethod === 'user') {
        script.script = script.script.replace(`{{login}}`, `m365 login --authType password --userName '$(UserName)' --password '$(Password)'`);
        pipeline.variables = pipeline.variables.filter(v =>
          v.name !== 'CertificateBase64Encoded' &&
          v.name !== 'CertificateSecureFileId' &&
          v.name !== 'CertificatePassword' &&
          v.name !== 'EntraAppId' &&
          v.name !== 'TenantId'
        );
      }
      else {
        script.script = script.script.replace(`{{login}}`, `m365 login --authType certificate --certificateBase64Encoded '$(CertificateBase64Encoded)' --password '$(CertificatePassword)' --appId '$(EntraAppId)' --tenant '$(TenantId)'`);
        pipeline.variables = pipeline.variables.filter(v =>
          v.name !== 'UserName' &&
          v.name !== 'Password'
        );
      }

      if (options.scope === 'sitecollection') {
        script.script = script.script.replace(`{{deploy}}`, `m365 spo app deploy --name '$(PackageName)' --appCatalogScope sitecollection --appCatalogUrl '$(SiteAppCatalogUrl)'`);
        script.script = script.script.replace(`{{addApp}}`, `m365 spo app add --filePath '$(Build.SourcesDirectory)/sharepoint/solution/$(PackageName)' --appCatalogScope sitecollection --appCatalogUrl '$(SiteAppCatalogUrl)' --overwrite`);
        this.assignPipelineVariables(pipeline, 'SiteAppCatalogUrl', options.siteUrl);
      }
      else {
        script.script = script.script.replace(`{{deploy}}`, `m365 spo app deploy --name '$(PackageName)' --appCatalogScope 'tenant'`);
        script.script = script.script.replace(`{{addApp}}`, `m365 spo app add --filePath '$(Build.SourcesDirectory)/sharepoint/solution/$(PackageName)' --overwrite`);
        pipeline.variables = pipeline.variables.filter(v => v.name !== 'SiteAppCatalogUrl');
      }

      if (solutionName) {
        this.assignPipelineVariables(pipeline, 'PackageName', `${solutionName}.sppkg`);
      }

      if (options.skipFeatureDeployment) {
        script.script = script.script.replace(`m365 spo app deploy `, `m365 spo app deploy --skipFeatureDeployment `);
      }
    }
  }

  private assignPipelineVariables(pipeline: AzureDevOpsPipeline, variableName: string, newVariableValue: string): void {
    const variable = pipeline.variables.find(v => v.name === variableName);
    if (variable) {
      variable.value = newVariableValue;
    }
  }

  private getScriptAction(pipeline: AzureDevOpsPipeline): AzureDevOpsPipelineStep {
    const steps = this.getPipelineSteps(pipeline);
    return steps.find(step => step.script)!;
  }

  private getPipelineSteps(pipeline: AzureDevOpsPipeline): AzureDevOpsPipelineStep[] {
    return pipeline.stages[0].jobs[0].steps;
  }
}

export default new SpfxProjectAzureDevOpsPipelineAddCommand();
