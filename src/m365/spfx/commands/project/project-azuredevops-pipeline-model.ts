export interface AzureDevOpsPipeline {
  name?: string;
  trigger: {
    branches: {
      include: string[];
    };
    workflow_dispatch?: any;
  };
  pool: {
    vmImage: string;
  };
  variables: {
    name: string;
    value: string;
  }[];
  stages: {
    stage: string;
    jobs: {
      job: string;
      steps: AzureDevOpsPipelineStep[];
    }[];
  }[];
}

export interface AzureDevOpsPipelineStep {
  task?: string;
  script?: string;
  displayName?: string;
  inputs?: {
    versionSpec?: string;
    command?: string;
    gulpFile?: string;
    targets?: string;
    arguments?: string;
    verbose?: boolean;
    customCommand?: string;
  }
}