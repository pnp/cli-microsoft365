export interface gitHubWorkflow {
  name: string,
  on: {
    push: {
      branches: string[];
    },
    workflow_dispatch?: any
  },
  jobs: {
    "build-and-deploy": {
      "runs-on": string,
      steps: gitHubWorkflowStep[]
    }
  }
}

export interface gitHubWorkflowStep {
  name?: string,
  run?: string,
  uses?: string,
  with?: {
    "node-version"?: string,
    CERTIFICATE_ENCODED?: string,
    CERTIFICATE_PASSWORD?: string,
    ADMIN_USERNAME?: string,
    ADMIN_PASSWORD?: string,
    APP_ID?: string,
    APP_FILE_PATH?: string,
    SKIP_FEATURE_DEPLOYMENT?: boolean,
    OVERWRITE?: boolean,
    SCOPE?: string,
    SITE_COLLECTION_URL?: string
  }
}