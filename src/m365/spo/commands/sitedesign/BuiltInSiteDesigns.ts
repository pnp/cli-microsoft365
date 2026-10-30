// GUID-to-name mapping of the built-in Microsoft site designs available in the SharePoint
// site template store (store 1). Ported from PnP PowerShell's BuiltInSiteTemplateSettings so
// the same friendly names can be used across both tools.
// See https://learn.microsoft.com/powershell/module/sharepoint-online/set-spobuiltinsitetemplatesettings?view=sharepoint-ps#description
export const builtInSiteDesigns: { id: string; template: string }[] = [
  { id: '00000000-0000-0000-0000-000000000000', template: 'All' },
  { id: '9522236e-6802-4972-a10d-e98dc74b3344', template: 'EventPlanning' },
  { id: 'f0a3abf4-afe8-4409-b7f3-484113dee93e', template: 'ProjectManagement' },
  { id: '695e52c9-8af7-4bd3-b7a5-46aca95e1c7e', template: 'TrainingAndCourses' },
  { id: '64aaa31e-7a1e-4337-b646-0b700aa9a52c', template: 'TrainingAndDevelopmentTeam' },
  { id: 'e4ec393e-da09-4816-b6b2-195393656edd', template: 'RetailManagement' },
  { id: 'af9037eb-09ef-4217-80fe-465d37511b33', template: 'EmployeeOnboardingTeam' },
  { id: '33537eba-a7d6-4d76-96cc-ee1930bd3907', template: 'SetUpYourHomePage' },
  { id: 'fb513aef-c06f-4dc3-b08c-963a2d2360c1', template: 'CrisisCommunicationTeam' },
  { id: '71308406-f31d-445f-85c7-b31942d1508c', template: 'ITHelpDesk' },
  { id: '2a7dd756-75f6-4f0f-a06a-a672939ea2a3', template: 'ContractsManagement' },
  { id: '403ffe4e-12d4-41a2-8153-208069eaf2b8', template: 'AccountsPayable' },
  { id: 'c8b3137a-ca4c-48a9-b356-a8e7987dd693', template: 'StandardTeam' },
  { id: '951190b8-8541-4f8c-8e8a-10a17c466c94', template: 'CrisisManagement' },
  { id: '73495f08-0140-499b-8927-dd26a546f26a', template: 'Department' },
  { id: 'cd4c26b2-b231-419a-8bb4-9b1d9b83aef6', template: 'LeadershipConnection' },
  { id: 'b8ef3134-92a2-4c9d-bca6-c2f14e79fe98', template: 'LearningCentral' },
  { id: '2a23fa44-52b0-4814-baba-06fef1ab931e', template: 'NewEmployeeOnboarding' },
  { id: '6142d2a0-63a5-4ba0-aede-d9fefca2c767', template: 'Showcase' },
  { id: '811ecf9a-b33f-44e6-81bd-da77729906dc', template: 'StoreCollaboration' },
  { id: '34a39504-194c-4605-87be-d48d00070c67', template: 'VolunteerCenter' },
  { id: 'f6cc5403-0d63-442e-96c0-285923709ffc', template: 'Blank' },
  { id: 'f2c6bb0c-9234-40c2-9ec3-ee86a70330fb', template: 'BrandCentral' },
  { id: '96c933ac-3698-44c7-9f4a-5fd17d71af9e', template: 'StandardCommunication' },
  { id: '3d5ef50b-88a0-42a7-9fb2-8036009f6f42', template: 'Event' },
  { id: 'c298ddc9-628d-48bf-b1e5-5939a1962fb1', template: 'HumanResources' },
  { id: '30eebaf6-48ea-4af9-a564-a5c50297c826', template: 'OrganizationHome' },
  { id: '94e24f52-dfaf-40e4-b629-df2c85570adc', template: 'CopilotCampaign' },
  { id: 'da99c5d9-baad-4e81-81f6-03a061972d49', template: 'VivaCampaign' }
];

export function getBuiltInSiteDesignId(template: string): string | undefined {
  return builtInSiteDesigns.find(d => d.template === template)?.id;
}

export function getBuiltInSiteDesignTemplateName(id: string): string | undefined {
  return builtInSiteDesigns.find(d => d.id.toLowerCase() === id.toLowerCase())?.template;
}

// 'All' (GUID 00000000-0000-0000-0000-000000000000) is a placeholder used by
// Get/Set-SPOBuiltInSiteTemplateSettings to refer to every built-in template at once. It doesn't
// identify an actual applicable site design, so it's excluded from the list of applyable templates.
export function getBuiltInSiteDesignTemplateNames(): string[] {
  return builtInSiteDesigns.filter(d => d.template !== 'All').map(d => d.template);
}
