export interface Plan {
  "@odata.etag": string;
  createdDateTime: string;
  id: string;
  owner: string;
  title: string;
  createdBy: {
    application: {
      displayName: string | undefined;
      id: string;
    },
    user: {
      displayName: string | undefined;
      id: string;
    }
  },
}