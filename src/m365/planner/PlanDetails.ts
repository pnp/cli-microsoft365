export interface PlanDetails {
  id: string;
  "@odata.etag"?: string;
  "@odata.context"?: string;
  sharewith:[string];
  categoryDescriptions: {
    category1?: string;
    category2?: string;
    category3?: string;
    category4?: string;
    category5?: string;
    category6?: string;
  }
}
