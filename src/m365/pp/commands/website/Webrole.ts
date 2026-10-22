export interface Webrole {
  mspp_webroleid: string;
  mspp_name: string;
  mspp_description: string | null;
  mspp_key: string | null;
  mspp_authenticatedusersrole: boolean;
  mspp_anonymoususersrole: boolean;
  mspp_createdon: string;
  mspp_modifiedon: string;
  statecode: number;
  statuscode: number;
  _mspp_websiteid_value: string;
  _mspp_createdby_value: string;
  _mspp_modifiedby_value: string;
}