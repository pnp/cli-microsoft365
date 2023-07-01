import { Hash } from '../../utils/types.js';

export interface M365RcJson {
  apps?: M365RcJsonApp[];
  context?: Hash;
}

export interface M365RcJsonApp {
  appId: string;
  name: string;
}