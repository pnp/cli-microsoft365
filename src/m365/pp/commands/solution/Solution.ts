export interface Solution {
  solutionid: string;
  publisherid: Publisher;
  version: string;
  uniquename: string;
  installedon: string;
  solutionpackageversion: string;
  friendlyname: string;
  versionnumber: number;
}

export interface Publisher {
  friendlyname: string;
  publisherid: string;
}