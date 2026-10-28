import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISalesProps {
  description: string;
  userDisplayName: string;
  spHttpClient:any;
  spContext:any;
  context:WebPartContext;
}

