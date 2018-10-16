import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITrainingModuleProps {
 description: string;
 siteUrl: string;
 listname:string;
 context: WebPartContext;
}
