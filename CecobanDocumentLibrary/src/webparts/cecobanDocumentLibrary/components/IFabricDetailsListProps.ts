import { ServiceScope } from "@microsoft/sp-core-library";

export interface IFabricDetailsListProps {
  spcontect?:any|null;
  StartFilesPath:string;
  TitleDocuments:string;
  LimitDocuments: number;
  DocumentsFilter: string;
  UserName: string;
  ServiceScope: ServiceScope;
  FilterProperty: string;
}
