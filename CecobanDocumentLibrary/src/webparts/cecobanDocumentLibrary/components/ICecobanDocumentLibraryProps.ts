import { ServiceScope } from '@microsoft/sp-core-library';

export interface ICecobanDocumentLibraryProps {
  description: string;
  userName: string;
  serviceScope: ServiceScope;
  filterProperty: string;
}
