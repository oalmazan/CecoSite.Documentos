import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";  
import { IUserProfile } from '../components/IUserProfile';
import { IUserProfileService } from './IUserProfileService'; 
import { BaseService } from "./BaseService";
 
export class UserProfileService extends BaseService<IUserProfile> implements IUserProfileService {
    public static readonly serviceKey: ServiceKey<IUserProfileService> = ServiceKey.create<IUserProfileService>('userProfle:data-service', UserProfileService);  
 
    constructor(serviceScope: ServiceScope) {  
      super(serviceScope);
    }

    public getUserProfileProperties(): Promise<IUserProfile> {
      return this.get(`${this._currentWebUrl}/_api/SP.UserProfiles.PeopleManager/getmyproperties`);
    }
}