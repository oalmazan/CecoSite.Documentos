import * as React from 'react';
import styles from './CecobanDocumentLibrary.module.scss';
import { ICecobanDocumentLibraryProps } from './ICecobanDocumentLibraryProps';
import { IUserProfile } from './IUserProfile';
import { IUserProfileViewerState } from './IUserProfileViewerState';
import { IUserProfileService } from '../services/IUserProfileService';
import { ServiceScope } from '@microsoft/sp-core-library';
import { UserProfileService } from '../services/UserProfileService';

export class UserProfile implements IUserProfile {
  public FirstName: string;
  public LastName: string;    
  public Email: string;
  public Title: string;
  public WorkPhone: string;
  public DisplayName: string;
  public Department: string;
  public PictureURL: string;    
  public UserProfileProperties: Array<any>;
  public FilterProperty: string;    
}


export default class CecobanDocumentLibrary extends React.Component<ICecobanDocumentLibraryProps, IUserProfileViewerState> {
  private dataCenterServiceInstance: IUserProfileService;

  constructor(props: ICecobanDocumentLibraryProps, state: IUserProfileViewerState) {  
    super(props); 

    let userProfile: IUserProfile = new UserProfile();
    userProfile.FirstName = "";
    userProfile.LastName = "";
    userProfile.Email = "";
    userProfile.Title = "";
    userProfile.WorkPhone = "";
    userProfile.DisplayName = "";
    userProfile.Department = "";
    userProfile.PictureURL = "";
    userProfile.UserProfileProperties = [];
    userProfile.FilterProperty = "";
    
    this.state = {  
      userProfileItems: userProfile
    };     
  }

  public componentWillMount(): void {
    let serviceScope: ServiceScope = this.props.serviceScope;  
    this.dataCenterServiceInstance = serviceScope.consume(UserProfileService.serviceKey);

    this.dataCenterServiceInstance.getUserProfileProperties().then((userProfileItems: IUserProfile) => {  
      for (let i: number = 0; i < userProfileItems.UserProfileProperties.length; i++) {
        if (userProfileItems.UserProfileProperties[i].Key == "FirstName") {
          userProfileItems.FirstName = userProfileItems.UserProfileProperties[i].Value;
        }
        if (userProfileItems.UserProfileProperties[i].Key == "LastName") {
          userProfileItems.LastName = userProfileItems.UserProfileProperties[i].Value;
        }
        if (userProfileItems.UserProfileProperties[i].Key == "WorkPhone") {
          userProfileItems.WorkPhone = userProfileItems.UserProfileProperties[i].Value;
        }
        if (userProfileItems.UserProfileProperties[i].Key == "Department") {
          userProfileItems.Department = userProfileItems.UserProfileProperties[i].Value;
        }
        if (userProfileItems.UserProfileProperties[i].Key == "PictureURL") {
          userProfileItems.PictureURL = userProfileItems.UserProfileProperties[i].Value;
        }
        if(this.props.filterProperty == userProfileItems.UserProfileProperties[i].Key ){
          userProfileItems.FilterProperty = userProfileItems.UserProfileProperties[i].Value;
        }
      }
      this.setState({ userProfileItems: userProfileItems });  
    }); 
  }

  public render(): React.ReactElement<ICecobanDocumentLibraryProps> {
    return (
      <div className={ styles.cecobanDocumentLibrary }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Fetch User Profile Properties</p>
              
              <img src={this.state.userProfileItems.PictureURL}></img>
              
              <p> 
                Name: {this.state.userProfileItems.LastName}, {this.state.userProfileItems.FirstName}
              </p>

              <p>
                WorkPhone: {this.state.userProfileItems.WorkPhone}
              </p>
              
              <p>
                Department: {this.state.userProfileItems.Department}
              </p>              
            </div>
          </div>
        </div>
      </div>
    );
  }
}
