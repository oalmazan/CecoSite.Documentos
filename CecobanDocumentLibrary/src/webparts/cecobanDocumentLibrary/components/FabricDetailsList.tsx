import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Environment, EnvironmentType, ServiceScope } from "@microsoft/sp-core-library";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  SelectionMode
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import { IFabricDetailsListProps } from "./IFabricDetailsListProps";
import pnp from "sp-pnp-js";
import { IUserProfile } from "./IUserProfile";
import { UserProfileService } from "../services/ListService";
import { IUserProfileService } from "../services/IUserProfileService";
import { ActionButton, BaseButton, Button } from "office-ui-fabric-react";
import '../loc/estilos.css'; // Import regular stylesheet
import * as moment from 'moment';

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: "16px"
  },
  fileIconCell: {
    textAlign: "center",
    selectors: {
      "&:before": {
        content: ".",
        display: "inline-block",
        verticalAlign: "middle",
        height: "100%",
        width: "0px",
        visibility: "hidden"
      }
    }
  },
  fileIconImg: {
    verticalAlign: "middle",
    maxHeight: "16px",
    maxWidth: "16px"
  },
  controlWrapper: {
    flexWrap: "nowrap",
    flexShrink: "0",
    display: "inherit"
  },
  divHeaderRow: {
    position: "relative",
    display: "flex",
    flexWrap: "nowrap",
    flexGrow: "1",
    alignItems: "stretch",
  },
  exampleToggle: {
    display: "inline-block",
    marginBottom: "10px",
    marginRight: "30px"
  },
  selectionDetails: {
    marginBottom: "20px"
  }
});
const controlStyles = {
  root: {
    margin: "0 30px 20px 0",
    maxWidth: "300px"
  }
};

export interface IDetailsListDocumentsExampleState {
  userProfileItems: IUserProfile;
  columns: IColumn[];
  items: IDocument[];
  allItems: IDocument[];
  selectionDetails: string;
  subElement:string;
  showAreaItem: boolean;
}

export interface IDocument {
  id: number;
  num: number;
  name: string;
  value: string;
  iconName: string;
  fileType: string;
  modifiedBy: string;
  dateModified: string;
  dateModifiedValue: number;
  dateRevision: string;
  dateRevisionValue: number;
  version: number;
  fileSize: string;
  fileSizeRaw: number;
  itemChildCount: number;
  folderChildCount: number;
  referenceThis: any;
  serverRelativeUrl: string;
}

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


export default class FabricDetailsList extends React.Component<IFabricDetailsListProps, IDetailsListDocumentsExampleState> {
  private _selection: Selection;
  private _allItems: IDocument[];
  private dataCenterServiceInstance: IUserProfileService;

  constructor(props: IFabricDetailsListProps, state: IDetailsListDocumentsExampleState) {
    super(props);

    this._allItems = [];

    const columns: IColumn[] = [
      {
        key: "itemChildCount",
        name: "N°",
        fieldName: "num",
        minWidth: 30,
        maxWidth: 30,
        isResizable: true,
        data: "number",
        isPadded: true

      },
      {
        key: "name",
        name: "Nombre del Documento",
        fieldName: "name",
        minWidth: 380,
        maxWidth: 480,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: "A a Z",
        sortDescendingAriaLabel: "Z a A",
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      },
      {
        key: "dateRevisionValue",
        name: "​​Fecha de última revisión",
        fieldName: "dateRevisionValue",
        minWidth: 180,
        maxWidth: 180,
        isResizable: true,
        data: "number",
        onRender: (item: IDocument) => {
          return <span>{item.dateRevision}</span>;
        },
        isPadded: true

      },
      {
        key: "dateModifiedValue",
        name: "Fecha de última actualización​",
        fieldName: "dateModifiedValue",
        minWidth: 180,
        maxWidth: 180,
        isResizable: true,
        data: "number",
        onRender: (item: IDocument) => {
          return <span>{item.dateModified}</span>;
        },
        isPadded: true
      },
      {
        key: "version",
        name: "Versión",
        fieldName: "version",
        minWidth: 30,
        maxWidth: 30,
        isResizable: true,
        data: "string",
        onRender: (item: IDocument) => {
          return <span>{item.version}</span>;
        },
        isPadded: true
      }
      /*,
      {
        key: "itemChildCount",
        name: "Elementos",
        fieldName: "itemChildCount",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "number",
        isPadded: true
      },
      {
        key: "folderChildCount",
        name: "Carpetas",
        fieldName: "folderChildCount",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "number",
        isPadded: true
      }*/
      /*,
      {
        key: "column4",
        name: "Modificado Por",
        fieldName: "modifiedBy",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: "string",
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.modifiedBy}</span>;
        },
        isPadded: true
      },
      {
        key: "column5",
        name: "Tamaño",
        fieldName: "fileSizeRaw",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: "number",
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.fileSize}</span>;
        }
      }
      */
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });

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
      userProfileItems: userProfile,
      items: this._allItems,
      allItems : this._allItems,
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      subElement:"",
      showAreaItem : true
    };
  }

  public render() {
    const { columns, items, selectionDetails } = this.state;

    return (
      <Fabric>
        <style>
        </style>
        <div className={classNames.divHeaderRow}>
          <div className={classNames.controlWrapper}>
            <ActionButton 
            label="Ir atras"
            text="Ir atras"
            disabled={this.state && this.state.subElement?false:true}
            onClick={this._onRestore}
            styles={controlStyles}
            className="actionButtonCls"
            />
          </div>
          <div className={classNames.controlWrapper}>
            <TextField
              label="Filtro:"
              onChange={this._onChangeText}
              styles={controlStyles}
              className="textFieldCls"
            />
          </div>
        </div>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={items}
            columns={columns}
            setKey="set"
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={this._onItemInvoked}
            onActiveItemChanged={this._onItemInvoked}
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn=""
            ariaLabelForSelectAllCheckbox=""
            className="DetailsListCCBCls"
    
          />
        </MarqueeSelection>
      </Fabric>
    );
  }

  public componentWillMount(): void {
    let serviceScope: ServiceScope = this.props.ServiceScope;  
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
        if(this.props.FilterProperty == userProfileItems.UserProfileProperties[i].Key ){
          userProfileItems.FilterProperty = userProfileItems.UserProfileProperties[i].Value;
        }
      }
      pnp.sp.web.lists.getById(this.props.spcontect.pageContext.list.id.toString())  
      .items.getById(this.props.spcontect.pageContext.listItem.id)  
        .select("*")  
          .get()  
            .then(d => {  
              const showAreaItem = this.CompareStringC(d.Title,userProfileItems.FilterProperty);
              const newState = {userProfileItems : userProfileItems, showAreaItem:showAreaItem}; 
              this.setState( newState );  
              this._allItems = _generateDocuments(this.props, "", this);        
            });
    }); 
  }
  private RemoveAccents(strAccentsP: string) {
    const strAccents = strAccentsP.split('');
    const strAccentsOut = [];
    const strAccentsLen = strAccents.length;
    const accents = 'ÀÁÂÃÄÅàáâãäåÒÓÔÕÕÖØòóôõöøÈÉÊËèéêëðÇçÐÌÍÎÏìíîïÙÚÛÜùúûüÑñŠšŸÿýŽž';
    const accentsOut = "AAAAAAaaaaaaOOOOOOOooooooEEEEeeeeeCcDIIIIiiiiUUUUuuuuNnSsYyyZz";
    for (var y = 0; y < strAccentsLen; y++) {
      if (accents.indexOf(strAccents[y]) != -1) {
        strAccentsOut[y] = accentsOut.substr(accents.indexOf(strAccents[y]), 1);
      } else {
        strAccentsOut[y] = strAccents[y];
      }
    }
    return strAccentsOut.join('');
  }

  private CompareStringC(str1: string, str2: string){
   return str1 && str2 && this.RemoveAccents(str1.toUpperCase()) ==  this.RemoveAccents(str2.toUpperCase());
  }

  private _onChangeText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    const items  = this.state.allItems;
    this.setState({
      items: text
        ? items.filter(i => i.name.toLowerCase().indexOf(text) > -1)
        : items
    });
  }

 
  private _onRestore = (event: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button, MouseEvent>) : void => {
    _generateDocuments(this.props, "", this);
  }
  
  private _onItemInvoked(item: any): void {
    if(!item.fileType){
      _generateDocuments(item.referenceThis.props, item.name, item.referenceThis);
    }else{
      window.open(item.serverRelativeUrl);
    }
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return "Sin seleccion";
      case 1:
        return (
          "1 elemento seleccionado : " +
          (this._selection.getSelection()[0] as IDocument).name
        );
      default:
        return `${selectionCount} elementos seleccionados.`;
    }
  }

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(
      currCol => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(
      items,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }
}

function _copyAndSort<T>(
  items: T[],
  columnKey: string,
  isSortedDescending?: boolean
): T[] {
  const key = columnKey as keyof T;
  return items
    .slice(0)
    .sort((a: T, b: T) =>
      (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
    );
}

function _generateDocuments(props: IFabricDetailsListProps, subElement: String, obj: any) {
  let items: IDocument[] = [];
  if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
    const caml =  {
      ViewXml: "<View>"
                  +"<ViewFields><FieldRef Name='"+props.DocumentsFilter+"' /><FieldRef Name='Version' /><FieldRef Name='Responsable' /><FieldRef Name='FecUltRevision' /><FieldRef Name='FechaModificacion' /><FieldRef Name='ItemChildCount' /><FieldRef Name='FolderChildCount' /><FieldRef Name='Editor' /><FieldRef Name='ContentTypeId' /><FieldRef Name='Title' /><FieldRef Name='FileDirRef' /><FieldRef Name='FileRef' /></ViewFields>"
                  +"<Query>"
                    +"<Where>"
                      +"<Or>"
                        +"<Or>"
                          +"<Eq><FieldRef Name='"+props.DocumentsFilter+"'/><Value Type='Text'></Value></Eq>"
                          +"<Eq><FieldRef Name='"+props.DocumentsFilter+"'/><Value Type='Text'>"+obj.state.userProfileItems.FilterProperty+"</Value></Eq>"
                        +"</Or>"
                        +"<IsNull><FieldRef Name='"+props.DocumentsFilter+"'/></IsNull>"
                      +"</Or>"
                    +"</Where>"
                  +"</Query>"
                  +"<RowLimit>"+(props.LimitDocuments||500)+"</RowLimit>"
                +"</View>",    
      FolderServerRelativeUrl: (props.StartFilesPath||"/sites/CecoSiteDes2/Documentos compartidos/Directorio/03_Documentacion/01_AyF")+(subElement?("/"+subElement):""),
      ListItemCollectionPosition: null,
      DatesInUtc: false
    }; 
     pnp.sp.web.lists.getByTitle(props.TitleDocuments||"Documentos").getItemsByCAMLQuery(caml,"Folder","File","FolderChildCount","ItemChildCount","FileRef","FileDirRef").then(listItems => {
      if(listItems.error){
        alert(listItems.error.message);
      }else{
        const tempList = listItems.sort((a, b) => a.File && b.File && a.File.Name < b.File.Name ? -1 : 1);
        let tempId = 0;
        tempList.forEach(element => {
          let iconurl:string;
          let serverRelativeUrl: string; 
          let file = element.File;
          let isFolder = false;
          if(file){
            const ext = element.File.Name.split('.').pop().toLowerCase();
            serverRelativeUrl = element.File.ServerRelativeUrl;
            if(ext === 'pdf'){
              iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/48/pdf.svg';
            }else{
              if(ext==="png"||ext==="jpeg"||ext==="jpg"){
                iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types-fluent/20/photo.svg';
              } else{
                iconurl =`https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${element.File.Name.split('.').pop()}_16x1.svg`;
              }
            } 
            if(element.ServerRedirectedEmbedUrl){
              serverRelativeUrl = element.ServerRedirectedEmbedUrl;
            }  
          }else{
            isFolder = true;
            file = element.Folder;
            serverRelativeUrl = '';
            iconurl = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types-fluent/20/folder.svg";
          }
          if(file){
            const areaArray= file.Name.split('_');
            const areaText = areaArray[areaArray.length-1];
            const showAreaComplexItem = isFolder && file.Name.toLowerCase().indexOf("area") !=-1 && obj.CompareStringC(areaText, obj.state.userProfileItems.FilterProperty);
            const showAreaItem = isFolder && file.Name.toLowerCase() =="area" && obj.state.showAreaItem;
            const showOrgItem = isFolder && file.Name.toLowerCase()=="organizacionales";
            if((!isFolder) || showOrgItem || showAreaComplexItem || showAreaItem ){  
              tempId++;          
              items.push({
                id : element.id||element.ID,
                num: tempId,
                name: file.Name.replace('Area_',''),
                value: file.Name.replace('Area_',''),
                iconName: iconurl,
                fileType: file.Name.indexOf(".")==-1?"":file.Name.split('.')[1],
                modifiedBy: element.Editor?element.Editor.Title:"",
                dateModified: moment((element['FechaModificacion']?new Date(element['FechaModificacion']):new Date(element.Modified))).format('DD/MM/YYYY')  ,
                dateModifiedValue: (element['FechaModificacion']?new Date(element['FechaModificacion']):new Date(element.Modified)).valueOf(),
                dateRevision:   element['FecUltRevision']?moment(new Date(element['FecUltRevision'])).format('DD/MM/YYYY'):""  ,
                dateRevisionValue: element['FecUltRevision']?new Date(element['FecUltRevision']).valueOf():0,
                version: element['Version']?element['Version']:'1.0',
                fileSize: readableFileSize(file.Length),
                fileSizeRaw: file.Length,
                itemChildCount: element.ItemChildCount,
                folderChildCount: element.FolderChildCount,
                referenceThis: obj,
                serverRelativeUrl: serverRelativeUrl
              });  
            }
          }else{
            console.log('????', element);
          }
        });
        obj.setState({ items: items, allItems: items, subElement: subElement});
      }
    });          
  }
  else if (Environment.type === EnvironmentType.Local) {
    for (let i = 0; i < 500; i++) {
      const randomDate = _randomDate(new Date(2012, 0, 1), new Date());
      const randomFileSize = _randomFileSize();
      const randomFileType = _randomFileIcon();
      let fileName = _lorem(2);
      fileName =
        fileName.charAt(0).toUpperCase() +
        fileName.slice(1).concat(`.${randomFileType.docType}`);
      let userName = _lorem(2);
      userName = userName
        .split(" ")
        .map((name: string) => name.charAt(0).toUpperCase() + name.slice(1))
        .join(" ");
      items.push({
        id: 0,
        num: 0,
        name: fileName,
        value: fileName,
        iconName: randomFileType.url,
        fileType: randomFileType.docType,
        modifiedBy: userName,
        dateModified: randomDate.dateFormatted,
        dateModifiedValue: randomDate.value,
        dateRevision: randomDate.dateFormatted,
        dateRevisionValue: randomDate.value,
        version: randomDate.value,
        fileSize: randomFileSize.value,
        fileSizeRaw: randomFileSize.rawSize,
        itemChildCount: 0,
        folderChildCount: 0,
        referenceThis: obj,
        serverRelativeUrl:""
      });
    }
  }
  return items;
}

function readableFileSize(size) {
  if(!size){
    size = 0;
  }
  var units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
  var i = 0;
  while(size >= 1024) {
      size /= 1024;
      ++i;
  }
  return size.toFixed(1) + ' ' + units[i];
}

function _randomDate(
  start: Date,
  end: Date
): { value: number; dateFormatted: string } {
  const date: Date = new Date(
    start.getTime() + Math.random() * (end.getTime() - start.getTime())
  );
  return {
    value: date.valueOf(),
    dateFormatted: date.toLocaleDateString()
  };
}

const FILE_ICONS: { name: string }[] = [
  { name: "accdb" },
  { name: "csv" },
  { name: "docx" },
  { name: "dotx" },
  { name: "mpt" },
  { name: "odt" },
  { name: "one" },
  { name: "onepkg" },
  { name: "onetoc" },
  { name: "pptx" },
  { name: "pub" },
  { name: "vsdx" },
  { name: "xls" },
  { name: "xlsx" },
  { name: "xsn" }
];

function _randomFileIcon(): { docType: string; url: string } {
  const docType: string =
    FILE_ICONS[Math.floor(Math.random() * FILE_ICONS.length)].name;
  return {
    docType,
    url: `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${docType}_16x1.svg`
  };
}

function _randomFileSize(): { value: string; rawSize: number } {
  const fileSize: number = Math.floor(Math.random() * 100) + 30;
  return {
    value: `${fileSize} KB`,
    rawSize: fileSize
  };
}

const LOREM_IPSUM = (
  "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut " +
  "labore et dolore magna aliqua ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut " +
  "aliquip ex ea commodo consequat duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore " +
  "eu fugiat nulla pariatur excepteur sint occaecat cupidatat non proident sunt in culpa qui officia deserunt "
).split(" ");
let loremIndex = 0;
function _lorem(wordCount: number): string {
  const startIndex =
    loremIndex + wordCount > LOREM_IPSUM.length ? 0 : loremIndex;
  loremIndex = startIndex + wordCount;
  return LOREM_IPSUM.slice(startIndex, loremIndex).join(" ");
}
