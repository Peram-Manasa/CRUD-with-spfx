

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { escape } from '@microsoft/sp-lodash-subset';

import styles from './components/Sharepointframe.module.scss';
import * as strings from 'SharepointframeWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IGetListItemFromSharePointListWebPartProps {
  description: string;
}
export interface ISPLists 
{
  value: ISPList[];
}
export interface ISPList 
{
  Title: string;
  Body: string;
  Images : string
}
export default class GetListItemFromSharePointListWebPart extends BaseClientSideWebPart <IGetListItemFromSharePointListWebPartProps> {

  private _getListData(): Promise<ISPLists>
  {
   return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Announcements')/Items?$select=Title,Body,Images",
       SPHttpClient.configurations.v1
   )
   .then((response: SPHttpClientResponse) => 
       {
       return response.json();
        console.log(response.json())
       });
   }
   private _renderListAsync(): void
   {
    if (Environment.type === EnvironmentType.SharePoint || 
             Environment.type === EnvironmentType.ClassicSharePoint) {
     this._getListData()
       .then((response) => {
         this._renderList(response.value);
         console.log(response.value);
       }).catch((err)=>{console.log(err)})
}
 }
 private _renderList(items: ISPList[]): void 
 {
  
let  html: string = '<table border=2 width=100% style="font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;">';
  

  html += '<b><th style="background-color: #af534c;" >Title</th> <th style="background-color: #af534c;">Body </th><th style="background-color: #af534c;">Images </th></b>';
   console.log(items)
  items.forEach((item: ISPList) => {
    // const imgurl = item.Images.Url
    html += `
    <tr>             
        <td>${item.Title}</td>
        <td>${item.Body}</td>
        <td>${item.Images}</td>
        </tr>
        `;
  });
  
  html += "</table>";
  
  const listContainer: Element = this.domElement.querySelector('#BindspListItems');
  listContainer.innerHTML = html;
  
}


  public render(): void {
    this.domElement.innerHTML = `
      <div class={styles.sharepointframe}>
    <div class={ styles.container }>
      <div class={ styles.row }>
        <div class="${ styles.column }">
          <span class="${ styles.title }">Today's Announcement </span>
          </div>
          <br/>
          <br/>
          <br/>

          <div id="BindspListItems" />
          </div>
          </div>
          </div>`;
          
          this._renderListAsync();
    
  }

  protected get dataVersion(): Version {
  return Version.parse('1.0');
}
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              })
            ]
          }
        ]
      }
    ]
  };
}
}