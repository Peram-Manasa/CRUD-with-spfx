

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption ,

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
  webURL:string;

  
}
export interface ISPLists 
{
  value: ISPList[];
}
export interface ISPList 
{
  Title: string;
  Amount: number;
  Purpose:string;
  Category:string;
}
export default class GetListItemFromSharePointListWebPart extends BaseClientSideWebPart <IGetListItemFromSharePointListWebPartProps> {
// private _siteLists: string[];

public dropdownOptions: IPropertyPaneDropdownOption[] = []; 


  private _getListData(): Promise<ISPLists>
  {
   return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('ZEA Fund Details')/Items?$select=Title,Amount,Purpose,Category",
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
  let amountexpenses=0;
      let amountfunding=0;
      let balance=0;
     
     
      
          
        
      let  html: string = ' <div style="border:2px solid black;background-color:pink; width:500px;"><th  width=60%   style="font-family: "Trebuchet MS", Arial, Helvetica, sans-serif; justify-content:space-between;">';
  


   html += '<b style="display:flex; margin:20px;  justify-content: space-between; font-size:20px;"><p style="color:red;">Expenses</p><p style="color:green; ">Funding </p><p style="color: blue;">Balance </p></b>';
   console.log(items)
  items.forEach((item: ISPList) => {
    if(item.Category==="Expenses"){
      amountexpenses+=item.Amount;
    }
   else if (item.Category==="Funding"){
    amountfunding+=item.Amount
   }
  }); 
  if(amountexpenses>amountfunding)
  {
    balance=amountexpenses-amountfunding;
  }
  else  if(amountfunding>amountexpenses){
    balance=amountfunding-amountexpenses;
  }
  else if(amountfunding===amountexpenses)
  {
    balance=0;
  }
    // const imgurl = item.Images.Url
    html += `
        <br/> 
        <div  style="display:flex; justify-content:space-between; margin:20px; font-size:30px; margin-top:-30px; font-weight:700;">   
        <tr>
        <th>₹ ${amountexpenses}</th><div style="margin-top:-50px;"><svg width="40px" height="40px"  viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M12 21C10.22 21 8.47991 20.4722 6.99987 19.4832C5.51983 18.4943 4.36628 17.0887 3.68509 15.4442C3.0039 13.7996 2.82567 11.99 3.17294 10.2442C3.5202 8.49836 4.37737 6.89472 5.63604 5.63604C6.89472 4.37737 8.49836 3.5202 10.2442 3.17294C11.99 2.82567 13.7996 3.0039 15.4442 3.68509C17.0887 4.36628 18.4943 5.51983 19.4832 6.99987C20.4722 8.47991 21 10.22 21 12C21 14.387 20.0518 16.6761 18.364 18.364C16.6761 20.0518 14.387 21 12 21ZM12 4.5C10.5166 4.5 9.0666 4.93987 7.83323 5.76398C6.59986 6.58809 5.63856 7.75943 5.07091 9.12988C4.50325 10.5003 4.35473 12.0083 4.64411 13.4632C4.9335 14.918 5.64781 16.2544 6.6967 17.3033C7.7456 18.3522 9.08197 19.0665 10.5368 19.3559C11.9917 19.6453 13.4997 19.4968 14.8701 18.9291C16.2406 18.3614 17.4119 17.4001 18.236 16.1668C19.0601 14.9334 19.5 13.4834 19.5 12C19.5 10.0109 18.7098 8.10323 17.3033 6.6967C15.8968 5.29018 13.9891 4.5 12 4.5Z" fill="#000000"/>
        <path d="M16 12.75H8C7.80109 12.75 7.61032 12.671 7.46967 12.5303C7.32902 12.3897 7.25 12.1989 7.25 12C7.25 11.8011 7.32902 11.6103 7.46967 11.4697C7.61032 11.329 7.80109 11.25 8 11.25H16C16.1989 11.25 16.3897 11.329 16.5303 11.4697C16.671 11.6103 16.75 11.8011 16.75 12C16.75 12.1989 16.671 12.3897 16.5303 12.5303C16.3897 12.671 16.1989 12.75 16 12.75Z" fill="#000000"/>
        </svg></div>
        <th>₹ ${amountfunding}</th><div style="margin-top:-50px;"><svg width="40px" height="40px" viewBox="0 0 72 72" id="emoji" version="1.1" xmlns="http://www.w3.org/2000/svg">
        <g id="color">
          <circle cx="36" cy="36" r="26.68" fill="#FFFFFF" fill-rule="evenodd" paint-order="normal"/>
        </g>
        <g id="line">
          <circle cx="36" cy="36" r="26.68" fill="none" stroke="#000000" stroke-linecap="round" stroke-linejoin="round" stroke-width="4.74" paint-order="normal"/>
          <path fill="none" stroke="#000000" stroke-linecap="round" stroke-linejoin="round" stroke-width="8.031" d="m28.03 42.18h15.95" clip-rule="evenodd"/>
          <path fill="none" stroke="#000000" stroke-linecap="round" stroke-linejoin="round" stroke-width="8.031" d="m28.03 29.82h15.95" clip-rule="evenodd"/>
        </g>
      </svg></div>
        <th>₹ ${balance}</th>
        </tr>
        </div>
      
        
        `;
  
  
  html += "</th></div>";
  
  const listContainer: Element = this.domElement.querySelector('#BindspListItems');
  listContainer.innerHTML = html;
  
}
//---------------------------------------------------------
//we get the total list in our site
//---------------------------------------------------------
// protected async onInit(): Promise<void> {

//   console.log("init");
//   this._siteLists= await this._getSiteLists();
//   return super.onInit();
// }

// private async _getSiteLists():Promise<string[]>{
//   const endpoint:string=`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title  &$orderby=Title &$top=10`;
//   const rawResponse:SPHttpClientResponse=await this.context.spHttpClient.get(endpoint,SPHttpClient.configurations.v1);
//   return(await rawResponse.json()).value.map((list:{Title:string})=>{return list.Title})
// }






  public render(): void {
    this.domElement.innerHTML = `
      <div class={styles.sharepointframe}>
    <div class={ styles.container }>
      <div class={ styles.row }>
        <div class="${ styles.column }">
          <span class="${ styles.title }">EVENT FUND DETAILS </span>
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
              }),
              // PropertyPaneTextField('siteUrl', {
              //   label: strings.webURLFieldLabel,
              //   onGetErrorMessage: this.onSiteUrlGetErrorMessage.bind(this),
              //   value: this.context.pageContext.site.absoluteUrl,
              //   deferredValidationTime: 1200,
              // }),
              // PropertyPaneDropdown('list', {
              //   label: strings.ListFieldLabel,
              //   options: this.lists,
              //   disabled: this.listsDropdownDisabled,
              // }),
              //PropertyPaneTextField('webURL', {
               // text: this.context.pageContext.web.absoluteUrl
               //label: strings.webURLFieldLabel
              //}),
              // PropertyPaneDropdown('selectedList', {
              //   label:"site lists",options:this._siteLists.map((list:string)=>{return <IPropertyPaneDropdownOption>{key:list,text:list}})
          
              // })
            //   PropertyPaneDropdown('libraryname', {  
            //     label: 'Select Library',  
            //     options: this.dropdownOptions 
            // }),
            //   PropertyPaneTextField('SiteUrl', {  
                 
            //     validateOnFocusOut: true, 
            //      label:strings.DescriptionFieldLabel,
              // text:document.getElementById('#SiteUrl').innerText;
               // onchange:this.bindLibraryNames1.bind(this);
               // onGetErrorMessage: this.bindLibraryNames1.bind(this)  
            //}),  
    
            
             
            ]
          }
        ]
      }
    ]
  };
}
}