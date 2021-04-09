import * as React from 'react';
import * as ReactDom from 'react-dom';
import {IUserItem} from '../spFxAadHttpClient/Models/IUserItem';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import{AadHttpClient,HttpClientResponse,MSGraphClient} from '@microsoft/sp-http'
  

import * as strings from 'SpFxAadHttpClientWebPartStrings';
import SpFxAadHttpClient from './components/SpFxAadHttpClient';
import { ISpFxAadHttpClientProps } from './components/ISpFxAadHttpClientProps';

export interface ISpFxAadHttpClientWebPartProps {
  description : string;
}

export default class SpFxAadHttpClientWebPart extends BaseClientSideWebPart<ISpFxAadHttpClientWebPartProps> {

  
  public render(): void {
     
    if(!this.renderedOnce){
      let msClient:MSGraphClient = null;
    this.context.msGraphClientFactory.getClient().then((client:any)=>{
     msClient = client;
    })
      this._getUser().then((results:IUserItem[])=>{

        const element: React.ReactElement<ISpFxAadHttpClientProps> = React.createElement(
          SpFxAadHttpClient,
          {
            userItems: results,
            graphClient:msClient
            
          }
        );
        ReactDom.render(element, this.domElement);
      })
    }
   
       }
       private _getUser():Promise<IUserItem[]>{
        return new Promise<IUserItem[]>((resolve,reject)=>{
         this.context.aadHttpClientFactory.getClient('https://graph.microsoft.com')
         .then((aadClient:AadHttpClient)=>{
           const endpoint:string = 'https://graph.microsoft.com/v1.0/users?$top=10&$select=id,displayName,mail';
           aadClient.get(endpoint,AadHttpClient.configurations.v1)
           .then((rawResponse:HttpClientResponse)=>{
             return rawResponse.json();
           })
           .then((jsonResponse:any)=>{
             resolve(jsonResponse.value);
            console.log("Value : " + jsonResponse.value)
           })
           .catch((error:any)=>{
             reject(error.value);
                  
           })
         })
      
        })
      
        }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
