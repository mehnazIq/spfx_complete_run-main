import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudWrbpartWebPart.module.scss';
import * as strings from 'CrudWrbpartWebPartStrings';
export interface ICrudWrbpartWebPartProps {
  description: string;
}

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

interface IRegistrationDetails {
  Title: string;
  Department: string;
  ProjectName: string;
  Expense: number;
  Remarks: string;

  }

   
export default class CrudWrbpartWebPart extends BaseClientSideWebPart<ICrudWrbpartWebPartProps> {

  //private _isDarkTheme: boolean = false;
  //private _environmentMessage: string = '';
  private Listname: string = "Employee";
  private listItemId: number = 0;

  /*protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }*/


  public render(): void {
    this.domElement.innerHTML =`<div>
    <table>
    <tr>
    <td>Full Name</td>
    <td><input type="text" id="idFullName" name="fullName" placeholder="Enter Full Name"></td>
    </tr>
    <tr>
    <td>Department</td>
    <td><input type="text" id="idDepartment" name="department" placeholder ="Enter your Department"></td>
    </tr>
    <tr>
    <td>Project Name</td>
    <td><input type="text" id="idProjectName" name="projectname" placeholder="What is your ProjectName"></td>
    </tr>
    <tr>
    <td>Expense</td>
    <td><input type="number" id="idExpense" name="expense" placeholder="How much Expense" pattern="[0-9]{3}-[0-9]{3}-[0-9]{4}" required</td>
    </tr>
    <tr>
    <td>Remarks</td>
    <td><input type="text" id="idRemarks" name="remarks" placeholder="Write your remarks.."></td>
    </tr>
    </table>
    <table>
    <tr>
    <td><button class="${styles.button} find-Button" >Find</button></td>
    <td><button class="${styles.button} create-Button">Create</button></td>
    <td><button class="${styles.button} update-Button">Update</button></td>
    <td><button class="${styles.button} delete-Button">Delete</button></td>
    <td><button class="${styles.button} clear-Button">Clear</button></td>
    </tr>
    </table>
    <div id="tblRegistrationDetails"></div>
    </div>
    `;
    this.setButtonsEventHandlers();
    this.getListData();  
  }


  private setButtonsEventHandlers(): void {
    const webPart: CrudWrbpartWebPart = this;
    this.domElement.querySelector('button.find-Button').addEventListener('click', () => { webPart.find(); });
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.create(); });
    //this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.update(); });
    //this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.delete(); });
    this.domElement.querySelector('button.clear-Button').addEventListener('click', () => { webPart.clear(); });
  }

  private find(): void {
    
    let emailId = prompt("Enter the Email ID");
    var siteUrl = this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/
    items?$select=*&$filter=EmailID eq '${emailId}'`, SPHttpClient.configurations.v1)
    .then(response => {
    return response.json()
    .then((item: any): void => {
      console.log(siteUrl);
      (document.getElementById('idFullName') as HTMLInputElement).value =item.value[0].Title;
    //document.getElementById('idFullName')["value"] = item.value[0].Title;
    (document.getElementById('idDepartment') as HTMLInputElement).value  = item.value[0].Department;
    //document.getElementById('idAddress')["value"] = item.value[0].Address;
    (document.getElementById('idProjectName') as HTMLInputElement).value = item.value[0].ProjectName;
    // document.getElementById('idPhoneNumber')["value"] = item.value[0].Mobile;
    (document.getElementById('idExpense') as HTMLInputElement).value = item.value[0].Expense;
      //document.getElementById('idEmailId')["value"] = item.value[0].EmailID;
      (document.getElementById('idRemarks') as HTMLInputElement).value = item.value[0].Remarks;

   
    this.listItemId = item.value[0].Id;
    });
    });
    }
         // This Function All List Item 
    private getListData() {
      
      let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
     // html += '<th>Full Name</th><th>Address</th><th>Phone Number</th> <th>Email ID</th> ';
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`, SPHttpClient.configurations.v1)
      .then(response => {
      return response.json()
      .then((items: any): void => {
      console.log('items.value: ', items.value);
      const listItems: IRegistrationDetails[] = items.value;
      console.log('list items: ', listItems);
       
      listItems.forEach((item: IRegistrationDetails) => {
       html += `
      <tr>
      <td>${item.Title}</td>
      <td>${item.Department}</td>
      <td>${item.ProjectName}</td>
      <td>${item.Expense}</td>
      <td>${item.Remarks}</td>
      
      </tr>
      `;
      });
      html += '</table>';
      const listContainer: Element = this.domElement.querySelector('#tblRegistrationDetails');
      listContainer.innerHTML = html;
      });
      });
    }
  

    private create(): void {
     // this Method Created By Gulam Khan 
      //   Declare All Variable Locally   *//
      let  Fullname    = (document.getElementById("idFullName")as HTMLInputElement).value;
      let Department   = (document.getElementById("idDepartment") as HTMLInputElement).value;
      let ProjectName  = (document.getElementById("idProjectName") as HTMLInputElement).value;
      let Expense      = (document.getElementById("idExpense")as HTMLInputElement).value;
      let Remarks      = (document.getElementById("idRemarks")as HTMLInputElement).value;
     

         //console.log("hello"+emailRegexp.test(Email));
    if (Fullname  =="") {
      alert("Enter Full Name!");
    }
    else if(Department  ==""){
      alert("Enter your Department!")
    }
    else if(ProjectName =="") {
         
      alert("Enter your Project Name");
    } 

    else if(Expense  =="") {
         
      alert("Enter Expense !");
    } 

    else if(Remarks =="") {
         
      alert("Add Remarks!");
    } 

    else 
    
    {
        //console.log("else parts");
      const body: string = JSON.stringify({
      'Title': (Fullname ),
      'Department':(Department),
      'ProjectName': (ProjectName),
      'Expense': (Expense),
      'Remarks': (Remarks)
      });
       
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'X-HTTP-Method': 'POST'
      },
      body: body
      }).then((response: SPHttpClientResponse): void => {
        
      this.getListData(); 
      
      this.clear();
      alert('Item has been successfully Saved ');
      }, (error: any): void => {
      alert(`${error}`);
      });
    } 
    }

    
    private clear(): void {
      (document.getElementById('idFullName') as HTMLInputElement).value  ="";
     (document.getElementById('idDepartment') as HTMLInputElement).value    ="";
     (document.getElementById('idProjectName') as HTMLInputElement).value  ="";
     (document.getElementById('idExpense') as HTMLInputElement).value    ="";
     (document.getElementById('idRemarks') as HTMLInputElement).value    ="";
     
     
 }

      /*private update(): void {
       
       
     
        const body: string = JSON.stringify({
          'Title': (document.getElementById('idFullName') as HTMLInputElement).value,
          'Address': (document.getElementById('idAddress') as HTMLInputElement).value,
          'Mobile': (document.getElementById('idPhoneNumber') as HTMLInputElement).value,
          'EmailID': (document.getElementById('idEmailId') as HTMLInputElement).value
          
        });
         
        var current_url = this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${this.listItemId})`,
        
        SPHttpClient.configurations.v1,
        {
        headers: {
        'Accept': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'PATCH'
        },
        body: body
        }).then((response: SPHttpClientResponse): void => {
        this.getListData();
        this.clear();
        console.log(current_url);
        alert(`Item successfully updated`);
        }, (error): void => {
        alert(`${error}`);
        });
        }


        private delete(): void {
          
          if (!window.confirm('Are you sure you want to delete the latest item?')) {
          return;
          }
           
          this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${this.listItemId})`,
          SPHttpClient.configurations.v1,
          {
          headers: {
          'Accept': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
          }
          }).then((response: SPHttpClientResponse): void => {
          alert(`Item successfully Deleted`);
          this.getListData();
          this.clear();
          }, (error: any): void => {
          alert(`${error}`);
          });
          }*/




 /* private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }*/

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
