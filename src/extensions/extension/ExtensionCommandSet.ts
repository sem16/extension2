import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {IViewInfo, sp, ViewScope} from '@pnp/sp-commonjs';
import * as strings from 'ExtensionCommandSetStrings';
import * as React from 'react';
import {CustomDialog} from './ExtensionDialog';
import * as ReactDOM from 'react-dom';
import { Convert } from './ConvertExcel';
import { ConvertToXlsx } from './ConvertToXlsx';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IExtensionCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ExtensionCommandSet';


export default class ExtensionCommandSet extends BaseListViewCommandSet<IExtensionCommandSetProperties> {

  private data: {}[];

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ExtensionCommandSet');
    sp.setup({pageContext: {web: {absoluteUrl: this.context.pageContext.web.absoluteUrl}}});
    console.log(this.context);
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length > 0;
      this.data = [];
      console.log(event.selectedRows);
      console.log(this.context);
      console.log(sp.web.getParentWeb());
      console.log(this.context.dynamicDataProvider.getAvailableSources().map(el => el.metadata.instanceId))
      event.selectedRows.forEach((row,i) => {
        let values: any = {};
        try{
        values['Id'] = row.getValueByName('ID');
        values['Nome società'] = row.getValueByName('Nome_societa_quick');
        }

        catch{}
        row.fields.forEach((field) => {
          let keyName = field.displayName;
          values[keyName]  = row.getValue(field);
        });
        if(values['Nome società'] === undefined){
          delete values['Nome società'];
        }
        try{
          delete values['Attachments'];
          delete values['Allegati'];
        }catch{}
        values['Modificato'] = 'no';
        this.data[i] = values;
      });
      console.log(this.data);
    }
  }

  getTitle(): string{
    const itemKeys =  Object.keys(window['$ic'].states[0].itemMap['items'][4])
    let listFacet: string;
    let listTitle: string;
    itemKeys.forEach(el => {
      if(el.match('listFacet_') !== null){
        listTitle = window['$ic'].states[0].itemMap['items'][4][el].title;
      }
    });
    return listTitle;
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters){
    let listTitle;
    switch (event.itemId) {
      case 'COMMAND_1':
        let url= this.context.pageContext.site.serverRequestPath;
        let arrOfStr:string[] = url.split("/");
        let listName: String;
        console.log(arrOfStr);
        for(let I=0; I<arrOfStr.length;I++){​​​​
          if(arrOfStr[I]==="Lists" || arrOfStr[I]==="SitePages"){​​​​
          console.log(arrOfStr[I+1])
           listName=arrOfStr[I+1];
           listName = listName.replace(".aspx","");
          }​​​​
        }​​​​
        ConvertToXlsx.convertToXslx(this.data,listName);
        break;
      case 'COMMAND_2':
        const dialogPlaceHolder = document.body.appendChild(document.createElement("div"));
        const lists = await sp.web.lists.filter('(hidden eq false) and (BaseTemplate eq 100)').get();
        console.log('ee');
        try{
          console.log(window['$ic'].states[0].itemMap['items'][4].listFacet_385.title);
        }catch(e){}
        if(window.location.href.match('SitePages') != null ){
          try{
            listTitle = this.getTitle();
          }catch(e){}
        }else{
            listTitle = this.context.pageContext.list.title;
        }
        console.log(listTitle);
        // let views: IViewInfo[][] = [];
        // views = await Promise.all(lists.map(list => (sp.web.lists.getById(list.Id).views.select('Title').get())));
        // for(let i = 0; i < lists.length; i++){
        //   views[i] = sp.web.lists.getById(lists[i].Id).views.select('Title').get();
        // }
        // console.log(views)
        const convert: Convert = new Convert(this.context)
        convert.title = listTitle;
        const element: React.ReactElement<{}> = React.createElement(
          CustomDialog,{
            hide: false,
            convert: convert,
            export: new ConvertToXlsx(),
            lists: lists,
            // views: views
          }
        );
        ReactDOM.render(element,dialogPlaceHolder);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
