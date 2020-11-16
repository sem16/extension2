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
import {CustomDialog} from './ExtensionDialog'
import * as ReactDOM from 'react-dom';
import { Convert } from './ConvertExcel';
import { ExportService } from './export-excel/exportService';

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
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters){
    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`${this.properties.sampleTextOne}`);
        break;
      case 'COMMAND_2':
        const dialogPlaceHolder = document.body.appendChild(document.createElement("div"));
        const lists = await sp.web.lists.filter('(hidden eq false) and (BaseTemplate eq 100)').get();
        // let views: IViewInfo[][] = [];
        // views = await Promise.all(lists.map(list => (sp.web.lists.getById(list.Id).views.select('Title').get())));
        // for(let i = 0; i < lists.length; i++){
        //   views[i] = sp.web.lists.getById(lists[i].Id).views.select('Title').get();
        // }
        // console.log(views)
        const element: React.ReactElement<{}> = React.createElement(
          CustomDialog,{
            hide: false,
            convert: new Convert(this.context),
            export: new ExportService(this.context),
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
