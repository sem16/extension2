import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  ColorPicker,
  PrimaryButton,
  Button,
  DialogFooter,
  Dialog,
  DialogContent,
  DialogType
} from 'office-ui-fabric-react';
import {Convert} from './ConvertExcel';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { sp } from '@pnp/sp-commonjs';
import styles from './Extension.module.scss';
import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import { ExportService } from './export-excel/exportService';

interface testProps{
  convert: Convert;
  export: ExportService;
}



class Test extends React.Component<testProps,{}>{
  public render(): JSX.Element{
    return (
      <DialogContent>
        <div>
          <p>insersci il file</p>
          <input type="file" id="fileUpload" onChange={e => this.props.convert.ConvertAndInsert(e)}></input>
          <p>esporta contenuto lista</p>
          <button onClick={ () => this.props.export.getService()}></button>
        </div>

      <DialogFooter>

      </DialogFooter>
      </DialogContent>
    );
  }
}

interface CustomDialogProps{
  hide: boolean;
}
export class CustomDialog extends React.Component<CustomDialogProps, {}>{
  state= {
    show: this.props.hide
  }

  dialogContentProps = {
    type: DialogType.normal,
    title: 'Missing Subject',
    closeButtonAriaLabel: 'Close',
    subText: 'Do you want to send this message without a subject?',
  }
  modalprops  = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } },
    containerClassName: styles.alert
  }

  public render(): React.ReactElement{

    return (
      <>
        <Dialog
          hidden={this.state.show}
          dialogContentProps={this.dialogContentProps}
          modalProps={this.modalprops}
          onDismiss={() => this.setState({show: true})}>
          <input type="file"></input>
        </Dialog>
      </>
    );
  }
}

export class FileDialog extends BaseDialog {
  export: ExportService;
  convert: Convert;
  constructor(context:ListViewCommandSetContext){
    super();
    //this.context = context;
    this.convert = new Convert(context);
    this.export = new ExportService(context);
  }


  // public context: ListViewCommandSetContext;
  public render(){
    ReactDOM.render(
        <Test convert={this.convert} export={this.export}></Test>
    ,this.domElement);
  }
}

