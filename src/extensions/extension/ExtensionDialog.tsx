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


interface CustomDialogProps{
  hide: boolean;
  convert: Convert;
  export: ExportService;
}
export class CustomDialog extends React.Component<CustomDialogProps, {}>{
  state= {
    show: this.props.hide
  }

  dialogContentProps = {
    type: DialogType.normal,
    title: 'Quick Import',
    closeButtonAriaLabel: 'Close',
  }
  modalprops  = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } },
    containerClassName: styles.alert,
    className: styles.alertBackground

  }

  public render(): React.ReactElement{

    return (
      <>
        <Dialog
          hidden={this.state.show}
          dialogContentProps={this.dialogContentProps}
          modalProps={this.modalprops}
          onDismiss={() => this.setState({show: true})}>
          <p>insersci il file</p>
          <input type="file" id="fileUpload" onChange={e => this.props.convert.ConvertAndInsert(e)}></input>
          <p>esporta contenuto lista</p>
          <button onClick={ () => this.props.export.getService()}></button>
        </Dialog>
      </>
    );
  }
}


