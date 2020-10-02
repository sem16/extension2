import { BaseDialog, Dialog,IDialogConfiguration } from '@microsoft/sp-dialog';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  ColorPicker,
  PrimaryButton,
  Button,
  DialogFooter,
  DialogContent
} from 'office-ui-fabric-react';
import {Convert} from './ConvertExcel';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { sp } from '@pnp/sp-commonjs';
import styles from './Extension.module.scss';

interface testProps{
  convert: Convert;
}

class Test extends React.Component<testProps,{}>{
  public render(): JSX.Element{
    return (
      <DialogContent className={styles.alert}>
        <div>
          <p>insersci il file</p>
          <input type="file" id="fileUpload" onChange={e => this.props.convert.ConvertAndInsert(e)}></input>

        </div>
      </DialogContent>
    );
  }
}

export class FileDialog extends BaseDialog {
  constructor(context:ListViewCommandSetContext){
    super();
    this.context = context;
    this.convert = new Convert(context);
  }
  convert: Convert;
  public context: ListViewCommandSetContext;

  public render(){
    ReactDOM.render(
      <Test  convert={this.convert}></Test>
    ,this.domElement);
  }
}
