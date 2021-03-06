import * as React from 'react';
import {
  Dialog,
  DialogType
} from 'office-ui-fabric-react';
import {Convert} from './ConvertExcel';
import { IListInfo, IViewInfo, Lists, sp } from '@pnp/sp-commonjs';
import styles from './Extension.module.scss';
import { ConvertToXlsx } from './ConvertToXlsx';

interface CustomDialogProps{
  hide: boolean;
  convert: Convert;
  export: ConvertToXlsx;
  lists: IListInfo[];
}
export class CustomDialog extends React.Component<CustomDialogProps, {}>{
  private file: FileList;

  public state= {
    show: this.props.hide,
    views: undefined,
    message: 'trascina un file qui',
    errors: Object.keys(this.props.convert.errors),
    okay: 0,
    hideErrors: true
  };

  private dialogContentProps = {
    type: DialogType.normal,
    title: 'Quick Import',
    closeButtonAriaLabel: 'Close',
  };
  private modalprops  = {
    isBlocking: false,
    containerClassName: styles.alert,
    className: styles.alertBackground

  };
  public render(): JSX.Element{
    console.log(this.props.lists);
    let optionGroup = (T,i) => {return (<optgroup label={T.Title}>
      {this.state.views !== undefined ?
        this.state.views[i].map(view => (
          <option key={view.Title} value={this.props.lists[i].Title}>
            {view.Title}
          </option>))
        : []
        }
      </optgroup>);
    };
    return (
      <>
        <Dialog
          hidden={this.state.show}
          dialogContentProps={this.dialogContentProps}
          modalProps={this.modalprops}
          onDismiss={() => this.setState({show: true})}>
          <p className={styles.warning}>Prima di selezionare il file con i contenuti da importare, selezionare una o più righe della lista e cliccare su "Admin Export".</p>
          <div className={styles.dragAndDrop}
          onDrop={e => {
            e.preventDefault()
            this.file = e.dataTransfer.files;
            e.currentTarget.style.border = '2px solid #ffffff00';
            this.setState({message: this.file[0].name});
          }}
          onDragOver={e => {e.preventDefault(); e.currentTarget.style.border = '2px solid #3369ff';}}
          onDragLeave={e => {e.preventDefault(); e.currentTarget.style.border = '2px solid #ffffff00';} }
          >
            <p>{this.state.message}</p>
          </div>

          <input type="file"
          id="fileUpload"
          onChange={e => {
            this.file = e.target.files;
            this.setState({message: this.file[0].name});
          }}
          style ={{display: 'none'}}/>

          <button onClick={() => document.getElementById('fileUpload').click()}
           className={styles.fileInput}>oppure scegli un file</button>

          <button onClick={() => {
            this.props.convert.ConvertAndInsert(this.file)
            .then(res => {
              this.setState({okay: res[0]});
              this.setState({errors: res[1]});
              this.setState({hideErrors: false});
          });

          }}
           className={styles.import}>import</button>
          <p  className={styles.status} hidden={this.state.hideErrors}>
            {`riusciti: ${this.state.okay} `} {this.props.convert.length !== undefined ? `/ ${this.props.convert.length }` : null}
            <br/>
            {this.state.errors.length > 0 ? `errore durante l inserimento delle seguenti righe: ${this.state.errors}` : null}
          </p>
        </Dialog>

      </>
    );
  }
}

