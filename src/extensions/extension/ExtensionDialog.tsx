import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  Dialog,
  DialogType
} from 'office-ui-fabric-react';
import {Convert} from './ConvertExcel';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { IListInfo, IViewInfo, Lists, sp } from '@pnp/sp-commonjs';
import styles from './Extension.module.scss';
import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
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
    message: 'trascina un file qui'
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
  public async  componentDidMount(){
    let views: IViewInfo[][] = [];
    views = await Promise.all(this.props.lists.map(list => (sp.web.lists.getById(list.Id).select('Title').views.get())));
    this.setState({views: views});
  }
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
          onChange={e => {this.file = e.target.files;
            this.setState({message: this.file[0].name})}}
          style ={{display: 'none'}}/>

          <button onClick={() => document.getElementById('fileUpload').click()} className={styles.fileInput}>oppure scegli un file</button>
          {/* <select onChange={e => this.props.convert.title = e.target.value}>{this.props.lists.map(optionGroup)}</select> */}

          <button onClick={() => this.props.convert.ConvertAndInsert(this.file) } className={styles.import}>import</button>
        </Dialog>
      </>
    );
  }
}

