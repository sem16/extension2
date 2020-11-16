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
import { IListInfo, IViewInfo, Lists, sp } from '@pnp/sp-commonjs';
import styles from './Extension.module.scss';
import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import { ExportService } from './export-excel/exportService';


interface CustomDialogProps{
  hide: boolean;
  convert: Convert;
  export: ExportService;
  lists: IListInfo[];
  // views: IViewInfo[][];
}
export class CustomDialog extends React.Component<CustomDialogProps, {}>{
  public state= {
    show: this.props.hide,
    views: undefined
  };

  private dialogContentProps = {
    type: DialogType.normal,
    title: 'Quick Import',
    closeButtonAriaLabel: 'Close',
  };
  private modalprops  = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } },
    containerClassName: styles.alert,
    className: styles.alertBackground

  };
  public async  componentDidMount(){
    let views: IViewInfo[][] = [];
    views = await Promise.all(this.props.lists.map(list => (sp.web.lists.getById(list.Id).views.select('Title').get())));
    this.setState({views: views});
    console.log(this.state);
  }
  public render(): JSX.Element{
    console.log(this.props.lists);
    let file;
    let option = (T) => {return <option>{T.Title}</option>}
    let optionGroup = (T,i) => {return (<optgroup label={T.Title}>
      {this.state.views !== undefined ?
        this.state.views[i].map(option)
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
          <p>insersci il file</p>
          <input type="file" id="fileUpload" onChange={e => file = e}></input>
          <select onChange={e => console.log(e.target.value)}>{this.props.lists.map(optionGroup)}</select>
          <p>esporta contenuto lista</p>
          <button onClick={ () => this.props.export.getService()}>export</button>
          <button onClick={() => this.props.convert.ConvertAndInsert(file)}>import</button>
        </Dialog>
      </>
    );
  }
}

// interface ViewSelectInterface{
//   lists: IListInfo[];
//   views: IViewInfo[];
// }

// class ViewSelect extends React.Component<{},{}>{
//   componentDidMount(){}
//   public render(): JSX.Element {
//     let option = (T) => {return <option key={T.Title}>{T.Title}</option>}
//     let optionGroup = (T,i) => {return (<optgroup label={T.Title}>
//       {this.state.views[i].map(option)}
//       </optgroup>);
//     };
//     return (
//       <>
//         <select>
//         {this.props.lists.map(optionGroup)}
//         </select>
//       </>
//     );
//   }
// }
