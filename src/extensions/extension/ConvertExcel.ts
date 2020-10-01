import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { sp } from '@pnp/sp-commonjs';
import * as XLSX from 'xlsx';

export class Convert{
  constructor(context:ListViewCommandSetContext){
    this.context = context;
    //saydsauidhasdjasdhaiusfaiPollo
  }
  public context: ListViewCommandSetContext;

  public async GetTableFromExcel(data) {
  console.log(data.target.files.item(0).name)
  const data1 = data = await data.target.files.item(0).arrayBuffer();
  var workbook = XLSX.read(data1, {
  type: 'array'
  });

  //get the name of First Sheet.
  var Sheet = workbook.SheetNames[0];

  //Read all rows from First Sheet into an JSON array.
  var excelRows: any = XLSX.utils.sheet_to_json(workbook.Sheets[Sheet]);
  // var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[Sheet]);
  console.log(excelRows);
  const allowed: string[] = ['__rowNum__'];
  //Create a HTML Table element.;
 let object: {};
 excelRows.forEach(el => {
    object =Object.keys(el).filter(key => allowed.indexOf(key)).reduce((obj, key) => {
    obj[key] = el[key];
    return obj;

  }, {});
  console.log(object);
  this.insertInList(object);
 });
 }

 insertInList(object: {}){
   console.log(object);

    sp.web.lists.getByTitle(this.context.pageContext.list.title).items.add(object).then(res => console.log(res));

}

ConvertAndInsert(fileUpload: React.ChangeEvent<HTMLInputElement>){
  console.log(this.context.pageContext.list.title)
  this.GetTableFromExcel(fileUpload);
}
}

