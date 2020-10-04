import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { sp } from '@pnp/sp-commonjs';
import * as XLSX from 'xlsx';

export class Convert{
  constructor(context:ListViewCommandSetContext){
    this.context = context;
  }
  public context: ListViewCommandSetContext;

  public async GetTableFromExcel(data) {
  console.log(data.target.files.item(0).name)
  const data1 = data = await data.target.files.item(0).arrayBuffer();
  var workbook = XLSX.read(data1, {
  type: 'array'
  });


  var Sheet = workbook.SheetNames[0];


  var excelRows: any = XLSX.utils.sheet_to_json(workbook.Sheets[Sheet]);

  console.log(excelRows);
  const allowed: string[] = ['__rowNum__'];

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

 insertInList(object: any){
   console.log(object);

    sp.web.lists.getByTitle(this.context.pageContext.list.title).items.add(object)
    .then(res => {console.log('succes'+res)}, res => {
      console.log('title: '+object.Title);
      console.log('res'+res);
      sp.web.lists.getByTitle(this.context.pageContext.list.title).items
        .filter(`Title eq '${object.Title}'`).get().then(res => {
          console.log(res)
          console.log(res[0].ID)
          sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(res[0].ID).update(object);
        });
      // const result = sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getAll().then(res =>{
      //   console.log(res);
      //   res.forEach(el => {
      //     if(object.Title === el.Title){
      //       sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(el.ID).update(object);
      //     }
      //   })
      // });
    });
}

ConvertAndInsert(fileUpload: React.ChangeEvent<HTMLInputElement>){
  console.log(this.context.pageContext.list.title)
  this.GetTableFromExcel(fileUpload);
}
}

