import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { IFieldInfo, IFields, sp } from "@pnp/sp-commonjs";
import * as XLSX from "xlsx";

export class Convert {
  public context: ListViewCommandSetContext;
  public title: string;
  constructor(context: ListViewCommandSetContext) {
    this.context = context;
  }
  private async shouldBeUpdated(newItem){
    const keys = Object.keys(newItem);
    const item = await sp.web.lists.getByTitle(this.title).items.select(...keys).get();
    let temp: boolean;
    console.log(keys);
    console.log(item);
    item.forEach(el => {
      console.log('item: ' + el + 'newItem' + newItem[el]);
      if(el !== newItem[el]){
        temp = true;
      }
    });
    return temp;
  }
  public async GetTableFromExcel(data) {
    //legge il file excel
    console.log(data.item(0).name);
    const data1 = (data = await data.item(0).arrayBuffer());
    var workbook = XLSX.read(data1, {
      type: "array",
    });

    //estrae i dati dal file e gli inserisce in excelRows
    var Sheet = workbook.SheetNames[0];
    var excelRows: any = XLSX.utils.sheet_to_json(workbook.Sheets[Sheet]);
    console.log(excelRows);
    const allowed: string[] = ["__rowNum__"];

    //filtra ed inserisce nel array object i dati elimindano la proprieta 'rowNum'
    let object: {}[] = [];
    excelRows.forEach((el) => {
      object.push(
        Object.keys(el)
          .filter((key) => allowed.indexOf(key))
          .reduce((obj, key) => {
            obj[key] = el[key];
            return obj;
          }, {})
      );
      console.log(object);
    });
    //converte il titolo delle colonne ad nomi interni
    console.log(object);
    return object;
  }

  public insertInList(objects: any[]) {

    console.log(objects);

    objects.forEach((object,i) => {
      this.shouldBeUpdated(object).then(res =>
      {if(res){
        console.log('modif',i);
      }
    });

      if(object['Modificato'] === undefined || null){
        delete object['Modificato'];
        sp.web.lists
        .getByTitle(this.title)
        .items.add(object)
        .then(
          (res) => {
            console.log("succes" + res);
        });
      }
      else if(object['Modificato'].toLowerCase() === 'no'){
        delete objects[object];
      }else if (object['Modificato'].toLowerCase() === 'sì' || 'si'){
      delete object['Modificato'];
        sp.web.lists
          .getByTitle(this.title)
          .items.filter(`Id eq '${parseInt(object.Id)}'`)
          .get()
          .then((res) => {
            console.log(res);
            console.log(res[0].Id);
            sp.web.lists
              .getByTitle(this.title)
              .items.getById(res[0].ID)
              .update(object);
        });
      }
    });
  }

    async fixNameAndType(object: any[]){
      object.forEach((el) => {
        if(el['Modificato'] === undefined){
          delete el['Modificato'];
        }
      })

      await sp.web.lists
      .getByTitle(this.title)
      .fields.get()
      .then((fields) => {
        console.log(fields)
        fields.forEach((field) => {
          object.forEach((el) => {
            Object.keys(el).forEach((key) => {
              if (field.Title === key) {
                if(field.InternalName !== key){
                  el[field.InternalName] = el[key];
                  delete el[key];
                }

                switch (field.TypeAsString) {
                  case 'Boolean':
                    try{
                      switch (el[field.InternalName].toLowerCase()) {
                        case 'sì' || 'si' || 'true' || 'yes':
                          el[field.InternalName] = true;
                          break;
                        case 'no ':
                          el[field.InternalName] = false;
                          break;
                        default:
                          break;
                      }
                    }catch(e){
                      throw new Error(`errore nel campo ${key}; puo avere solo valori si/no` + e)
                    }

                    break;
                  case 'DateTime':
                    if(typeof el[field.InternalName] === 'string'){
                      let splitedDate = el[field.InternalName].split('/');
                      let date: Date = new Date();
                      date.setFullYear(
                        splitedDate[2],
                        splitedDate[1]-1,
                        splitedDate[0]);
                      el[field.InternalName]  = date;
                    }else{
                      el[field.InternalName] = new Date((el[field.InternalName] - (25567 + 1))*86400*1000);
                    }

                    break;
                  case 'Choice':
                    console.log(Object.keys(field)['Choices'])
                    break;
                  default:
                    break;
                }
                console.log(object);
              }
            });
          });
        });
      });
      return object;
    }

  ConvertAndInsert(fileUpload: React.ChangeEvent<HTMLInputElement>) {
    console.log(this.context);
    // this.title = this.context.dynamicDataProvider.getAvailableSources()[1].metadata.title;
    // this.title = this.context.pageContext.list.title;
    // console.log(this.title);
    this.GetTableFromExcel(fileUpload).then((result) =>
      this.fixNameAndType(result).then(fixed =>
         {this.insertInList(fixed)
         console.log(fixed)}
         )
    );
  }
}


