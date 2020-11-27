import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import {  sp } from "@pnp/sp-commonjs";
import * as XLSX from "xlsx";

export class Convert {
  public context: ListViewCommandSetContext;
  public title: string;
  public errors: any[] = [];
  public success: number;
  public length: number;
  constructor(context: ListViewCommandSetContext) {
    this.context = context;
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
    var excelRows: any = XLSX.utils.sheet_to_json(workbook.Sheets[Sheet], {
      defval: null,
    });
    console.log(excelRows);
    const allowed: string[] = ["__rowNum__"];

    //filtra ed inserisce nel array object i dati elimindano la proprieta 'rowNum'
    let object: {}[] = [];
    excelRows.forEach((row) => {
      object.push(
        Object.keys(row)
          .filter((key) => allowed.indexOf(key))
          .reduce((obj, key) => {
            obj[key] = row[key];
            return obj;
          }, {})
      );
      console.log(object);
    });
    //converte il titolo delle colonne ad nomi interni
    this.success = object.length;
    this.length = object.length;
    console.log(object);
    return object;
  }

  public async insertInList(objects: any[]) {
    console.log(objects);
    this.errors = [];
    for (let i = 0; i < objects.length; i++) {
      let object = objects[i];
      if (object["Modificato"] === undefined || object["Modificato"] === null) {
        delete object["Modificato"];
        delete object["Id"];
        const item = await sp.web.lists.getByTitle(this.title);
        let res;
        try{
          res = await item.items.add(object);
        }catch(e){
          console.log(e);
          this.errors[i] = object;
          this.success -= 1;
          continue;
        }
        if (res) {
          console.log("succes" + res);
        } else {
          this.errors.push(object);
          this.success -= 1;
        }
      } else if (object["Modificato"].toLowerCase() === "no") {
        delete objects[object];
      } else if (object["Modificato"].toLowerCase() === "sì" || "si") {
        delete object["Modificato"];
        const filteredList = await sp.web.lists
          .getByTitle(this.title)
          .items.filter(`Id eq '${parseInt(object.Id)}'`)
          .get();
        let result;
        try{
          result = await sp.web.lists
            .getByTitle(this.title)
            .items.getById(filteredList[0].ID)
            .update(object);
        }catch(e){
          console.log(e);
          this.errors[i] = object;
          console.log(this.errors);
          this.success -= 1;
          continue;
        }
        if (result !== undefined) {
          console.log(result);
        } else {
          this.errors[i] = object;
          console.log(this.errors);
          this.success -= 1;
        }

      }
    }

  }

  async fixNameAndType(object: any[]) {
    object.forEach((el) => {
      if (el["Modificato"] === undefined) {
        delete el["Modificato"];
      }
    });

    await sp.web.lists
      .getByTitle(this.title)
      .fields.get()
      .then((fields) => {
        console.log(fields);
        fields.forEach((field) => {
          object.forEach((el) => {
            Object.keys(el).forEach((key) => {
              if (field.Title === key) {
                if (field.InternalName !== key) {
                  el[field.InternalName] = el[key];
                  delete el[key];
                }

                switch (field.TypeAsString) {
                  case "Boolean":
                    try {
                      switch (el[field.InternalName].toLowerCase()) {
                        case "sì" || "si" || "true" || "yes":
                          el[field.InternalName] = true;
                          break;
                        case "no":
                          el[field.InternalName] = false;
                          break;
                        default:
                          el[field.InternalName] = null;
                          break;
                      }
                    } catch (e) {
                      //throw new Error(`errore nel campo ${key}; puo avere solo valori si/no` + e)
                    }

                    break;
                  case "DateTime":
                    if(el[field.InternalName] === ""){
                      el[field.InternalName] = null;
                    }
                    if (el[field.InternalName] !== null) {

                      if (typeof el[field.InternalName] === "string") {
                        let splitedDate = el[field.InternalName].split("/");
                        let date: Date = new Date();
                        date.setFullYear(
                          splitedDate[2],
                          splitedDate[1] - 1,
                          splitedDate[0]
                        );
                        el[field.InternalName] = date;
                      } else {
                        el[field.InternalName] = new Date(
                          (el[field.InternalName] - (25567 + 1)) * 86400 * 1000
                        );
                      }
                    }
                    if(isNaN(el[field.InternalName])){
                      el[field.InternalName] = 'Darai errore';
                    }
                    break;

                  case "Choice":
                    console.log(Object.keys(field)["Choices"]);
                    break;

                  case "URL":
                    if (el[field.InternalName] === "") {
                      el[field.InternalName] = null;
                    } else {
                      el[field.InternalName] = {
                        __metadata: { type: "SP.FieldUrlValue" },
                        Url: el[field.InternalName],
                      };
                    }
                    break;
                  case "Number":
                    if (el[field.InternalName] === "") {
                      el[field.InternalName] = null;
                    } else {
                      try {
                        el[field.InternalName] = parseInt(
                          el[field.InternalName]
                        );
                      } catch (e) {}
                    }
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

  public async ConvertAndInsert(fileUpload) {
    let temp: number;
    console.log(this.context);
    const result = await this.GetTableFromExcel(fileUpload);
    const fixed = await this.fixNameAndType(result);
    const t = await this.insertInList(fixed);
    return [this.success,Object.keys(this.errors).map(err => (parseInt(err) + 2))];
  }
}
