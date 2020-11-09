import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { sp } from "@pnp/sp-commonjs";
import { ConvertToXlsx } from "./ConvertToXlsx";
import { exclude } from "./excluded";

export class ExportService {
  private context;
  constructor(_context) {
    this.context = _context;
  }

  async cheangeColumnName(jsonList: any) {
    let keys;
    let fields = await sp.web.lists
      .getByTitle(this.context.pageContext.list.title)
      .fields.get();
    for (let i = 0; i < jsonList.length; i++) {
      exclude.forEach((element) => {
        try {
          delete jsonList[i][element];
        } catch (e) {
          console.log(e);
        }
      });

    }
    jsonList.forEach((column) => {
      fields.forEach((res) => {
        keys = Object.keys(column);
        keys.forEach((el) => {
          if (el === res.StaticName) {
            column[res.Title] = column[el];
            delete column[el];
          }
        });
      });
    });
  }

  public getService() {
      sp.web.lists
        .getByTitle(this.context.pageContext.list.title)
        .items
        .get()
        .then((res) => {
          console.log(res)
          this.cheangeColumnName(res).then(() =>
            ConvertToXlsx.convertToXslx(res,'test')
          );
          console.log(res);
        });
  }
}
