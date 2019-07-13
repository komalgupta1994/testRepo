const options = {
    filename: "./language.xlsx",
    useStyles: true,
    useSharedStrings: true
};
const workbook = new Excel.stream.xlsx.WorkbookWriter(options);

let commonSheet = workbook.addWorksheet("common");
let cartosSheet = workbook.addWorksheet("cartos");
let porticoSheet = workbook.addWorksheet("portico");
let pharosSheet = workbook.addWorksheet("pharos");
const sheetColumn = [
    { header: "Key", Key: "key", width: 40 },
    { header: "Language", Key: "lang", width: 100 }
];

commonSheet.columns = cartosSheet.columns = porticoSheet.columns = pharosSheet.columns = sheetColumn;

this.commonKeys.map(this.addRow.bind(null, commonSheet));
this.uniqueCartosKeys.map(this.addRow.bind(null, cartosSheet));
this.uniquePorticoKeys.map(this.addRow.bind(null, porticoSheet));
this.uniquePharosKeys.map(this.addRow.bind(null, pharosSheet));

commonSheet.commit();
cartosSheet.commit();
porticoSheet.commit();
pharosSheet.commit();
workbook.commit();

addRow(sheetName: any, row: object) {
    sheetName.addRow(row);

}