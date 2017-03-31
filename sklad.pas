unit sklad;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Grids, DBGrids, ComObj;

type
  TForm1 = class(TForm)
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    ADOConnection1: TADOConnection;
    Button1: TButton;
    ADOQuery1: TADOQuery;
    Button2: TButton;
    Button3: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
 
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin

ADOConnection1.Connected := true;
adoquery1.Active := true;
end;



procedure TForm1.Button2Click(Sender: TObject);
var
StarOffice, StarDesktop, Document, Sheets, Sheet, MyStruct, Cell: variant;
J, I: integer;
begin
  StarOffice := CreateOleObject('com.sun.star.ServiceManager');

  StarDesktop := StarOffice.createInstance('com.sun.star.frame.Desktop');
  Document := StarDesktop.LoadComponentFromURL(
                  'private:factory/scalc', '_blank', 0,
                  VarArrayCreate([0, -1], varVariant));
  MyStruct := StarOffice.Bridge_GetStruct('com.sun.star.table.BorderLine');

    try

    Sheets := Document.getSheets;
    Sheet := Sheets.getByName('ÀËÒÚ1');
    Cell := Sheet.getCellRangeByName('C10:C13');// cell := sheet.getCellRangeByPosition(0, 1, 0, 3);
    Cell.setPropertyValue('CellBackColor', 255);//  Cell.cellBackColor :=200;
 //   Cell.setPropertyValue('GoToCell');  //Í‡Í Ò‰

    J := 1; I := 1;
    Cell := Sheet.getCellByPosition(j, i);
   // Cell.charColor := 255;
    Cell.charHeight :=  14;
    Cell.CharWeight:= 150;
    Cell.charUnderline := true;
    Cell.charPosture := true;
    Cell.charStrikeout  := true;


    Cell.SetString('—¿À‹ƒŒ¬Œ-√–”œœ»–Œ¬Œ◊Õ¿ﬂ ¬≈ƒŒÃŒ—“‹ œŒ ”—À”√≈:');

    Cell := Sheet.getCellByPosition(j, i+1);
      //MyStruct.color := 255;
    MyStruct.lineDistance  := 0;
    MyStruct.innerLineWidth := 0;
    MyStruct.outerLineWidth := 1;
    Cell.leftBorder :=  MyStruct;
    Cell.rightBorder := MyStruct;
    Cell.topBorder := MyStruct;
    Cell.bottomBorder := MyStruct;

    Cell.SetString('ƒÓÏ: ');


  finally
    StarOffice := Unassigned;
  end;
end;

procedure TForm1.Button3Click(Sender: TObject);

var
i,j,index: Integer;
ExcelApp,sheet: Variant;

begin
 ExcelApp := CreateOleObject('Excel.Application');
 ExcelApp.Visible := true;
 ExcelApp.WorkBooks.Add(-4167);
 ExcelApp.WorkBooks[1].WorkSheets[1].name := 'Export';

 sheet:=ExcelApp.WorkBooks[1].WorkSheets['Export'];
 index:=3; //c ????? ?????? ????????? ? excel'e
 form1.DBGrid1.DataSource.DataSet.First;
 //form5.DBGrideh1.DataSource.DataSet.First;
   for i:=1 to  form1.DBGrid1.DataSource.DataSet.RecordCount do
      begin
   for j:=1 to form1.DBGrid1.FieldCount do
   sheet.cells[index,j]:=DBGrid1.fields[j-1].asstring;
   inc(index);
   form1.DBGrid1.DataSource.DataSet.Next;
      end;
end;

end.
