unit main;              //XPBCJRSKKURVXMT

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBGridEhGrouping, ToolCtrlsEh, Printers, DBGridEhImpExp,
  GridsEh,  DBGridEh, ADODB, DB, StdCtrls, ComCtrls,   Menus,
  TeEngine, Series, ExtCtrls, TeeProcs, Chart, DbChart
  , ComObj, StrUtils,EhLibADO, BubbleCh, DBGridEhToolCtrls, DynVarsEh,
  EhLibVCL, DBAxisGridsEh, PrnDbgeh;

type
  TForm1 = class(TForm)
    Button1: TButton;
    ListBox1: TListBox;
    GroupBox1: TGroupBox;
    Button4: TButton;
    DS_S: TDataSource;
    Qur_S: TADOQuery;
    ADOC_S: TADOConnection;
    DbgEh_S: TDBGridEh;
    StatusBar1: TStatusBar;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    ListBox2: TListBox;
    Button2: TButton;
    Button3: TButton;
    Button5: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    XT1: TMenuItem;
    DBChart1: TDBChart;
    N5: TMenuItem;
    N6: TMenuItem;
    Series1: TPointSeries;
    Memo1: TMemo;
    Memo2: TMemo;
    ADOQuery2: TADOQuery;
    ADOConnection2: TADOConnection;
    DataSource2: TDataSource;
    PrintDBGridEh2: TPrintDBGridEh;
    Button6: TButton;
    procedure FormCreate(Sender: TObject);
    procedure StatusBar1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure ListBox2DblClick(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure ListBox2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure XT1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure DBChart1DblClick(Sender: TObject);
    procedure DbgEh_SDblClick(Sender: TObject);
    procedure Button6Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  E_App, E_Books, E_Sheet: OleVariant;
  TT,C_PL:string;
implementation

{$R *.dfm}

function DATA_Connect(Data_Name,Tab_Name:string;TF:Boolean):Boolean;      //accessConnect函数目的：连接Access
var i,m:Integer;
begin
        // 以下为初始化数据库联结
 Form1.ADOC_S.Connected:=False;       //设置数据库联结为False，关闭联结，以便于下一步联结
 Form1.ADOC_S.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_name+';Persist Security Info=False';
 Form1.ADOC_S.LoginPrompt:=false;        //不显示登录框

 Form1.Qur_S.Connection:=Form1.ADOC_S;
 Form1.DS_S.DataSet:=Form1.Qur_S;

 Form1.ListBox1.Clear;
 with Form1.Qur_S do                         //查出存在的剖面条数及编号
 begin
   Close;
   SQL.Clear;
   SQL.Add('select distinct 孔号 from '+Tab_Name);  //Caran_GJL');
   Open;
 end;
  m:=Form1.Qur_S.RecordCount-1;

  for i:=0 to m do                                   //将剖面编号加入Combobox1的选项中
  begin
    Form1.ListBox1.Items.Add(Form1.Qur_S.fieldbyname('孔号').AsString);
//    ListBox2.ItemIndex:=0;
    Form1.Qur_S.Next;
  end;
 result:=true;                     //函数必须有返回值，为不现出警告信息，返回true   }
end;

function Ex_Tro(E_Name:String):Boolean;      //读取Excel表
 var maxR,i:Integer;
  Str_Q:string;
begin

 Result:=True;
end;


function Ex_TroG(E_Name:String):Boolean;      //读取Excel表
 var maxR,i:Integer;
  Str_Q:string;
begin
  maxR:= E_Sheet.usedrange.rows.count;     //有数据的行数
  Str_Q:='insert into Caran_GJL (孔号,段次,灌浆起深,灌浆终深,灌浆段长,孔径,透水率,水灰比起,水灰比终,注入率起,注入率终,水泥注浆量,水泥注灰量,水泥废弃量,水泥合计,单位注入量,灌浆压力,时间起,时间终,时长,备注) values';
  Str_Q:=Str_Q+'(:I_D,:D_i,:D_S,:D_E,:D_L,:D_R,:K_L,:K_s,:K_e,:Z_S,:Z_E,:V_a,:V_b,:V_c,:V_v,:V_m,:V_p,:T_S,:T_E,:T_T,:K_B)';
 for i:=6 to maxR do
 begin
  if not TryStrToInt(Trim(E_Sheet.cells.item[i,1]),maxR) then continue;                     //如果段次不是数字则跳出本次循环
  if (Trim(E_Sheet.cells.item[i,2])<>'') and (Trim(E_Sheet.cells.item[i,3])<>'') and  (Trim(E_Sheet.cells.item[i,1])<>'fk') then       //如果本行不为空则运行
  with Form1.Qur_S do
  begin
    Close;
    SQL.Clear;
    SQL.Add(Str_Q);
		Parameters.ParamByName('I_D').Value:=E_Name;		//孔号
 		Parameters.ParamByName('D_i').Value:=StrToInt(E_Sheet.cells.item[i,1]);  //E_Sheet.cells[i,1].value;		//段次
 		Parameters.ParamByName('D_S').Value:=E_Sheet.cells[i,2].value;		//灌浆起深
 		Parameters.ParamByName('D_E').Value:=E_Sheet.cells[i,3].value;		//灌浆终深
		Parameters.ParamByName('D_L').Value:=E_Sheet.cells[i,4].value;		//灌浆段长
    Parameters.ParamByName('D_R').Value:=StrToInt(RightStr(E_Sheet.cells[i,5].value, 2));		//孔径
    if Trim(E_Sheet.cells.item[i,6])='/' then Parameters.ParamByName('K_L').Value:=0 else Parameters.ParamByName('K_L').Value:=E_Sheet.cells[i,6].value;		//透水率
 		Parameters.ParamByName('K_s').Value:=E_Sheet.cells[i,7].value;		//水灰比起
		Parameters.ParamByName('K_e').Value:=E_Sheet.cells[i,8].value;		//水灰比终
		Parameters.ParamByName('Z_S').Value:=E_Sheet.cells[i,9].value;		//注入率起
		Parameters.ParamByName('Z_E').Value:=E_Sheet.cells[i,10].value;		//注入率终
		Parameters.ParamByName('V_a').Value:=E_Sheet.cells[i,11].value;		//水泥量浆
		Parameters.ParamByName('V_b').Value:=E_Sheet.cells[i,12].value;		//水泥量灰
		Parameters.ParamByName('V_c').Value:=E_Sheet.cells[i,13].value;		//水泥量废
		Parameters.ParamByName('V_v').Value:=E_Sheet.cells[i,14].value;		//水泥量合
		Parameters.ParamByName('V_m').Value:=E_Sheet.cells[i,15].value;		//单位注入量
		Parameters.ParamByName('V_p').Value:=E_Sheet.cells[i,16].value;		//灌浆压力

    Parameters.ParamByName('T_S').Value:=E_Sheet.cells[i,17].value;
    Parameters.ParamByName('T_E').Value:=E_Sheet.cells[i,19].value;

{    Parameters.ParamByName('T_S').Value:=FormatdateTime('c',StrToDateTime(DateToStr(E_Sheet.cells[i,17].value)+' '+TimeToStr(E_Sheet.cells[i,18].value)));
    Parameters.ParamByName('T_E').Value:=FormatdateTime('c',StrToDateTime(DateToStr(E_Sheet.cells[i,19].value)+' '+TimeToStr(E_Sheet.cells[i,20].value)));
} 		Parameters.ParamByName('T_T').Value:=FormatdateTime('tt',E_Sheet.cells[i,21].value);		//时长
		Parameters.ParamByName('K_B').Value:=E_Sheet.cells[i,22].value;		//备注
    ExecSQL;                    //打开数据集，执行SQL语句
  end
  else Break;
 end;
 Result:=True;
end;

 function Data_EhH(D_Eh:TDBGridEh):Boolean;
 var Col:TColumnEh;
 begin
    D_Eh.UseMultiTitle:=True; //是否使用多行标题行
    D_Eh.TitleLines:=2; //标题行行数
    D_Eh.Flat:=True; //平面显示;False为立体显示
    D_Eh.Columns.Clear;

		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='段次';
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='灌浆孔段(m)|自'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='灌浆孔段(m)|至'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='灌浆孔段(m)|段长'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='孔径'+#13+'(mm)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='透水率'+#13+'(Lu)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='水灰比|起始'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='水灰比|终止'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='注入率'+#13+'(L/min)|起始'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='注入率'+#13+'(L/min)|终止'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='水泥用量|注浆'+#13+'(L)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='水泥用量|注灰'+#13+'(kg)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='水泥用量|废弃'+#13+'(kg)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='水泥用量|合计'+#13+'(kg)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='单位'+#13+'注入量'+#13+'(kg/m)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='灌浆'+#13+'压力'+#13+'(Mpa)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='灌浆时间|起始'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='灌浆时间|终止'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='灌浆时间|纯灌'+#13+'(hh:mm)'; //标题行文本
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='备注'; //标题行文本

   with D_Eh do
   begin
                 {
		Columns[0].Width:=20;
		Columns[1].Width:=30;
		Columns[2].Width:=30;
		Columns[3].Width:=30;
		Columns[4].Width:=30;
		Columns[5].Width:=40;
		Columns[6].Width:=35;
		Columns[7].Width:=40;
		Columns[8].Width:=40;
		Columns[9].Width:=40;
		Columns[10].Width:=50;
		Columns[11].Width:=50;
		Columns[12].Width:=35;
		Columns[13].Width:=50;
		Columns[14].Width:=45;
		Columns[15].Width:=40;
		Columns[16].Width:=110;
		Columns[17].Width:=110;
		Columns[18].Width:=40;
		Columns[19].Width:=75;
                       }
    Columns[6].Alignment:=taRightJustify;         //单元中内容居中
    Columns[7].Alignment:=taRightJustify;         //单元中内容居中
    Columns[18].Alignment:=taRightJustify;         //单元中内容居中

//    Columns[0].ReadOnly:=True;                  //设置表格的部分列不可写
   end;
    Result:=True;
 end;

 function Data_EhR(D_Eh:TDBGridEh):Boolean;
 begin
   with D_Eh do
   begin
    UseMultiTitle:=True; //是否使用多行标题行
    TitleLines:=2; //标题行行数
    Flat:=True; //平面显示;False为立体显示
    Columns.Clear;

		Columns[0].Title.Caption:='孔号';
		Columns[1].Title.Caption:='桩号';
		Columns[2].Title.Caption:='排序';
		Columns[3].Title.Caption:='孔序';
		Columns[4].Title.Caption:='孔口高程'+#13+'(m)';
		Columns[5].Title.Caption:='混凝土厚'+#13+'(m)';
		Columns[6].Title.Caption:='灌浆孔段(m)|自'; //标题行文本
		Columns[7].Title.Caption:='灌浆孔段(m)|至'; //标题行文本
		Columns[8].Title.Caption:='灌浆孔段(m)|段长'; //标题行文本
		Columns[9].Title.Caption:='透水率'+#13+'(Lu)'; //标题行文本
		Columns[10].Title.Caption:='水泥用量|注浆'+#13+'(L)'; //标题行文本
		Columns[11].Title.Caption:='水泥用量|注灰'+#13+'(kg)'; //标题行文本
		Columns[12].Title.Caption:='水泥用量|废弃'+#13+'(kg)'; //标题行文本
		Columns[13].Title.Caption:='水泥用量|合计'+#13+'(kg)'; //标题行文本
		Columns[14].Title.Caption:='单位'+#13+'注入量'+#13+'(kg/m)'; //标题行文本
             {
		Columns[0].Width:=40;
		Columns[1].Width:=60;
		Columns[2].Width:=20;
		Columns[3].Width:=20; //孔序
		Columns[4].Width:=50;
		Columns[5].Width:=30;
		Columns[6].Width:=35;
		Columns[7].Width:=40;
		Columns[8].Width:=35;
		Columns[9].Width:=40;    //透水率
		Columns[10].Width:=60;
		Columns[11].Width:=60;
		Columns[12].Width:=40;
		Columns[13].Width:=60;
		Columns[14].Width:=50;
                    }
    Columns[15].Width:=80;
    Columns[1].Alignment:=taLeftJustify;         //单元中内容居中
    Columns[2].Alignment:=taCenter;         //单元中内容居中
    Columns[3].Alignment:=taCenter;         //单元中内容居中

//    Columns[0].ReadOnly:=True;                  //设置表格的部分列不可写
    Columns[2].STFilter.ListSource:=Form1.DS_S;
    Columns[3].STFilter.ListSource:=Form1.DS_S;

    FooterColor := clScrollBar;
    FooterRowCount :=1;
    SumList.Active := true;
    Columns[0].Footer.valuetype := fvtcount;
    Columns[8].Footer.valuetype := fvtsum;
    Columns[10].Footer.valuetype := fvtsum;     //注浆量
    Columns[11].Footer.valuetype := fvtsum;
    Columns[12].Footer.valuetype := fvtsum;
    Columns[13].Footer.valuetype := fvtsum;
    Columns[14].Footer.valuetype := fvtAvg;    //单位注浆量

   end;
    Result:=True;
 end;


 function Data_Eh(D_Eh:TDBGridEh):Boolean;
 begin
   with D_Eh do
   begin
    UseMultiTitle:=True; //是否使用多行标题行
    TitleLines:=2; //标题行行数
    Flat:=True; //平面显示;False为立体显示
    Columns.Clear;
		Columns[0].Title.Caption:='段次';
		Columns[1].Title.Caption:='灌浆孔段(m)|自'; //标题行文本
		Columns[2].Title.Caption:='灌浆孔段(m)|至'; //标题行文本
		Columns[3].Title.Caption:='灌浆孔段(m)|段长'; //标题行文本
		Columns[4].Title.Caption:='孔径'+#13+'(mm)'; //标题行文本
		Columns[5].Title.Caption:='透水率'+#13+'(Lu)'; //标题行文本
		Columns[6].Title.Caption:='水灰比|起始'; //标题行文本
		Columns[7].Title.Caption:='水灰比|终止'; //标题行文本
		Columns[8].Title.Caption:='注入率'+#13+'(L/min)|起始'; //标题行文本
		Columns[9].Title.Caption:='注入率'+#13+'(L/min)|终止'; //标题行文本
		Columns[10].Title.Caption:='水泥用量|注浆'+#13+'(L)'; //标题行文本
		Columns[11].Title.Caption:='水泥用量|注灰'+#13+'(kg)'; //标题行文本
		Columns[12].Title.Caption:='水泥用量|废弃'+#13+'(kg)'; //标题行文本
		Columns[13].Title.Caption:='水泥用量|合计'+#13+'(kg)'; //标题行文本
		Columns[14].Title.Caption:='单位'+#13+'注入量'+#13+'(kg/m)'; //标题行文本
		Columns[15].Title.Caption:='灌浆'+#13+'压力'+#13+'(Mpa)'; //标题行文本
		Columns[16].Title.Caption:='灌浆时间|起始'; //标题行文本
		Columns[17].Title.Caption:='灌浆时间|终止'; //标题行文本
		Columns[18].Title.Caption:='灌浆时间|纯灌'+#13+'(hh:mm)'; //标题行文本
		Columns[19].Title.Caption:='备注'; //标题行文本
                   {
		Columns[0].Width:=20;
		Columns[1].Width:=30;
		Columns[2].Width:=30;
		Columns[3].Width:=30;
		Columns[4].Width:=30;
		Columns[5].Width:=40;
		Columns[6].Width:=35;
		Columns[7].Width:=40;
		Columns[8].Width:=40;
		Columns[9].Width:=40;
		Columns[10].Width:=50;
		Columns[11].Width:=50;
		Columns[12].Width:=35;
		Columns[13].Width:=50;
		Columns[14].Width:=45;
		Columns[15].Width:=40;
		Columns[16].Width:=110;
		Columns[17].Width:=110;
		Columns[18].Width:=40;
		Columns[19].Width:=75;
                     }
    Columns[6].Alignment:=taRightJustify;         //单元中内容居中
    Columns[7].Alignment:=taRightJustify;         //单元中内容居中
    Columns[18].Alignment:=taRightJustify;         //单元中内容居中

//    Columns[0].ReadOnly:=True;                  //设置表格的部分列不可写    
    Columns[0].STFilter.ListSource:=Form1.DS_S;

    //    Columns[0].Width:=40;                   //点号       显示表Dbgrid中的列宽设置
//    Columns[0].Title.Alignment:=taCenter;        //表头文字居中
//    Columns[0].Alignment:=taCenter;         //单元中内容居中
//    Columns[0].Color:=ColorToRGB($ABABAB);         //列前景色
               {
    Columns[1].Footer.ValueType:=fvtStaticText;
    Columns[1].Footer.Value:='最小高程：';
    Columns[4].Footer.ValueType:=fvtMin;
    Columns[3].Footer.ValueType:=fvtStaticText;
    Columns[3].Footer.Value:='最大高程：';
    Columns[4].Footer.ValueType:=fvtmax;
    Columns[6].Footer.ValueType:=fvtStaticText;
    Columns[6].Footer.Value:='总长：';
    Columns[7].Footer.ValueType:=fvtmax;

     DbgEh_S.Columns[0].Footers.Add; // 加入Footer首行

// 设置首行第1列

  DbgEh_S.Columns[0].Footers[0].ValueType:=fvtStaticText; // 显示文本

  DbgEh_S.Columns[0].Footers[0].Value:='合计';

  DbgEh_S.Columns[0].Footers[0].Alignment:=taCenter; // 中心对齐

  DbgEh_S.Columns[0].Footers.Add; // 加入Footer次行

// 设置首行第2列

  DbgEh_S.Columns[0].Footers[1].ValueType:=fvtCount; // 计数

  DbgEh_S.Columns[0].Footers[1].FieldName:='编号'; // 字段名

  DbgEh_S.Columns[0].Footers[1].Alignment:=taCenter; // 中心对齐

  DbgEh_S.Columns[3].Footers.Add; // 加入首行第4列

// 设置首行第4列

  DbgEh_S.Columns[3].Footers[0].ValueType:=fvtSum; // 数据类型：合计

  DbgEh_S.Columns[3].Footers[0].FieldName:='金额'; // 字段名

  DbgEh_S.Columns[3].Footers[0].DisplayFormat:='#,###,###.00'; // 显示格式

  DbgEh_S.Columns[3].Footers.Add; // 加入次行第4列

// 设置次行第4列

  DbgEh_S.Columns[3].Footers[1].ValueType:=fvtFieldValue; // 数据类型：字段值

  DbgEh_S.Columns[3].Footers[1].FieldName:='账号'; // 字段名

  DbgEh_S.Columns[3].Footers[1].Font.Style:=[fsBold]; // 文字格式

  DbgEh_S.Columns[3].Footers[1].Font.Color:=clBlue; // 文字尺寸

  DbgEh_S.SumList.Active:=True; // 确定 统计合计

end;

             }

    FooterColor :=clInactiveCaption;        //clScrollBar;
    FooterRowCount :=1;
    SumList.Active := true;
    Columns[0].Footer.valuetype := fvtcount;
    Columns[3].Footer.valuetype := fvtsum;
    Columns[10].Footer.valuetype := fvtsum;     //注浆量
    Columns[11].Footer.valuetype := fvtsum;
    Columns[12].Footer.valuetype := fvtsum;
    Columns[13].Footer.valuetype := fvtsum;
    Columns[14].Footer.valuetype := fvtAvg;    //单位注浆量
    Columns[15].Footer.valuetype := fvtAvg;
   end;
    Result:=True;
 end;


 function Data_EhA(D_Eh:TDBGridEh):Boolean;
 begin
   with D_Eh do
   begin
    UseMultiTitle:=True; //是否使用多行标题行
    TitleLines:=2; //标题行行数
    Flat:=True; //平面显示;False为立体显示
    Columns.Clear;

		Columns[0].Title.Caption:='孔号';
		Columns[1].Title.Caption:='段次';
		Columns[2].Title.Caption:='灌浆孔段(m)|自'; //标题行文本
		Columns[3].Title.Caption:='灌浆孔段(m)|至'; //标题行文本
		Columns[4].Title.Caption:='灌浆孔段(m)|段长'; //标题行文本
		Columns[5].Title.Caption:='孔径'+#13+'(mm)'; //标题行文本
		Columns[6].Title.Caption:='透水率'+#13+'(Lu)'; //标题行文本
		Columns[7].Title.Caption:='水灰比|起始'; //标题行文本
		Columns[8].Title.Caption:='水灰比|终止'; //标题行文本
		Columns[9].Title.Caption:='注入率'+#13+'(L/min)|起始'; //标题行文本
		Columns[10].Title.Caption:='注入率'+#13+'(L/min)|终止'; //标题行文本
		Columns[11].Title.Caption:='水泥用量|注浆'+#13+'(L)'; //标题行文本
		Columns[12].Title.Caption:='水泥用量|注灰'+#13+'(kg)'; //标题行文本
		Columns[13].Title.Caption:='水泥用量|废弃'+#13+'(kg)'; //标题行文本
		Columns[14].Title.Caption:='水泥用量|合计'+#13+'(kg)'; //标题行文本
		Columns[15].Title.Caption:='单位'+#13+'注入量'+#13+'(kg/m)'; //标题行文本
		Columns[16].Title.Caption:='灌浆'+#13+'压力'+#13+'(Mpa)'; //标题行文本
		Columns[17].Title.Caption:='灌浆时间|起始'; //标题行文本
		Columns[18].Title.Caption:='灌浆时间|终止'; //标题行文本
		Columns[19].Title.Caption:='灌浆时间|纯灌'+#13+'(hh:mm)'; //标题行文本
		Columns[20].Title.Caption:='备注'; //标题行文本
                   {
		Columns[0].Width:=20;
		Columns[1].Width:=30;
		Columns[2].Width:=30;
		Columns[3].Width:=30;
		Columns[4].Width:=30;
		Columns[5].Width:=40;
		Columns[6].Width:=35;
		Columns[7].Width:=40;
		Columns[8].Width:=40;
		Columns[9].Width:=40;
		Columns[10].Width:=50;
		Columns[11].Width:=50;
		Columns[12].Width:=35;
		Columns[13].Width:=50;
		Columns[14].Width:=45;
		Columns[15].Width:=40;
		Columns[16].Width:=110;
		Columns[17].Width:=110;
		Columns[18].Width:=40;
		Columns[19].Width:=75;
                     }
    Columns[7].Alignment:=taRightJustify;         //单元中内容居中
    Columns[8].Alignment:=taRightJustify;         //单元中内容居中
    Columns[19].Alignment:=taRightJustify;         //单元中内容居中

//    Columns[0].ReadOnly:=True;                  //设置表格的部分列不可写
//    Columns[1].ReadOnly:=True;                  //设置表格的部分列不可写
    Columns[0].STFilter.ListSource:=Form1.DS_S;
    Columns[1].STFilter.ListSource:=Form1.DS_S;

    FooterColor := clScrollBar;
    FooterRowCount :=1;
    SumList.Active := true;
    Columns[0].Footer.valuetype := fvtcount;
    Columns[4].Footer.valuetype := fvtsum;
    Columns[11].Footer.valuetype := fvtsum;     //注浆量
    Columns[12].Footer.valuetype := fvtsum;
    Columns[13].Footer.valuetype := fvtsum;
    Columns[14].Footer.valuetype := fvtsum;
    Columns[15].Footer.valuetype := fvtAvg;    //单位注浆量
    Columns[16].Footer.valuetype := fvtAvg;
   end;
    Result:=True;
 end;

procedure TForm1.Button1Click(Sender: TObject);     //读取Excel文件
var I: Integer;
begin
  opendialog1.filter:='Excel(*.xls)|*.xls|Excel2007(*.xlsx)|*.xlsx';
  if opendialog1.Execute then
  begin
    ListBox2.Clear;
  //  files:=opendialog1.Files;     //获得所有文件名
    for i:=0 to opendialog1.Files.count-1 do
    ListBox2.Items.Add(opendialog1.Files[i])//list[i]就是存放所选文件的单个文件名
  end;
  ListBox2.ItemIndex:=0;
  Button4.Enabled:=True;
  Memo1.Clear;
  Memo2.Clear;
end;

procedure TForm1.FormCreate(Sender: TObject);
var i:Variant;
begin
  i:=Trunc(Screen.Height);           //桌面高度
  Form1.Top:=(i-Form1.Height)/2;
  i:=Trunc(Screen.Width);             //桌面宽度
  Form1.Left:=(i-Form1.Width)/2;
  DbChart1.Top:=168;
  DBChart1.Left:=0;   
  Button1.Enabled:=False;
  Button4.Enabled:=False;
  Button5.Enabled:=False;
//  Button6.Enabled:=False;
  ListBox1.Enabled:=False;
  ListBox2.Enabled:=False;
  ListBox2.Width:=392;
  DbChart1.Visible:=false;
  Memo1.Width:=145;
  Memo2.Width:=145;
  Memo1.Height:=80;
  Memo2.Height:=80;

//  Memo1.Visible:=False;
//  Memo2.Visible:=False;
  PopupMenu1.Items[2].Visible:=false;
  PopupMenu1.Items[3].Visible:=false;
  PopupMenu1.Items[4].Visible:=false;
  PopupMenu1.Items[5].Visible:=false;

      //OpenDialog的Options\ofAllowMultiSelect属性设为true，表示可以选择多个文件。
  OpenDialog1.Options:=OpenDialog1.Options+[ofAllowMultiSelect];
  DbgEh_S.Height:=410;
  DbgEh_S.Width:=595;
  DbgEh_S.Anchors:=[akTop,akLeft,akRight,akBottom];
  DbgEh_S.ReadOnly:=True;
  DbgEh_S.Color:=cl3DLight;
  DbgEh_S.PopupMenu:=PopupMenu1;
  PrintDBGridEh2.DBGridEh:=DbgEh_S;
  C_PL:='TPointSeries';
end;

function Data_Ado(ADO:TADOQuery):Boolean;
begin
	TFloatField(ADO.FieldByName('灌浆起深')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('灌浆终深')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('灌浆段长')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('透水率')).DisplayFormat:='#.##';
	TFloatField(ADO.FieldByName('注入率起')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('注入率终')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('水泥注浆量')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('水泥注灰量')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('水泥废弃量')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('水泥合计')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('单位注入量')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('灌浆压力')).DisplayFormat:='0.##';
//  TTimeField(ADO.FieldByName('时长')).DisplayFormat:='hh:mm';
 Result:=True;
end;

function Data_AdoR(ADO:TADOQuery):Boolean;
begin
 TFloatField(ADO.FieldByName('孔口高程')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('混凝土厚')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('灌浆起深')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('灌浆终深')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('灌浆段长')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('平均透水率')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('水泥注浆量')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('水泥注灰量')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('水泥废弃量')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('水泥合计')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('单位注入量')).DisplayFormat:='0.##';
 Result:=True;
end;

procedure TForm1.StatusBar1Click(Sender: TObject);
begin
  ShowMessage(StatusBar1.Panels[0].Text+#13+StatusBar1.Panels[1].Text+#13+StatusBar1.Panels[2].Text);
end;

function DATA_rre(D_Ado:TADOQuery;D_Dbge:TDBGridEh;D_Ds:TDataSource;D_Tab,data_Name:String):Boolean;   //刷新数据
 begin
   with D_Ado do
   begin
    close;
    sql.Clear;
    SQL.Add('select '+D_Tab+'='''+data_Name+'''');
    SQL.Add('order by ID');              //以ID号排序
    open;
   end;

   with D_Dbge do
   begin
    DataSource:=D_Ds;
    AutoFitColWidths:=True;                               //自适应单元格宽度
    ColumnDefValues.Title.TitleButton:=True;             //所有列都参与排序
// 		ColumnDefValues.Title.SortMarker:=smDownEh;         //指定排序标志
    OptionsEh:=OptionsEh+[dghAutoSortMarking];            //采用自动排序功能
    Options:=Options+[dgEditing]+[dgMultiSelect];   //  不要加入[dgRowSelect]，否则不能编辑表格数据
    SortLocal:=True;                                      //客户端排序
    STFilter.Visible := True;
    STFilter.Local := True;
   end;
   Result:=True;
 end;

 function DATA_re(D_Ado:TADOQuery;D_Dbge:TDBGridEh;D_Ds:TDataSource;D_Tab:String):Boolean;   //刷新数据
 begin
   with D_Ado do
   begin
    close;
    sql.Clear;
    SQL.Add('select '+D_Tab);
    SQL.Add('order by ID');              //以ID号排序
    open;
   end;

   with D_Dbge do
   begin
    DataSource:=D_Ds;
    AutoFitColWidths:=True;                               //自适应单元格宽度
    ColumnDefValues.Title.TitleButton:=True;             //所有列都参与排序
// 		ColumnDefValues.Title.SortMarker:=smDownEh;         //指定排序标志
    OptionsEh:=OptionsEh+[dghAutoSortMarking];            //采用自动排序功能
    Options:=Options+[dgEditing]+[dgMultiSelect];   //  不要加入[dgRowSelect]，否则不能编辑表格数据
    SortLocal:=True;                                      //客户端排序
    STFilter.Visible := True;
    STFilter.Local := True;
   end;
   Result:=True;
 end;

 function DATA_r(D_Ado:TADOQuery;D_Dbge:TDBGridEh;D_Ds:TDataSource;D_Tab:String):Boolean;   //刷新数据
 begin
   with D_Ado do
   begin
    close;
    sql.Clear;
    SQL.Add('select '+D_Tab);
    SQL.Add('order by ID');              //以ID号排序
    open;
   end;

   with D_Dbge do
   begin
    DataSource:=D_Ds;
    AutoFitColWidths:=True;                               //自适应单元格宽度
    ColumnDefValues.Title.TitleButton:=True;             //所有列都参与排序
// 		ColumnDefValues.Title.SortMarker:=smDownEh;         //指定排序标志
    OptionsEh:=OptionsEh+[dghAutoSortMarking];            //采用自动排序功能
    Options:=Options+[dgEditing]+[dgMultiSelect];   //  不要加入[dgRowSelect]，否则不能编辑表格数据
    SortLocal:=True;                                      //客户端排序
    STFilter.Visible := True;
    STFilter.Local := True;
   end;
   Result:=True;
 end;


procedure TForm1.ListBox1Click(Sender: TObject);
var
  Str_Q:string;
begin
 { DbChart1.Visible:=false;
  DbgEh_S.Height:=Form1.Height-100;
  DbgEh_S.Anchors:=[akLeft,akTop,akRight,akBottom];
  PopupMenu1.Items[1].Visible:=False;
  PopupMenu1.Items[2].Visible:=False;
  PopupMenu1.Items[3].Visible:=False;
         }

  PopupMenu1.Items[2].Visible:=True;
  PopupMenu1.Items[3].Visible:=True;
  PopupMenu1.Items[4].Visible:=True;
  PopupMenu1.Items[5].Visible:=True;

  Str_Q:='段次,灌浆起深,灌浆终深,灌浆段长,孔径,透水率,水灰比起,水灰比终,注入率起,注入率终,水泥注浆量,水泥注灰量,水泥废弃量,水泥合计,单位注入量,灌浆压力,时间起,时间终,时长,备注 from Caran_GJL where 孔号';
  DATA_rre(Qur_S,DbgEh_S,DS_S,Str_Q,ListBox1.Items[ListBox1.itemindex]);    //刷新数据
  Data_Ado(Qur_S);         //设置数据显示格式
  Data_Eh(DbgEh_S);         //设置数据显示格式
  StatusBar1.Panels[1].Text:=ListBox1.Items[ListBox1.itemindex]+' 孔详细数据';
end;

procedure TForm1.ListBox2DblClick(Sender: TObject);
var  I,N: Integer;//记录数据表的当前记录号
  F: TextFile;  //TextFile 和 Text 是一样的
  Str: string;
begin
  try                                            //链接到Excel
    E_App := CreateOleObject('Excel.Application');
    E_App.Visible := true;                       //显示打开的Excel文件
   if FileExists(ListBox2.Items[ListBox2.itemindex]) then
    begin
     E_App.WorkBooks.Open(ListBox2.Items[ListBox2.itemindex]);   //打开文件

     AssignFile(F,StatusBar1.Panels[2].Text);
     Append(F);  //打开准备追加
     for i:=1 to E_App.Sheets.Count do
     begin
       E_Sheet:=E_App.worksheets[i];
       E_Sheet.activate;
       if (Trim(E_Sheet.cells.item[7,2])='') and (Trim(E_Sheet.cells.item[7,3])='') then continue;
       if E_Sheet.Name<>'合计' then Ex_Tro(E_Sheet.Name);      //读取Excel表
       N:=E_Sheet.usedrange.rows.count;
       repeat N := N - 1 until Trim(E_Sheet.cells.item[N,1])<>'';

       Str:=E_Sheet.Name+#9+E_Sheet.cells[4,1].value+#9+FloatToStr(E_Sheet.cells[N,2].value)+#9+FloatToStr(E_Sheet.cells[N,3].value)+#9+FloatToStr(E_Sheet.cells[N,4].value)+#9+FloatToStr(E_Sheet.cells[N,6].value)+#9;
       Str:=Str+FloatToStr(E_Sheet.cells[N,11].value)+#9+FloatToStr(E_Sheet.cells[N,12].value)+#9+FloatToStr(E_Sheet.cells[N,13].value)+#9+FloatToStr(E_Sheet.cells[N,14].value)+#9+FloatToStr(E_Sheet.cells[N,15].value);
       Writeln(F, Str);
     end;

     CloseFile(F);

     E_App.WorkBooks.Close; //关闭工作簿
     E_App.Quit; //退出 Excel
     E_App:=Unassigned;//释放excel进程
     DATA_Connect(StatusBar1.Panels[0].Text,'Caran_GJL',True);           //链接数据库
     ShowMessage(ListBox2.Items[ListBox2.itemindex]+#13+'导入完成');
     StatusBar1.Panels[1].Text:=ListBox2.Items[ListBox2.Itemindex];
    end
  except ShowMessage('Excel文件调用出错或不存在！');
	end;
end;

function D_Create(D_ADO:TADOQuery;T_Name,T_Str:string):Boolean;
begin
  try
  with D_ADO do
  begin
    Close;
    SQL.Clear;
    Active:=false;                                          //创建数据库中的表     公司信息
    SQL.Add('create table '+T_Name+' (ID AUTOINCREMENT,'+T_Str+'primary key(ID))');
    ExecSQL;
  end;
 except
 end;
 Result:=True;
end;

function DATA_Create(Data_Name:string):Boolean;    //创建数据库
 var CreateAccess:OleVariant;
 Sql_Data:string;
begin
  CreateAccess:=CreateOleObject('ADOX.Catalog');
  CreateAccess.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_name+'.MDB');

  Form1.ADOC_S.Connected:=False;       //设置数据库联结为False，关闭联结，以便于下一步联结
  Form1.ADOC_S.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_name+'.mdb;Persist Security Info=False';
//  ADOC_S.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_bame+'.mdb;Persist Security Info=False;Jet OLEDB:Database Password=$#！ailihongbi 80！^%&';
  Form1.ADOC_S.LoginPrompt:=false;        //不显示登录框
  Form1.ADOC_S.Connected:=true;
  Form1.Qur_S.Connection:=Form1.ADOC_S;
                                          //创建数据库中的表
  Sql_Data:='孔号 varchar(8),段次 Byte,灌浆起深 Single,灌浆终深 Single,灌浆段长 Single,孔径 Byte,透水率 Single,水灰比起 varchar(8),水灰比终 varchar(8),注入率起 Single,注入率终 Single,';
  sql_Data:=Sql_Data+'水泥注浆量 Single,水泥注灰量 Single,水泥废弃量 Single,水泥合计 Single,单位注入量 Single,灌浆压力 Single,时间起 Date,时间终 Date,时长 Date,备注 varchar(20),';
  D_Create(Form1.Qur_S,'Caran_GJL',Sql_Data);

  Sql_Data:='孔号 varchar(8),段次 Byte,灌浆起深 Single,灌浆终深 Single,灌浆段长 Single,孔径 Byte,透水率 Single,水灰比起 varchar(8),水灰比终 varchar(8),注入率起 Single,注入率终 Single,';
  sql_Data:=Sql_Data+'水泥注浆量 Single,水泥注灰量 Single,水泥废弃量 Single,水泥合计 Single,单位注入量 Single,灌浆压力 Single,时间起 Date,时间终 Date,时长 Date,备注 varchar(20),';
  D_Create(Form1.Qur_S,'Caran_GJT',Sql_Data);

  Sql_Data:='孔号 varchar(8),桩号 varchar(12),排序 varchar(12),孔序 Byte,孔口高程 Single,混凝土厚 Single,灌浆起深 Single,灌浆终深 Single,灌浆段长 Single,平均透水率 Single,';
  sql_Data:=Sql_Data+'水泥注浆量 Single,水泥注灰量 Single,水泥废弃量 Single,水泥合计 Single,单位注入量 Single,备注 varchar(50),';
  D_Create(Form1.Qur_S,'Caran_T',Sql_Data);

  Result:=True;
end;


procedure TForm1.Button4Click(Sender: TObject);
var
  maxR,J,K,N,M: Integer;//记录数据表的当前记录号
  F_Name,Str:String;
  F:TextFile;
  List,Del_List:TStringList;
begin
 List:=TStringList.Create();
 Del_List:=TStringList.Create();
   for j:=0 to ListBox2.Count-1 do
   if FileExists(ListBox2.Items[j]) then
    try
     E_App := CreateOleObject('Excel.Application');
     E_App.Visible := true;                       //显示打开的Excel文件
     E_App.WorkBooks.Open(ListBox2.Items[j]);   //打开文件
     E_Sheet:=E_App.worksheets[1];
     E_Sheet.activate;
     F_Name:=Copy(StatusBar1.Panels[0].Text,1,Length(StatusBar1.Panels[0].Text)-4);
     AssignFile(F,F_Name+'.TXT');
     if not FileExists(F_Name+'.TXT') then Rewrite(F) else Append(F);   //打开存在的文件并把文件指针定位在文件尾。

          maxR:= E_Sheet.usedrange.rows.count;     //有数据的行数
          Str:='insert into Caran_GJL (孔号,段次,灌浆起深,灌浆终深,灌浆段长,孔径,透水率,水灰比起,水灰比终,注入率起,注入率终,水泥注浆量,水泥注灰量,水泥废弃量,水泥合计,单位注入量,灌浆压力,时间起,时间终,时长,备注) values';
          Str:=Str+'(:I_D,:D_i,:D_S,:D_E,:D_L,:D_R,:K_L,:K_s,:K_e,:Z_S,:Z_E,:V_a,:V_b,:V_c,:V_v,:V_m,:V_p,:T_S,:T_E,:T_T,:K_B)';

         for k:=7 to maxR do
         begin
          if not TryStrToInt(Trim(E_Sheet.cells.item[k,1]),maxR) then
          begin
           continue;                     //如果段次不是数字则跳出本次循环
          end;
          if (Trim(E_Sheet.cells.item[k,2])<>'') and (Trim(E_Sheet.cells.item[k,3])<>'') then       //如果本行不为空则运行
          try
           with Qur_S do
           begin
            Close;
            SQL.Clear;
            SQL.Add(Str);
        		Parameters.ParamByName('I_D').Value:=E_Sheet.Name;		//孔号
         		try Parameters.ParamByName('D_i').Value:=StrToInt(E_Sheet.cells.item[k,1]);  //E_Sheet.cells[i,1].value;		//段次
            except Del_List.Add(E_Sheet.Name+#9+'段次'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('D_S').Value:=E_Sheet.cells[k,2].value;		//灌浆起深
            except Del_List.Add(E_Sheet.Name+#9+'起深'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('D_E').Value:=E_Sheet.cells[k,3].value;		//灌浆终深
            except Del_List.Add(E_Sheet.Name+#9+'终深'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('D_L').Value:=E_Sheet.cells[k,4].value;		//灌浆段长
            except Del_List.Add(E_Sheet.Name+#9+'段长'+#9+inttostr(k)) end;
            try Parameters.ParamByName('D_R').Value:=E_Sheet.cells[k,5].value;		//孔径
            except Del_List.Add(E_Sheet.Name+#9+'孔径'+#9+inttostr(k)) end;
            try if Trim(E_Sheet.cells.item[k,6])='/' then Parameters.ParamByName('K_L').Value:=0 else Parameters.ParamByName('K_L').Value:=E_Sheet.cells[k,6].value;		//透水率
            except Del_List.Add(E_Sheet.Name+#9+'透水率'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('K_s').Value:=E_Sheet.cells[k,7].value;		//水灰比起
            except Del_List.Add(E_Sheet.Name+#9+'水灰比起'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('K_e').Value:=E_Sheet.cells[k,8].value;		//水灰比终
            except Del_List.Add(E_Sheet.Name+#9+'水灰比终'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('Z_S').Value:=E_Sheet.cells[k,9].value;		//注入率起
            except Del_List.Add(E_Sheet.Name+#9+'注入率起'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('Z_E').Value:=E_Sheet.cells[k,10].value;		//注入率终
            except Del_List.Add(E_Sheet.Name+#9+'注入率终'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_a').Value:=E_Sheet.cells[k,11].value;		//水泥量浆
            except Del_List.Add(E_Sheet.Name+#9+'浆量'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_b').Value:=E_Sheet.cells[k,12].value;		//水泥量灰
            except Del_List.Add(E_Sheet.Name+#9+'灰量'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_c').Value:=E_Sheet.cells[k,13].value;		//水泥量废
            except Del_List.Add(E_Sheet.Name+#9+'废灰'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_v').Value:=E_Sheet.cells[k,14].value;		//水泥量合
            except Del_List.Add(E_Sheet.Name+#9+'合计灰'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_m').Value:=E_Sheet.cells[k,15].value;		//单位注入量
            except Del_List.Add(E_Sheet.Name+#9+'单位注入量'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_p').Value:=E_Sheet.cells[k,16].value;		//灌浆压力
            except Del_List.Add(E_Sheet.Name+#9+'压力'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('T_S').Value:=E_Sheet.cells[k,17].value;		//时间起
            except Del_List.Add(E_Sheet.Name+#9+'起时'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('T_E').Value:=E_Sheet.cells[k,18].value;	 	//时间终
            except Del_List.Add(E_Sheet.Name+#9+'终时'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('T_T').Value:=E_Sheet.cells[k,19].value;	 	//S 时长
            except Del_List.Add(E_Sheet.Name+#9+'时长'+#9+inttostr(k)) end;
        		Parameters.ParamByName('K_B').Value:=E_Sheet.cells[k,20].value;		//T备注
            ExecSQL;                    //打开数据集，执行SQL语句
           end;
          except
            Del_List.Add(E_Sheet.Name+#9+inttostr(k));
            continue;
          end;
         end;

       N:=E_Sheet.usedrange.rows.count;
       repeat N := N - 1 until Trim(E_Sheet.cells.item[N,1])<>'';
       List.Add(E_Sheet.Name);
       Str:=E_Sheet.Name+#9+E_Sheet.cells[4,1].value+#9+FloatToStr(E_Sheet.cells[N,2].value)+#9+FloatToStr(E_Sheet.cells[N,3].value)+#9+FloatToStr(E_Sheet.cells[N,4].value)+#9+FloatToStr(E_Sheet.cells[N,6].value)+#9;
      try
       Str:=Str+FloatToStr(E_Sheet.cells[N,11].value)+#9+FloatToStr(E_Sheet.cells[N,12].value)+#9+FloatToStr(E_Sheet.cells[N,13].value)+#9+FloatToStr(E_Sheet.cells[N,14].value)+#9+FloatToStr(E_Sheet.cells[N,15].value);
      except
      end;
     Writeln(F, Str);     
     CloseFile(F);
     E_App.WorkBooks.Close;         //关闭工作簿
     E_App.Quit;                    //退出 Excel
     E_App:=Unassigned;             //释放excel进程
     DATA_Connect(StatusBar1.Panels[0].Text,'Caran_GJL',True);           //链接数据库
     StatusBar1.Panels[1].Text:=ListBox2.Items[j];
    except;
   	end;
  ListBox2.Width:=Form1.Width-320;
  Memo1.Lines:=List;
  Memo2.Lines:=Del_List;
  Button4.Enabled:=False;
//  Memo1.Visible:=True;
//  Memo2.Visible:=True;
  ShowMessage(IntToStr(List.Count)+'个文件全部完成！');
  List.Free;
  Del_List.Free;

end;

procedure TForm1.ListBox2Click(Sender: TObject);
begin
  StatusBar1.Panels[1].Text:=ListBox2.Items[ListBox2.Itemindex];
end;

procedure TForm1.Button2Click(Sender: TObject);
var F:TextFile;
    F_Name:string;
begin
 opendialog1.filter:='Access(*.Mdb)|*.MDB';
 if opendialog1.Execute then
 try
   DATA_Connect(OpenDialog1.FileName,'Caran_GJL',True);           //链接数据库
   Data_EhH(DbgEh_S);         //设置数据显示格式

   F_Name:=Copy(OpenDialog1.FileName,1,Length(OpenDialog1.FileName)-3)+'TXT';
   AssignFile(F,F_Name);
//   try Append(f) except Rewrite(F) end;  //新建文件，如果已存在则追加,否则新建
   if not FileExists(F_Name) then Rewrite(F) else Append(F);   //打开存在的文件并把文件指针定位在文件尾。
   CloseFile(F);
   TT:='Synopsis';
   Button1.Enabled:=True;
 //  Button4.Enabled:=True;
   Button5.Enabled:=True;
   ListBox1.Enabled:=True;
   ListBox2.Enabled:=True;
   StatusBar1.Panels[2].Text:=F_Name;
   StatusBar1.Panels[0].Text:=OpenDialog1.FileName;  
 except ShowMessage('打开数据库失败!');
 end;
end;

procedure TForm1.Button3Click(Sender: TObject);
var F:TextFile;
begin
 Savedialog1.filter:='Access(*.Mdb)|*.MDB';
 if SaveDialog1.Execute then
 try
   DATA_Create(SaveDialog1.FileName);

   AssignFile(F,Trim(SaveDialog1.FileName)+'.TXT');
//   try Append(f) except Rewrite(F) end;  //新建文件，如果已存在则追加,否则新建
   if not FileExists(Trim(SaveDialog1.FileName)+'.TXT') then Rewrite(F) else Append(F);   //打开存在的文件并把文件指针定位在文件尾。
   CloseFile(F);
   TT:='Synopsis';
   Button1.Enabled:=True;
   Button4.Enabled:=True;
   Button5.Enabled:=True;
   ListBox1.Enabled:=True;
   ListBox2.Enabled:=True;
   StatusBar1.Panels[0].Text:=Trim(SaveDialog1.FileName)+'.Mdb';     //获取保存路径
   StatusBar1.Panels[2].Text:=Trim(SaveDialog1.FileName)+'.TXT';
 except ShowMessage('新建数据库失败!');
 end;
end;

procedure TForm1.Button5Click(Sender: TObject);
var Str_Q:string;
begin
 if TT='Synopsis' then           //简况
 begin
  DbChart1.Visible:=false;
  DbgEh_S.Height:=Form1.Height-115;
  DbgEh_S.Anchors:=[akLeft,akTop,akRight,akBottom];
          {
  PopupMenu1.Items[2].Visible:=False;
  PopupMenu1.Items[3].Visible:=False;
  PopupMenu1.Items[4].Visible:=False;
  PopupMenu1.Items[5].Visible:=False;
           }
  PopupMenu1.Items[2].Visible:=True;
  PopupMenu1.Items[3].Visible:=True;
  PopupMenu1.Items[4].Visible:=True;
  PopupMenu1.Items[5].Visible:=True;

  Str_Q:='孔号,桩号,排序,孔序,孔口高程,混凝土厚,灌浆起深,灌浆终深,灌浆段长,平均透水率,水泥注浆量,水泥注灰量,水泥废弃量,水泥合计,单位注入量,备注 from Caran_T where 孔号';
  DATA_re(Qur_S,DbgEh_S,DS_S,Str_Q);    //刷新数据
  Data_AdoR(Qur_S);         //设置数据显示格式
  Data_EhR(DbgEh_S);         //设置数据显示格式
  TT:='Statistics';
  StatusBar1.Panels[1].Text:='灌浆孔简况';
 end else if TT='Statistics' then          //统计
 begin
  Str_Q:='孔号,段次,灌浆起深,灌浆终深,灌浆段长,孔径,透水率,水灰比起,水灰比终,注入率起,注入率终,水泥注浆量,水泥注灰量,水泥废弃量,水泥合计,单位注入量,灌浆压力,时间起,时间终,时长,备注 from Caran_GJT where 孔号';
  DATA_re(Qur_S,DbgEh_S,DS_S,Str_Q);    //刷新数据
  Data_Ado(Qur_S);         //设置数据显示格式
  Data_EhA(DbgEh_S);         //设置数据显示格式
  StatusBar1.Panels[1].Text:='灌浆孔统计数据';

  PopupMenu1.Items[2].Visible:=True;
  PopupMenu1.Items[3].Visible:=True;
  PopupMenu1.Items[4].Visible:=True;
  PopupMenu1.Items[5].Visible:=True;
  TT:='ALL';
 end else if TT='ALL' then
 begin
  Str_Q:='孔号,段次,灌浆起深,灌浆终深,灌浆段长,孔径,透水率,水灰比起,水灰比终,注入率起,注入率终,水泥注浆量,水泥注灰量,水泥废弃量,水泥合计,单位注入量,灌浆压力,时间起,时间终,时长,备注 from Caran_GJL where 孔号';
  DATA_re(Qur_S,DbgEh_S,DS_S,Str_Q);    //刷新数据
  Data_Ado(Qur_S);         //设置数据显示格式
  Data_EhA(DbgEh_S);         //设置数据显示格式
  StatusBar1.Panels[1].Text:='灌浆孔详细数据';
  PopupMenu1.Items[2].Visible:=True;
  PopupMenu1.Items[3].Visible:=True;
  PopupMenu1.Items[4].Visible:=True;
  PopupMenu1.Items[5].Visible:=True;
  TT:='Synopsis';
 end;
end;

procedure TForm1.N1Click(Sender: TObject);       //打印表格
begin
//  PrintDBGridEh2.Title.Text:='明细表';

  PrintDBGridEh2.PageHeader.CenterText.Clear;
  PrintDBGridEh2.PageHeader.CenterText.Add(trim(StatusBar1.Panels[1].Text));    //
  PrintDBGridEh2.PageHeader.Font.Style:=[fsBold];
  PrintDBGridEh2.PageHeader.Font.Name:='黑体';
  PrintDBGridEh2.PageHeader.Font.Size:=12;

  PrintDBGridEh2.PageFooter.CenterText.Clear;
  PrintDBGridEh2.PageFooter.CenterText.Add('第 &[Page] 页 / 共 &[Pages] 页');
  PrintDBGridEh2.PageFooter.RightText.Add(SysUtils.DateTimeToStr(Now()));
  PrintDBGridEh2.PageFooter.Font.Size:=7;
  PrintDBGridEh2.Preview; //打印预览
  //PrintDBGridEh2.Print; //直接输送到打印机上打印
end;

procedure TForm1.XT1Click(Sender: TObject);
var
  ExpClass: TDBGridEhExportclass;
  Ext: string;
  FSaveDialog: TSaveDialog;
begin
  try
    if DbgEh_S.DataSource.DataSet.IsEmpty then
    begin
      Application.MessageBox(PChar('没有可导出的数据'), PChar('提示'), MB_OK + MB_ICONINFORMATION);
      exit;
    end;
    FSaveDialog := TSaveDialog.Create(Self);
    FSaveDialog.Filter :=
      'Excel 文档 (*.xls)|*.XLS|Text files (*.txt)|*.TXT|Comma separated values (*.csv)|*.CSV|HTML file (*.htm)|*.HTM|Word 文档 (*.rtf)|*.RTF';
    if FSaveDialog.Execute and (trim(FSaveDialog.FileName) <> '') then
    begin
      case FSaveDialog.FilterIndex of
        1:
          begin
            ExpClass := TDBGridEhExportAsXLS;
            Ext := 'xls';
          end;
        2:
          begin
            ExpClass := TDBGridEhExportAsText;
            Ext := 'txt';
          end;
        3:
          begin
            ExpClass := TDBGridEhExportAsCSV;
            Ext := 'csv';
          end;
        4:
          begin
            ExpClass := TDBGridEhExportAsHTML;
            Ext := 'htm';
          end;
        5:
          begin
            ExpClass := TDBGridEhExportAsRTF;
            Ext := 'rtf';
          end;
      end;
      if ExpClass <> nil then
      begin
        if UpperCase(Copy(FSaveDialog.FileName, Length(FSaveDialog.FileName) -
          2, 3)) <> UpperCase(Ext) then
          FSaveDialog.FileName := FSaveDialog.FileName + '.' + Ext;
        if FileExists(FSaveDialog.FileName) then
        begin
          if application.MessageBox('文件名已存在，是否覆盖   ', '提示',
            MB_ICONASTERISK or MB_OKCANCEL) <> idok then
            exit;
        end;
        Screen.Cursor := crHourGlass;
        SaveDBGridEhToExportFile(ExpClass, DbgEh_S, FSaveDialog.FileName, true);
        Screen.Cursor := crDefault;
        MessageBox(Handle, '导出成功  ', '提示', MB_OK +
          MB_ICONINFORMATION);
      end;
    end;
    FSaveDialog.Destroy;
  except
    on e: exception do
    begin
      Application.MessageBox(PChar(e.message), '错误', MB_OK + MB_ICONSTOP);
    end;
  end;
end;

procedure TForm1.N2Click(Sender: TObject);
begin
  DbgEh_S.Anchors:=[akLeft,akTop,akRight];
  DbgEh_S.Height:=170;
  DbChart1.Visible:=True;
  DBChart1.Series[0].DataSource:=Qur_S;

  DbChart1.Title.Text.Clear;

  DBChart1.Title.Text.Add('压水深度(m)----吕荣值(Lu)');
  DBChart1.Series[0].XValues.ValueSource:='灌浆终深';
  if StatusBar1.Panels[1].Text='灌浆孔简况' then
    DBChart1.Series[0].YValues.ValueSource:='平均透水率'
  else DBChart1.Series[0].YValues.ValueSource:='透水率';
//  DbChart1.Series[0].XLabelsSource:='透水率';       //设置X轴
  DbChart1.LeftAxis.Title.Caption:='吕荣值 Lu';
  DBChart1.RefreshDataSet(Qur_S,DBChart1.Series[0]);
end;

procedure TForm1.N3Click(Sender: TObject);
begin
  DbgEh_S.Anchors:=[akLeft,akTop,akRight];
  DbgEh_S.Height:=170;
  DbChart1.Visible:=True;
  DBChart1.Series[0].DataSource:=Qur_S;

  DbChart1.Title.Text.Clear;
  DBChart1.Title.Text.Add('压水深度(m)----单位灌浆量(kg/m)');
  DBChart1.Series[0].XValues.ValueSource:='灌浆终深';
  DBChart1.Series[0].YValues.ValueSource:='单位注入量';
//  DbChart1.Series[0].XLabelsSource:='单位注入量';
  DbChart1.LeftAxis.Title.Caption:='单位注入量 kg/m';
  DBChart1.RefreshDataSet(Qur_S,DBChart1.Series[0]);
end;

procedure TForm1.N4Click(Sender: TObject);
begin
  DbgEh_S.Anchors:=[akLeft,akTop,akRight];
  DbgEh_S.Height:=170;
  DbChart1.Visible:=True;
  DBChart1.Series[0].DataSource:=Qur_S;

  DbChart1.Title.Text.Clear;
  DBChart1.Title.Text.Add('吕荣值(Lu)----单位灌浆量(kg/m)');
  if StatusBar1.Panels[1].Text='灌浆孔简况' then
    DBChart1.Series[0].YValues.ValueSource:='平均透水率'
  else  DBChart1.Series[0].XValues.ValueSource:='透水率';
  
 DBChart1.Series[0].YValues.ValueSource:='单位注入量';
//  DbChart1.Series[0].XLabelsSource:='单位注入量';
  DbChart1.LeftAxis.Title.Caption:='单位注入量 kg/m';
  DBChart1.RefreshDataSet(Qur_S,DBChart1.Series[0]);
end;

procedure TForm1.N5Click(Sender: TObject);
begin
  DbChart1.Visible:=false;
  DbgEh_S.Height:=Form1.Height-115;
  DbgEh_S.Anchors:=[akLeft,akTop,akRight,akBottom];
//  PopupMenu1.Items[2].Visible:=False;
end;

procedure TForm1.N6Click(Sender: TObject);
begin
  DbgEh_S.ReadOnly:=False;
end;

procedure TForm1.DBChart1DblClick(Sender: TObject);
var
  F_PChart:TChartSeries;
begin
  F_PChart:=DBChart1.Series[0];

  if C_PL='TPointSeries' then
   begin
    ChangeSeriesType(F_PChart,TFastLineSeries);
    C_PL:='TFastLineSeries'
   end else
   begin
    ChangeSeriesType(F_PChart,TPointSeries);
    C_PL:='TPointSeries'
   end;

  F_PChart.Active:=True;
  DBChart1.RefreshDataSet(Qur_S,F_PChart);
 // F_PChart.RefreshSeries;
end;

procedure TForm1.DbgEh_SDblClick(Sender: TObject);
var
  F_PChart:TChartSeries;
begin
  if DbChart1.Visible and DbgEh_S.ReadOnly then
  begin
   F_PChart:=DBChart1.Series[0];
   F_PChart.Active:=True;
   DBChart1.RefreshDataSet(Qur_S,F_PChart);
 // F_PChart.RefreshSeries;
  end;
end;

procedure TForm1.Button6Click(Sender: TObject);
var maxR,j,i,k: Integer;//记录数据表的当前记录号
  E_Target,E_TTarg: OleVariant;
begin
  opendialog1.filter:='Excel(*.xls)|*.xls|Excel2007(*.xlsx)|*.xlsx';
 if opendialog1.Execute then
 try
   E_App := CreateOleObject('Excel.Application');
   E_App.Visible := true;                       //显示打开的Excel文件
   E_App.WorkBooks.Open(OpenDialog1.FileName);   //打开文件
   E_Sheet:=E_App.worksheets['Caran_GJL'];
   E_Target:=E_App.worksheets['Caran_GJT'];
   E_TTarg:=E_App.worksheets['Caran_T'];
//   E_Sheet.activate;
   if E_Target.usedrange.rows.count=1 then
    maxR:= E_Sheet.usedrange.rows.count     //有数据的行数
   else maxR:= E_Target.usedrange.rows.count;
 except
 end;
 j:=1;
 k:=1;
 if E_Target.usedrange.rows.count=1 then
 for i:=1 to maxR-1 do
 begin                    //StrToInt(E_Sheet.cells.item[k,1]);
  StatusBar1.Panels[1].Text:=IntToStr(i);
  E_Target.activate;
//  if E_Sheet.cells[i,23].value<>E_Sheet.cells[i+1,23].value then
  if E_Sheet.cells[i,2].value<>E_Sheet.cells[i+1,2].value then
   begin
     E_Target.cells[j,1].value:=E_Sheet.cells[i,1].value;		//ID
     E_Target.cells[j,2].value:=E_Sheet.cells[i,2].value;		//孔号
     E_Target.cells[j,3].value:=E_Sheet.cells[i,3].value;		//段次
     E_Target.cells[j,4].value:=E_Sheet.cells[i,4].value;		//灌浆起深
     E_Target.cells[j,5].value:=E_Sheet.cells[i,5].value;		//灌浆终深
     E_Target.cells[j,6].value:=E_Sheet.cells[i,6].value;		//灌浆段长
     E_Target.cells[j,7].value:=E_Sheet.cells[i,7].value;		//孔径
     try if strtofloat(E_Target.cells[j,8].value)=0 then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value    	//透水率
     except if Trim(E_Target.cells[j,8].value)='' then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value end;
     try if strtofloat(E_Target.cells[j,9].value)=0 then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value     //水灰比起
     except if Trim(E_Target.cells[j,9].value)='' then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value end;
     E_Target.cells[j,10].value:=E_Sheet.cells[i,10].value;		//水灰比终
     try if strtofloat(E_Target.cells[j,11].value)=0 then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value    //注入率起
     except if Trim(E_Target.cells[j,11].value)='' then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value end;
     E_Target.cells[j,12].value:=E_Sheet.cells[i,12].value;		//注入率终
     try if strtofloat(E_Target.cells[j,13].value)=0 then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value		//水泥量浆
     except if Trim(E_Target.cells[j,13].value)='' then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value end;		//水泥量灰
     try if strtofloat(E_Target.cells[j,14].value)=0 then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value		//水泥量废
     except if Trim(E_Target.cells[j,14].value)='' then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value end;		//水泥量合
     try if strtofloat(E_Target.cells[j,15].value)=0 then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value		//单位注入量
     except if Trim(E_Target.cells[j,15].value)='' then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value end;		//水泥量浆
     try if strtofloat(E_Target.cells[j,16].value)=0 then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value		//水泥量灰
     except if Trim(E_Target.cells[j,16].value)='' then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value end;		//水泥量废
     try if strtofloat(E_Target.cells[j,17].value)=0 then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value		//水泥量合
     except if Trim(E_Target.cells[j,17].value)='' then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value end;		//单位注入量
     E_Target.cells[j,18].value:=E_Sheet.cells[i,18].value;		//灌浆压力
     try if strtofloat(E_Target.cells[j,19].value)=0 then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value    	//时间起
     except if Trim(E_Target.cells[j,19].value)='' then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value end;
     E_Target.cells[j,20].value:=E_Sheet.cells[i,20].value;		//时间终
     try if strtofloat(E_Target.cells[j,21].value)=0 then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value   	//时长
     except if Trim(E_Target.cells[j,21].value)='' then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value end;
     E_Target.cells[j,22].value:=E_Sheet.cells[i,22].value;		//备注
    j:=j+1;
   end else
  begin if E_Sheet.cells[i,3].value<>E_Sheet.cells[i+1,3].value then
    begin
     E_Target.cells[j,1].value:=E_Sheet.cells[i,1].value;		//ID
     E_Target.cells[j,2].value:=E_Sheet.cells[i,2].value;		//孔号
     E_Target.cells[j,3].value:=E_Sheet.cells[i,3].value;		//段次
     E_Target.cells[j,4].value:=E_Sheet.cells[i,4].value;		//灌浆起深
     E_Target.cells[j,5].value:=E_Sheet.cells[i,5].value;		//灌浆终深
     E_Target.cells[j,6].value:=E_Sheet.cells[i,6].value;		//灌浆段长
     E_Target.cells[j,7].value:=E_Sheet.cells[i,7].value;		//孔径
     try if strtofloat(E_Target.cells[j,8].value)=0 then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value    	//透水率
     except if Trim(E_Target.cells[j,8].value)='' then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value end;
     try if strtofloat(E_Target.cells[j,9].value)=0 then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value     //水灰比起
     except if Trim(E_Target.cells[j,9].value)='' then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value end;
     E_Target.cells[j,10].value:=E_Sheet.cells[i,10].value;		//水灰比终
     try if strtofloat(E_Target.cells[j,11].value)=0 then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value    //注入率起
     except if Trim(E_Target.cells[j,11].value)='' then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value end;
     E_Target.cells[j,12].value:=E_Sheet.cells[i,12].value;		//注入率终
     try if strtofloat(E_Target.cells[j,13].value)=0 then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value		//水泥量浆
     except if Trim(E_Target.cells[j,13].value)='' then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value end;		//水泥量灰
     try if strtofloat(E_Target.cells[j,14].value)=0 then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value		//水泥量废
     except if Trim(E_Target.cells[j,14].value)='' then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value end;		//水泥量合
     try if strtofloat(E_Target.cells[j,15].value)=0 then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value		//单位注入量
     except if Trim(E_Target.cells[j,15].value)='' then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value end;		//水泥量浆
     try if strtofloat(E_Target.cells[j,16].value)=0 then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value		//水泥量灰
     except if Trim(E_Target.cells[j,16].value)='' then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value end;		//水泥量废
     try if strtofloat(E_Target.cells[j,17].value)=0 then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value		//水泥量合
     except if Trim(E_Target.cells[j,17].value)='' then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value end;		//单位注入量
     E_Target.cells[j,18].value:=E_Sheet.cells[i,18].value;		//灌浆压力
     try if strtofloat(E_Target.cells[j,19].value)=0 then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value    	//时间起
     except if Trim(E_Target.cells[j,19].value)='' then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value end;
     E_Target.cells[j,20].value:=E_Sheet.cells[i,20].value;		//时间终
     try if strtofloat(E_Target.cells[j,21].value)=0 then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value   	//时长
     except if Trim(E_Target.cells[j,21].value)='' then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value end;
     E_Target.cells[j,22].value:=E_Sheet.cells[i,22].value;		//备注
     j:=j+1;
    end else
   begin
    E_Target.cells[j,1].value:=E_Sheet.cells[i,1].value;		//ID
    E_Target.cells[j,2].value:=E_Sheet.cells[i,2].value;		//孔号
    E_Target.cells[j,3].value:=E_Sheet.cells[i,3].value;		//段次
    E_Target.cells[j,4].value:=E_Sheet.cells[i,4].value;		//灌浆起深
    E_Target.cells[j,5].value:=E_Sheet.cells[i,5].value;		//灌浆终深
    E_Target.cells[j,6].value:=E_Sheet.cells[i,6].value;		//灌浆段长
    E_Target.cells[j,7].value:=E_Sheet.cells[i,7].value;		//孔径
     try if strtofloat(E_Target.cells[j,8].value)=0 then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value    	//透水率
     except if Trim(E_Target.cells[j,8].value)='' then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value end;
     try if strtofloat(E_Target.cells[j,9].value)=0 then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value     //水灰比起
     except if Trim(E_Target.cells[j,9].value)='' then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value end;
     E_Target.cells[j,10].value:=E_Sheet.cells[i,10].value;		//水灰比终
     try if strtofloat(E_Target.cells[j,11].value)=0 then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value    //注入率起
     except if Trim(E_Target.cells[j,11].value)='' then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value end;
    E_Target.cells[j,12].value:=E_Sheet.cells[i,12].value;		//注入率终
    E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value;		//水泥量浆
    E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value;		//水泥量灰
    E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value;		//水泥量废
    E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value;		//水泥量合
    E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value;		//单位注入量
    E_Target.cells[j,18].value:=E_Sheet.cells[i,18].value;		//灌浆压力
     try if strtofloat(E_Target.cells[j,19].value)=0 then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value    	//时间起
     except if Trim(E_Target.cells[j,19].value)='' then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value end;
    E_Target.cells[j,20].value:=E_Sheet.cells[i,20].value;		//时间终
    E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value;		//时长
    E_Target.cells[j,22].value:=E_Sheet.cells[i,22].value;		//备注
   end
  end
 end
 else
 for i:=1 to maxR-1 do
 begin                    //StrToInt(E_Sheet.cells.item[k,1]);
  StatusBar1.Panels[1].Text:=IntToStr(i);
  E_TTarg.activate;
  if E_Target.cells[i,2].value<>E_Target.cells[i+1,2].value then
   begin
    E_TTarg.cells[j,1].value:=E_Target.cells[i,1].value;		//ID
    E_TTarg.cells[j,2].value:=E_Target.cells[i,2].value;		//孔号
    try if strtofloat(E_TTarg.cells[j,8].value)=0 then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value     //起灌深
    except if Trim(E_TTarg.cells[j,8].value)='' then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value end;
    E_TTarg.cells[j,9].value:=E_Target.cells[i,5].value;		//终灌深
    try if strtofloat(E_TTarg.cells[j,10].value)=0 then E_TTarg.cells[j,10].value:=E_Target.cells[i,6].value else E_TTarg.cells[j,10].value:=E_TTarg.cells[j,10].value+E_Target.cells[i,6].value		//段长
    except if Trim(E_TTarg.cells[j,10].value)='' then E_TTarg.cells[j,10].value:=E_Target.cells[i,6].value else E_TTarg.cells[j,10].value:=E_TTarg.cells[j,10].value+E_Target.cells[i,6].value end;
    try if strtofloat(E_TTarg.cells[j,11].value)=0 then E_TTarg.cells[j,11].value:=E_Target.cells[i,8].value else E_TTarg.cells[j,11].value:=(E_TTarg.cells[j,11].value+E_Target.cells[i,8].value)/k		//平均透水率
    except if Trim(E_TTarg.cells[j,11].value)='' then E_TTarg.cells[j,11].value:=E_Target.cells[i,8].value else E_TTarg.cells[j,11].value:=(E_TTarg.cells[j,11].value+E_Target.cells[i,8].value)/k end;
    try if strtofloat(E_TTarg.cells[j,12].value)=0 then E_TTarg.cells[j,12].value:=E_Target.cells[i,13].value else E_TTarg.cells[j,12].value:=E_TTarg.cells[j,12].value+E_Target.cells[i,13].value		//水泥量浆
    except if Trim(E_TTarg.cells[j,12].value)='' then E_TTarg.cells[j,12].value:=E_Target.cells[i,13].value else E_TTarg.cells[j,12].value:=E_TTarg.cells[j,12].value+E_Target.cells[i,13].value end;
    try if strtofloat(E_TTarg.cells[j,13].value)=0 then E_TTarg.cells[j,13].value:=E_Target.cells[i,14].value else E_TTarg.cells[j,13].value:=E_TTarg.cells[j,13].value+E_Target.cells[i,14].value		//水泥量灰
    except if Trim(E_TTarg.cells[j,13].value)='' then E_TTarg.cells[j,13].value:=E_Target.cells[i,14].value else E_TTarg.cells[j,13].value:=E_TTarg.cells[j,13].value+E_Target.cells[i,14].value end;
    try if strtofloat(E_TTarg.cells[j,14].value)=0 then E_TTarg.cells[j,14].value:=E_Target.cells[i,15].value else E_TTarg.cells[j,14].value:=E_TTarg.cells[j,14].value+E_Target.cells[i,15].value		//水泥量废
    except if Trim(E_TTarg.cells[j,14].value)='' then E_TTarg.cells[j,14].value:=E_Target.cells[i,15].value else E_TTarg.cells[j,14].value:=E_TTarg.cells[j,14].value+E_Target.cells[i,15].value end;
    try if strtofloat(E_TTarg.cells[j,15].value)=0 then E_TTarg.cells[j,15].value:=E_Target.cells[i,16].value else E_TTarg.cells[j,15].value:=E_TTarg.cells[j,15].value+E_Target.cells[i,16].value		//水泥量合
    except if Trim(E_TTarg.cells[j,15].value)='' then E_TTarg.cells[j,15].value:=E_Target.cells[i,16].value else E_TTarg.cells[j,15].value:=E_TTarg.cells[j,15].value+E_Target.cells[i,16].value end;
    try if strtofloat(E_TTarg.cells[j,16].value)=0 then E_TTarg.cells[j,16].value:=E_Target.cells[i,17].value else E_TTarg.cells[j,16].value:=E_TTarg.cells[j,13].value/E_TTarg.cells[j,10].value		//单位注入量
    except if Trim(E_TTarg.cells[j,16].value)='' then E_TTarg.cells[j,16].value:=E_Target.cells[i,17].value else E_TTarg.cells[j,16].value:=E_TTarg.cells[j,13].value/E_TTarg.cells[j,10].value end;
    E_TTarg.cells[j,17].value:=IntToStr(k)+'段';		//备注
    j:=j+1; k:=1;
    end else
   begin
    E_TTarg.cells[j,1].value:=E_Target.cells[i,1].value;		//ID
    E_TTarg.cells[j,2].value:=E_Target.cells[i,2].value;		//孔号
    try if strtofloat(E_TTarg.cells[j,8].value)=0 then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value     //起灌深
    except if Trim(E_TTarg.cells[j,8].value)='' then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value end;
    E_TTarg.cells[j,9].value:=E_Target.cells[i,5].value;		//终灌深
    E_TTarg.cells[j,10].value:=E_TTarg.cells[j,10].value+E_Target.cells[i,6].value;		//段长  *
    E_TTarg.cells[j,11].value:=E_TTarg.cells[j,11].value+E_Target.cells[i,8].value;		//平均透水率
    E_TTarg.cells[j,12].value:=E_TTarg.cells[j,12].value+E_Target.cells[i,13].value;		//平均透水率
    E_TTarg.cells[j,13].value:=E_TTarg.cells[j,13].value+E_Target.cells[i,14].value;		//平均透水率
    E_TTarg.cells[j,14].value:=E_TTarg.cells[j,14].value+E_Target.cells[i,15].value;		//平均透水率
    E_TTarg.cells[j,15].value:=E_TTarg.cells[j,15].value+E_Target.cells[i,16].value;		//平均透水率
    E_TTarg.cells[j,16].value:=E_TTarg.cells[j,16].value+E_Target.cells[i,17].value;		//平均透水率
    k:=k+1;
   end;
 end;
 E_App.ActiveWorkBook.Save;
 E_App.WorkBooks.Close;         //关闭工作簿
 E_App.Quit;                    //退出 Excel
 E_App:=Unassigned;             //释放excel进程
 ShowMessage('数据整理完成！');
end;

end.
