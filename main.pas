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

function DATA_Connect(Data_Name,Tab_Name:string;TF:Boolean):Boolean;      //accessConnect����Ŀ�ģ�����Access
var i,m:Integer;
begin
        // ����Ϊ��ʼ�����ݿ�����
 Form1.ADOC_S.Connected:=False;       //�������ݿ�����ΪFalse���ر����ᣬ�Ա�����һ������
 Form1.ADOC_S.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_name+';Persist Security Info=False';
 Form1.ADOC_S.LoginPrompt:=false;        //����ʾ��¼��

 Form1.Qur_S.Connection:=Form1.ADOC_S;
 Form1.DS_S.DataSet:=Form1.Qur_S;

 Form1.ListBox1.Clear;
 with Form1.Qur_S do                         //������ڵ��������������
 begin
   Close;
   SQL.Clear;
   SQL.Add('select distinct �׺� from '+Tab_Name);  //Caran_GJL');
   Open;
 end;
  m:=Form1.Qur_S.RecordCount-1;

  for i:=0 to m do                                   //�������ż���Combobox1��ѡ����
  begin
    Form1.ListBox1.Items.Add(Form1.Qur_S.fieldbyname('�׺�').AsString);
//    ListBox2.ItemIndex:=0;
    Form1.Qur_S.Next;
  end;
 result:=true;                     //���������з���ֵ��Ϊ���ֳ�������Ϣ������true   }
end;

function Ex_Tro(E_Name:String):Boolean;      //��ȡExcel��
 var maxR,i:Integer;
  Str_Q:string;
begin

 Result:=True;
end;


function Ex_TroG(E_Name:String):Boolean;      //��ȡExcel��
 var maxR,i:Integer;
  Str_Q:string;
begin
  maxR:= E_Sheet.usedrange.rows.count;     //�����ݵ�����
  Str_Q:='insert into Caran_GJL (�׺�,�δ�,�ཬ����,�ཬ����,�ཬ�γ�,�׾�,͸ˮ��,ˮ�ұ���,ˮ�ұ���,ע������,ע������,ˮ��ע����,ˮ��ע����,ˮ�������,ˮ��ϼ�,��λע����,�ཬѹ��,ʱ����,ʱ����,ʱ��,��ע) values';
  Str_Q:=Str_Q+'(:I_D,:D_i,:D_S,:D_E,:D_L,:D_R,:K_L,:K_s,:K_e,:Z_S,:Z_E,:V_a,:V_b,:V_c,:V_v,:V_m,:V_p,:T_S,:T_E,:T_T,:K_B)';
 for i:=6 to maxR do
 begin
  if not TryStrToInt(Trim(E_Sheet.cells.item[i,1]),maxR) then continue;                     //����δβ�����������������ѭ��
  if (Trim(E_Sheet.cells.item[i,2])<>'') and (Trim(E_Sheet.cells.item[i,3])<>'') and  (Trim(E_Sheet.cells.item[i,1])<>'fk') then       //������в�Ϊ��������
  with Form1.Qur_S do
  begin
    Close;
    SQL.Clear;
    SQL.Add(Str_Q);
		Parameters.ParamByName('I_D').Value:=E_Name;		//�׺�
 		Parameters.ParamByName('D_i').Value:=StrToInt(E_Sheet.cells.item[i,1]);  //E_Sheet.cells[i,1].value;		//�δ�
 		Parameters.ParamByName('D_S').Value:=E_Sheet.cells[i,2].value;		//�ཬ����
 		Parameters.ParamByName('D_E').Value:=E_Sheet.cells[i,3].value;		//�ཬ����
		Parameters.ParamByName('D_L').Value:=E_Sheet.cells[i,4].value;		//�ཬ�γ�
    Parameters.ParamByName('D_R').Value:=StrToInt(RightStr(E_Sheet.cells[i,5].value, 2));		//�׾�
    if Trim(E_Sheet.cells.item[i,6])='/' then Parameters.ParamByName('K_L').Value:=0 else Parameters.ParamByName('K_L').Value:=E_Sheet.cells[i,6].value;		//͸ˮ��
 		Parameters.ParamByName('K_s').Value:=E_Sheet.cells[i,7].value;		//ˮ�ұ���
		Parameters.ParamByName('K_e').Value:=E_Sheet.cells[i,8].value;		//ˮ�ұ���
		Parameters.ParamByName('Z_S').Value:=E_Sheet.cells[i,9].value;		//ע������
		Parameters.ParamByName('Z_E').Value:=E_Sheet.cells[i,10].value;		//ע������
		Parameters.ParamByName('V_a').Value:=E_Sheet.cells[i,11].value;		//ˮ������
		Parameters.ParamByName('V_b').Value:=E_Sheet.cells[i,12].value;		//ˮ������
		Parameters.ParamByName('V_c').Value:=E_Sheet.cells[i,13].value;		//ˮ������
		Parameters.ParamByName('V_v').Value:=E_Sheet.cells[i,14].value;		//ˮ������
		Parameters.ParamByName('V_m').Value:=E_Sheet.cells[i,15].value;		//��λע����
		Parameters.ParamByName('V_p').Value:=E_Sheet.cells[i,16].value;		//�ཬѹ��

    Parameters.ParamByName('T_S').Value:=E_Sheet.cells[i,17].value;
    Parameters.ParamByName('T_E').Value:=E_Sheet.cells[i,19].value;

{    Parameters.ParamByName('T_S').Value:=FormatdateTime('c',StrToDateTime(DateToStr(E_Sheet.cells[i,17].value)+' '+TimeToStr(E_Sheet.cells[i,18].value)));
    Parameters.ParamByName('T_E').Value:=FormatdateTime('c',StrToDateTime(DateToStr(E_Sheet.cells[i,19].value)+' '+TimeToStr(E_Sheet.cells[i,20].value)));
} 		Parameters.ParamByName('T_T').Value:=FormatdateTime('tt',E_Sheet.cells[i,21].value);		//ʱ��
		Parameters.ParamByName('K_B').Value:=E_Sheet.cells[i,22].value;		//��ע
    ExecSQL;                    //�����ݼ���ִ��SQL���
  end
  else Break;
 end;
 Result:=True;
end;

 function Data_EhH(D_Eh:TDBGridEh):Boolean;
 var Col:TColumnEh;
 begin
    D_Eh.UseMultiTitle:=True; //�Ƿ�ʹ�ö��б�����
    D_Eh.TitleLines:=2; //����������
    D_Eh.Flat:=True; //ƽ����ʾ;FalseΪ������ʾ
    D_Eh.Columns.Clear;

		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�δ�';
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�ཬ�׶�(m)|�γ�'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�׾�'+#13+'(mm)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='͸ˮ��'+#13+'(Lu)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ˮ�ұ�|��ʼ'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ˮ�ұ�|��ֹ'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ע����'+#13+'(L/min)|��ʼ'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ע����'+#13+'(L/min)|��ֹ'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ˮ������|ע��'+#13+'(L)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ˮ������|ע��'+#13+'(kg)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ˮ������|����'+#13+'(kg)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='ˮ������|�ϼ�'+#13+'(kg)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='��λ'+#13+'ע����'+#13+'(kg/m)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�ཬ'+#13+'ѹ��'+#13+'(Mpa)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�ཬʱ��|��ʼ'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�ཬʱ��|��ֹ'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='�ཬʱ��|����'+#13+'(hh:mm)'; //�������ı�
		Col:=D_Eh.Columns.Add;
		Col.Title.Caption:='��ע'; //�������ı�

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
    Columns[6].Alignment:=taRightJustify;         //��Ԫ�����ݾ���
    Columns[7].Alignment:=taRightJustify;         //��Ԫ�����ݾ���
    Columns[18].Alignment:=taRightJustify;         //��Ԫ�����ݾ���

//    Columns[0].ReadOnly:=True;                  //���ñ��Ĳ����в���д
   end;
    Result:=True;
 end;

 function Data_EhR(D_Eh:TDBGridEh):Boolean;
 begin
   with D_Eh do
   begin
    UseMultiTitle:=True; //�Ƿ�ʹ�ö��б�����
    TitleLines:=2; //����������
    Flat:=True; //ƽ����ʾ;FalseΪ������ʾ
    Columns.Clear;

		Columns[0].Title.Caption:='�׺�';
		Columns[1].Title.Caption:='׮��';
		Columns[2].Title.Caption:='����';
		Columns[3].Title.Caption:='����';
		Columns[4].Title.Caption:='�׿ڸ߳�'+#13+'(m)';
		Columns[5].Title.Caption:='��������'+#13+'(m)';
		Columns[6].Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Columns[7].Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Columns[8].Title.Caption:='�ཬ�׶�(m)|�γ�'; //�������ı�
		Columns[9].Title.Caption:='͸ˮ��'+#13+'(Lu)'; //�������ı�
		Columns[10].Title.Caption:='ˮ������|ע��'+#13+'(L)'; //�������ı�
		Columns[11].Title.Caption:='ˮ������|ע��'+#13+'(kg)'; //�������ı�
		Columns[12].Title.Caption:='ˮ������|����'+#13+'(kg)'; //�������ı�
		Columns[13].Title.Caption:='ˮ������|�ϼ�'+#13+'(kg)'; //�������ı�
		Columns[14].Title.Caption:='��λ'+#13+'ע����'+#13+'(kg/m)'; //�������ı�
             {
		Columns[0].Width:=40;
		Columns[1].Width:=60;
		Columns[2].Width:=20;
		Columns[3].Width:=20; //����
		Columns[4].Width:=50;
		Columns[5].Width:=30;
		Columns[6].Width:=35;
		Columns[7].Width:=40;
		Columns[8].Width:=35;
		Columns[9].Width:=40;    //͸ˮ��
		Columns[10].Width:=60;
		Columns[11].Width:=60;
		Columns[12].Width:=40;
		Columns[13].Width:=60;
		Columns[14].Width:=50;
                    }
    Columns[15].Width:=80;
    Columns[1].Alignment:=taLeftJustify;         //��Ԫ�����ݾ���
    Columns[2].Alignment:=taCenter;         //��Ԫ�����ݾ���
    Columns[3].Alignment:=taCenter;         //��Ԫ�����ݾ���

//    Columns[0].ReadOnly:=True;                  //���ñ��Ĳ����в���д
    Columns[2].STFilter.ListSource:=Form1.DS_S;
    Columns[3].STFilter.ListSource:=Form1.DS_S;

    FooterColor := clScrollBar;
    FooterRowCount :=1;
    SumList.Active := true;
    Columns[0].Footer.valuetype := fvtcount;
    Columns[8].Footer.valuetype := fvtsum;
    Columns[10].Footer.valuetype := fvtsum;     //ע����
    Columns[11].Footer.valuetype := fvtsum;
    Columns[12].Footer.valuetype := fvtsum;
    Columns[13].Footer.valuetype := fvtsum;
    Columns[14].Footer.valuetype := fvtAvg;    //��λע����

   end;
    Result:=True;
 end;


 function Data_Eh(D_Eh:TDBGridEh):Boolean;
 begin
   with D_Eh do
   begin
    UseMultiTitle:=True; //�Ƿ�ʹ�ö��б�����
    TitleLines:=2; //����������
    Flat:=True; //ƽ����ʾ;FalseΪ������ʾ
    Columns.Clear;
		Columns[0].Title.Caption:='�δ�';
		Columns[1].Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Columns[2].Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Columns[3].Title.Caption:='�ཬ�׶�(m)|�γ�'; //�������ı�
		Columns[4].Title.Caption:='�׾�'+#13+'(mm)'; //�������ı�
		Columns[5].Title.Caption:='͸ˮ��'+#13+'(Lu)'; //�������ı�
		Columns[6].Title.Caption:='ˮ�ұ�|��ʼ'; //�������ı�
		Columns[7].Title.Caption:='ˮ�ұ�|��ֹ'; //�������ı�
		Columns[8].Title.Caption:='ע����'+#13+'(L/min)|��ʼ'; //�������ı�
		Columns[9].Title.Caption:='ע����'+#13+'(L/min)|��ֹ'; //�������ı�
		Columns[10].Title.Caption:='ˮ������|ע��'+#13+'(L)'; //�������ı�
		Columns[11].Title.Caption:='ˮ������|ע��'+#13+'(kg)'; //�������ı�
		Columns[12].Title.Caption:='ˮ������|����'+#13+'(kg)'; //�������ı�
		Columns[13].Title.Caption:='ˮ������|�ϼ�'+#13+'(kg)'; //�������ı�
		Columns[14].Title.Caption:='��λ'+#13+'ע����'+#13+'(kg/m)'; //�������ı�
		Columns[15].Title.Caption:='�ཬ'+#13+'ѹ��'+#13+'(Mpa)'; //�������ı�
		Columns[16].Title.Caption:='�ཬʱ��|��ʼ'; //�������ı�
		Columns[17].Title.Caption:='�ཬʱ��|��ֹ'; //�������ı�
		Columns[18].Title.Caption:='�ཬʱ��|����'+#13+'(hh:mm)'; //�������ı�
		Columns[19].Title.Caption:='��ע'; //�������ı�
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
    Columns[6].Alignment:=taRightJustify;         //��Ԫ�����ݾ���
    Columns[7].Alignment:=taRightJustify;         //��Ԫ�����ݾ���
    Columns[18].Alignment:=taRightJustify;         //��Ԫ�����ݾ���

//    Columns[0].ReadOnly:=True;                  //���ñ��Ĳ����в���д    
    Columns[0].STFilter.ListSource:=Form1.DS_S;

    //    Columns[0].Width:=40;                   //���       ��ʾ��Dbgrid�е��п�����
//    Columns[0].Title.Alignment:=taCenter;        //��ͷ���־���
//    Columns[0].Alignment:=taCenter;         //��Ԫ�����ݾ���
//    Columns[0].Color:=ColorToRGB($ABABAB);         //��ǰ��ɫ
               {
    Columns[1].Footer.ValueType:=fvtStaticText;
    Columns[1].Footer.Value:='��С�̣߳�';
    Columns[4].Footer.ValueType:=fvtMin;
    Columns[3].Footer.ValueType:=fvtStaticText;
    Columns[3].Footer.Value:='���̣߳�';
    Columns[4].Footer.ValueType:=fvtmax;
    Columns[6].Footer.ValueType:=fvtStaticText;
    Columns[6].Footer.Value:='�ܳ���';
    Columns[7].Footer.ValueType:=fvtmax;

     DbgEh_S.Columns[0].Footers.Add; // ����Footer����

// �������е�1��

  DbgEh_S.Columns[0].Footers[0].ValueType:=fvtStaticText; // ��ʾ�ı�

  DbgEh_S.Columns[0].Footers[0].Value:='�ϼ�';

  DbgEh_S.Columns[0].Footers[0].Alignment:=taCenter; // ���Ķ���

  DbgEh_S.Columns[0].Footers.Add; // ����Footer����

// �������е�2��

  DbgEh_S.Columns[0].Footers[1].ValueType:=fvtCount; // ����

  DbgEh_S.Columns[0].Footers[1].FieldName:='���'; // �ֶ���

  DbgEh_S.Columns[0].Footers[1].Alignment:=taCenter; // ���Ķ���

  DbgEh_S.Columns[3].Footers.Add; // �������е�4��

// �������е�4��

  DbgEh_S.Columns[3].Footers[0].ValueType:=fvtSum; // �������ͣ��ϼ�

  DbgEh_S.Columns[3].Footers[0].FieldName:='���'; // �ֶ���

  DbgEh_S.Columns[3].Footers[0].DisplayFormat:='#,###,###.00'; // ��ʾ��ʽ

  DbgEh_S.Columns[3].Footers.Add; // ������е�4��

// ���ô��е�4��

  DbgEh_S.Columns[3].Footers[1].ValueType:=fvtFieldValue; // �������ͣ��ֶ�ֵ

  DbgEh_S.Columns[3].Footers[1].FieldName:='�˺�'; // �ֶ���

  DbgEh_S.Columns[3].Footers[1].Font.Style:=[fsBold]; // ���ָ�ʽ

  DbgEh_S.Columns[3].Footers[1].Font.Color:=clBlue; // ���ֳߴ�

  DbgEh_S.SumList.Active:=True; // ȷ�� ͳ�ƺϼ�

end;

             }

    FooterColor :=clInactiveCaption;        //clScrollBar;
    FooterRowCount :=1;
    SumList.Active := true;
    Columns[0].Footer.valuetype := fvtcount;
    Columns[3].Footer.valuetype := fvtsum;
    Columns[10].Footer.valuetype := fvtsum;     //ע����
    Columns[11].Footer.valuetype := fvtsum;
    Columns[12].Footer.valuetype := fvtsum;
    Columns[13].Footer.valuetype := fvtsum;
    Columns[14].Footer.valuetype := fvtAvg;    //��λע����
    Columns[15].Footer.valuetype := fvtAvg;
   end;
    Result:=True;
 end;


 function Data_EhA(D_Eh:TDBGridEh):Boolean;
 begin
   with D_Eh do
   begin
    UseMultiTitle:=True; //�Ƿ�ʹ�ö��б�����
    TitleLines:=2; //����������
    Flat:=True; //ƽ����ʾ;FalseΪ������ʾ
    Columns.Clear;

		Columns[0].Title.Caption:='�׺�';
		Columns[1].Title.Caption:='�δ�';
		Columns[2].Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Columns[3].Title.Caption:='�ཬ�׶�(m)|��'; //�������ı�
		Columns[4].Title.Caption:='�ཬ�׶�(m)|�γ�'; //�������ı�
		Columns[5].Title.Caption:='�׾�'+#13+'(mm)'; //�������ı�
		Columns[6].Title.Caption:='͸ˮ��'+#13+'(Lu)'; //�������ı�
		Columns[7].Title.Caption:='ˮ�ұ�|��ʼ'; //�������ı�
		Columns[8].Title.Caption:='ˮ�ұ�|��ֹ'; //�������ı�
		Columns[9].Title.Caption:='ע����'+#13+'(L/min)|��ʼ'; //�������ı�
		Columns[10].Title.Caption:='ע����'+#13+'(L/min)|��ֹ'; //�������ı�
		Columns[11].Title.Caption:='ˮ������|ע��'+#13+'(L)'; //�������ı�
		Columns[12].Title.Caption:='ˮ������|ע��'+#13+'(kg)'; //�������ı�
		Columns[13].Title.Caption:='ˮ������|����'+#13+'(kg)'; //�������ı�
		Columns[14].Title.Caption:='ˮ������|�ϼ�'+#13+'(kg)'; //�������ı�
		Columns[15].Title.Caption:='��λ'+#13+'ע����'+#13+'(kg/m)'; //�������ı�
		Columns[16].Title.Caption:='�ཬ'+#13+'ѹ��'+#13+'(Mpa)'; //�������ı�
		Columns[17].Title.Caption:='�ཬʱ��|��ʼ'; //�������ı�
		Columns[18].Title.Caption:='�ཬʱ��|��ֹ'; //�������ı�
		Columns[19].Title.Caption:='�ཬʱ��|����'+#13+'(hh:mm)'; //�������ı�
		Columns[20].Title.Caption:='��ע'; //�������ı�
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
    Columns[7].Alignment:=taRightJustify;         //��Ԫ�����ݾ���
    Columns[8].Alignment:=taRightJustify;         //��Ԫ�����ݾ���
    Columns[19].Alignment:=taRightJustify;         //��Ԫ�����ݾ���

//    Columns[0].ReadOnly:=True;                  //���ñ��Ĳ����в���д
//    Columns[1].ReadOnly:=True;                  //���ñ��Ĳ����в���д
    Columns[0].STFilter.ListSource:=Form1.DS_S;
    Columns[1].STFilter.ListSource:=Form1.DS_S;

    FooterColor := clScrollBar;
    FooterRowCount :=1;
    SumList.Active := true;
    Columns[0].Footer.valuetype := fvtcount;
    Columns[4].Footer.valuetype := fvtsum;
    Columns[11].Footer.valuetype := fvtsum;     //ע����
    Columns[12].Footer.valuetype := fvtsum;
    Columns[13].Footer.valuetype := fvtsum;
    Columns[14].Footer.valuetype := fvtsum;
    Columns[15].Footer.valuetype := fvtAvg;    //��λע����
    Columns[16].Footer.valuetype := fvtAvg;
   end;
    Result:=True;
 end;

procedure TForm1.Button1Click(Sender: TObject);     //��ȡExcel�ļ�
var I: Integer;
begin
  opendialog1.filter:='Excel(*.xls)|*.xls|Excel2007(*.xlsx)|*.xlsx';
  if opendialog1.Execute then
  begin
    ListBox2.Clear;
  //  files:=opendialog1.Files;     //��������ļ���
    for i:=0 to opendialog1.Files.count-1 do
    ListBox2.Items.Add(opendialog1.Files[i])//list[i]���Ǵ����ѡ�ļ��ĵ����ļ���
  end;
  ListBox2.ItemIndex:=0;
  Button4.Enabled:=True;
  Memo1.Clear;
  Memo2.Clear;
end;

procedure TForm1.FormCreate(Sender: TObject);
var i:Variant;
begin
  i:=Trunc(Screen.Height);           //����߶�
  Form1.Top:=(i-Form1.Height)/2;
  i:=Trunc(Screen.Width);             //������
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

      //OpenDialog��Options\ofAllowMultiSelect������Ϊtrue����ʾ����ѡ�����ļ���
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
	TFloatField(ADO.FieldByName('�ཬ����')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('�ཬ����')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('�ཬ�γ�')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('͸ˮ��')).DisplayFormat:='#.##';
	TFloatField(ADO.FieldByName('ע������')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('ע������')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('ˮ��ע����')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('ˮ��ע����')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('ˮ�������')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('ˮ��ϼ�')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('��λע����')).DisplayFormat:='0.##';
	TFloatField(ADO.FieldByName('�ཬѹ��')).DisplayFormat:='0.##';
//  TTimeField(ADO.FieldByName('ʱ��')).DisplayFormat:='hh:mm';
 Result:=True;
end;

function Data_AdoR(ADO:TADOQuery):Boolean;
begin
 TFloatField(ADO.FieldByName('�׿ڸ߳�')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('��������')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('�ཬ����')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('�ཬ����')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('�ཬ�γ�')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('ƽ��͸ˮ��')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('ˮ��ע����')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('ˮ��ע����')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('ˮ�������')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('ˮ��ϼ�')).DisplayFormat:='0.##';
 TFloatField(ADO.FieldByName('��λע����')).DisplayFormat:='0.##';
 Result:=True;
end;

procedure TForm1.StatusBar1Click(Sender: TObject);
begin
  ShowMessage(StatusBar1.Panels[0].Text+#13+StatusBar1.Panels[1].Text+#13+StatusBar1.Panels[2].Text);
end;

function DATA_rre(D_Ado:TADOQuery;D_Dbge:TDBGridEh;D_Ds:TDataSource;D_Tab,data_Name:String):Boolean;   //ˢ������
 begin
   with D_Ado do
   begin
    close;
    sql.Clear;
    SQL.Add('select '+D_Tab+'='''+data_Name+'''');
    SQL.Add('order by ID');              //��ID������
    open;
   end;

   with D_Dbge do
   begin
    DataSource:=D_Ds;
    AutoFitColWidths:=True;                               //����Ӧ��Ԫ����
    ColumnDefValues.Title.TitleButton:=True;             //�����ж���������
// 		ColumnDefValues.Title.SortMarker:=smDownEh;         //ָ�������־
    OptionsEh:=OptionsEh+[dghAutoSortMarking];            //�����Զ�������
    Options:=Options+[dgEditing]+[dgMultiSelect];   //  ��Ҫ����[dgRowSelect]�������ܱ༭�������
    SortLocal:=True;                                      //�ͻ�������
    STFilter.Visible := True;
    STFilter.Local := True;
   end;
   Result:=True;
 end;

 function DATA_re(D_Ado:TADOQuery;D_Dbge:TDBGridEh;D_Ds:TDataSource;D_Tab:String):Boolean;   //ˢ������
 begin
   with D_Ado do
   begin
    close;
    sql.Clear;
    SQL.Add('select '+D_Tab);
    SQL.Add('order by ID');              //��ID������
    open;
   end;

   with D_Dbge do
   begin
    DataSource:=D_Ds;
    AutoFitColWidths:=True;                               //����Ӧ��Ԫ����
    ColumnDefValues.Title.TitleButton:=True;             //�����ж���������
// 		ColumnDefValues.Title.SortMarker:=smDownEh;         //ָ�������־
    OptionsEh:=OptionsEh+[dghAutoSortMarking];            //�����Զ�������
    Options:=Options+[dgEditing]+[dgMultiSelect];   //  ��Ҫ����[dgRowSelect]�������ܱ༭�������
    SortLocal:=True;                                      //�ͻ�������
    STFilter.Visible := True;
    STFilter.Local := True;
   end;
   Result:=True;
 end;

 function DATA_r(D_Ado:TADOQuery;D_Dbge:TDBGridEh;D_Ds:TDataSource;D_Tab:String):Boolean;   //ˢ������
 begin
   with D_Ado do
   begin
    close;
    sql.Clear;
    SQL.Add('select '+D_Tab);
    SQL.Add('order by ID');              //��ID������
    open;
   end;

   with D_Dbge do
   begin
    DataSource:=D_Ds;
    AutoFitColWidths:=True;                               //����Ӧ��Ԫ����
    ColumnDefValues.Title.TitleButton:=True;             //�����ж���������
// 		ColumnDefValues.Title.SortMarker:=smDownEh;         //ָ�������־
    OptionsEh:=OptionsEh+[dghAutoSortMarking];            //�����Զ�������
    Options:=Options+[dgEditing]+[dgMultiSelect];   //  ��Ҫ����[dgRowSelect]�������ܱ༭�������
    SortLocal:=True;                                      //�ͻ�������
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

  Str_Q:='�δ�,�ཬ����,�ཬ����,�ཬ�γ�,�׾�,͸ˮ��,ˮ�ұ���,ˮ�ұ���,ע������,ע������,ˮ��ע����,ˮ��ע����,ˮ�������,ˮ��ϼ�,��λע����,�ཬѹ��,ʱ����,ʱ����,ʱ��,��ע from Caran_GJL where �׺�';
  DATA_rre(Qur_S,DbgEh_S,DS_S,Str_Q,ListBox1.Items[ListBox1.itemindex]);    //ˢ������
  Data_Ado(Qur_S);         //����������ʾ��ʽ
  Data_Eh(DbgEh_S);         //����������ʾ��ʽ
  StatusBar1.Panels[1].Text:=ListBox1.Items[ListBox1.itemindex]+' ����ϸ����';
end;

procedure TForm1.ListBox2DblClick(Sender: TObject);
var  I,N: Integer;//��¼���ݱ�ĵ�ǰ��¼��
  F: TextFile;  //TextFile �� Text ��һ����
  Str: string;
begin
  try                                            //���ӵ�Excel
    E_App := CreateOleObject('Excel.Application');
    E_App.Visible := true;                       //��ʾ�򿪵�Excel�ļ�
   if FileExists(ListBox2.Items[ListBox2.itemindex]) then
    begin
     E_App.WorkBooks.Open(ListBox2.Items[ListBox2.itemindex]);   //���ļ�

     AssignFile(F,StatusBar1.Panels[2].Text);
     Append(F);  //��׼��׷��
     for i:=1 to E_App.Sheets.Count do
     begin
       E_Sheet:=E_App.worksheets[i];
       E_Sheet.activate;
       if (Trim(E_Sheet.cells.item[7,2])='') and (Trim(E_Sheet.cells.item[7,3])='') then continue;
       if E_Sheet.Name<>'�ϼ�' then Ex_Tro(E_Sheet.Name);      //��ȡExcel��
       N:=E_Sheet.usedrange.rows.count;
       repeat N := N - 1 until Trim(E_Sheet.cells.item[N,1])<>'';

       Str:=E_Sheet.Name+#9+E_Sheet.cells[4,1].value+#9+FloatToStr(E_Sheet.cells[N,2].value)+#9+FloatToStr(E_Sheet.cells[N,3].value)+#9+FloatToStr(E_Sheet.cells[N,4].value)+#9+FloatToStr(E_Sheet.cells[N,6].value)+#9;
       Str:=Str+FloatToStr(E_Sheet.cells[N,11].value)+#9+FloatToStr(E_Sheet.cells[N,12].value)+#9+FloatToStr(E_Sheet.cells[N,13].value)+#9+FloatToStr(E_Sheet.cells[N,14].value)+#9+FloatToStr(E_Sheet.cells[N,15].value);
       Writeln(F, Str);
     end;

     CloseFile(F);

     E_App.WorkBooks.Close; //�رչ�����
     E_App.Quit; //�˳� Excel
     E_App:=Unassigned;//�ͷ�excel����
     DATA_Connect(StatusBar1.Panels[0].Text,'Caran_GJL',True);           //�������ݿ�
     ShowMessage(ListBox2.Items[ListBox2.itemindex]+#13+'�������');
     StatusBar1.Panels[1].Text:=ListBox2.Items[ListBox2.Itemindex];
    end
  except ShowMessage('Excel�ļ����ó���򲻴��ڣ�');
	end;
end;

function D_Create(D_ADO:TADOQuery;T_Name,T_Str:string):Boolean;
begin
  try
  with D_ADO do
  begin
    Close;
    SQL.Clear;
    Active:=false;                                          //�������ݿ��еı�     ��˾��Ϣ
    SQL.Add('create table '+T_Name+' (ID AUTOINCREMENT,'+T_Str+'primary key(ID))');
    ExecSQL;
  end;
 except
 end;
 Result:=True;
end;

function DATA_Create(Data_Name:string):Boolean;    //�������ݿ�
 var CreateAccess:OleVariant;
 Sql_Data:string;
begin
  CreateAccess:=CreateOleObject('ADOX.Catalog');
  CreateAccess.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_name+'.MDB');

  Form1.ADOC_S.Connected:=False;       //�������ݿ�����ΪFalse���ر����ᣬ�Ա�����һ������
  Form1.ADOC_S.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_name+'.mdb;Persist Security Info=False';
//  ADOC_S.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+data_bame+'.mdb;Persist Security Info=False;Jet OLEDB:Database Password=$#��ailihongbi 80��^%&';
  Form1.ADOC_S.LoginPrompt:=false;        //����ʾ��¼��
  Form1.ADOC_S.Connected:=true;
  Form1.Qur_S.Connection:=Form1.ADOC_S;
                                          //�������ݿ��еı�
  Sql_Data:='�׺� varchar(8),�δ� Byte,�ཬ���� Single,�ཬ���� Single,�ཬ�γ� Single,�׾� Byte,͸ˮ�� Single,ˮ�ұ��� varchar(8),ˮ�ұ��� varchar(8),ע������ Single,ע������ Single,';
  sql_Data:=Sql_Data+'ˮ��ע���� Single,ˮ��ע���� Single,ˮ������� Single,ˮ��ϼ� Single,��λע���� Single,�ཬѹ�� Single,ʱ���� Date,ʱ���� Date,ʱ�� Date,��ע varchar(20),';
  D_Create(Form1.Qur_S,'Caran_GJL',Sql_Data);

  Sql_Data:='�׺� varchar(8),�δ� Byte,�ཬ���� Single,�ཬ���� Single,�ཬ�γ� Single,�׾� Byte,͸ˮ�� Single,ˮ�ұ��� varchar(8),ˮ�ұ��� varchar(8),ע������ Single,ע������ Single,';
  sql_Data:=Sql_Data+'ˮ��ע���� Single,ˮ��ע���� Single,ˮ������� Single,ˮ��ϼ� Single,��λע���� Single,�ཬѹ�� Single,ʱ���� Date,ʱ���� Date,ʱ�� Date,��ע varchar(20),';
  D_Create(Form1.Qur_S,'Caran_GJT',Sql_Data);

  Sql_Data:='�׺� varchar(8),׮�� varchar(12),���� varchar(12),���� Byte,�׿ڸ߳� Single,�������� Single,�ཬ���� Single,�ཬ���� Single,�ཬ�γ� Single,ƽ��͸ˮ�� Single,';
  sql_Data:=Sql_Data+'ˮ��ע���� Single,ˮ��ע���� Single,ˮ������� Single,ˮ��ϼ� Single,��λע���� Single,��ע varchar(50),';
  D_Create(Form1.Qur_S,'Caran_T',Sql_Data);

  Result:=True;
end;


procedure TForm1.Button4Click(Sender: TObject);
var
  maxR,J,K,N,M: Integer;//��¼���ݱ�ĵ�ǰ��¼��
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
     E_App.Visible := true;                       //��ʾ�򿪵�Excel�ļ�
     E_App.WorkBooks.Open(ListBox2.Items[j]);   //���ļ�
     E_Sheet:=E_App.worksheets[1];
     E_Sheet.activate;
     F_Name:=Copy(StatusBar1.Panels[0].Text,1,Length(StatusBar1.Panels[0].Text)-4);
     AssignFile(F,F_Name+'.TXT');
     if not FileExists(F_Name+'.TXT') then Rewrite(F) else Append(F);   //�򿪴��ڵ��ļ������ļ�ָ�붨λ���ļ�β��

          maxR:= E_Sheet.usedrange.rows.count;     //�����ݵ�����
          Str:='insert into Caran_GJL (�׺�,�δ�,�ཬ����,�ཬ����,�ཬ�γ�,�׾�,͸ˮ��,ˮ�ұ���,ˮ�ұ���,ע������,ע������,ˮ��ע����,ˮ��ע����,ˮ�������,ˮ��ϼ�,��λע����,�ཬѹ��,ʱ����,ʱ����,ʱ��,��ע) values';
          Str:=Str+'(:I_D,:D_i,:D_S,:D_E,:D_L,:D_R,:K_L,:K_s,:K_e,:Z_S,:Z_E,:V_a,:V_b,:V_c,:V_v,:V_m,:V_p,:T_S,:T_E,:T_T,:K_B)';

         for k:=7 to maxR do
         begin
          if not TryStrToInt(Trim(E_Sheet.cells.item[k,1]),maxR) then
          begin
           continue;                     //����δβ�����������������ѭ��
          end;
          if (Trim(E_Sheet.cells.item[k,2])<>'') and (Trim(E_Sheet.cells.item[k,3])<>'') then       //������в�Ϊ��������
          try
           with Qur_S do
           begin
            Close;
            SQL.Clear;
            SQL.Add(Str);
        		Parameters.ParamByName('I_D').Value:=E_Sheet.Name;		//�׺�
         		try Parameters.ParamByName('D_i').Value:=StrToInt(E_Sheet.cells.item[k,1]);  //E_Sheet.cells[i,1].value;		//�δ�
            except Del_List.Add(E_Sheet.Name+#9+'�δ�'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('D_S').Value:=E_Sheet.cells[k,2].value;		//�ཬ����
            except Del_List.Add(E_Sheet.Name+#9+'����'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('D_E').Value:=E_Sheet.cells[k,3].value;		//�ཬ����
            except Del_List.Add(E_Sheet.Name+#9+'����'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('D_L').Value:=E_Sheet.cells[k,4].value;		//�ཬ�γ�
            except Del_List.Add(E_Sheet.Name+#9+'�γ�'+#9+inttostr(k)) end;
            try Parameters.ParamByName('D_R').Value:=E_Sheet.cells[k,5].value;		//�׾�
            except Del_List.Add(E_Sheet.Name+#9+'�׾�'+#9+inttostr(k)) end;
            try if Trim(E_Sheet.cells.item[k,6])='/' then Parameters.ParamByName('K_L').Value:=0 else Parameters.ParamByName('K_L').Value:=E_Sheet.cells[k,6].value;		//͸ˮ��
            except Del_List.Add(E_Sheet.Name+#9+'͸ˮ��'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('K_s').Value:=E_Sheet.cells[k,7].value;		//ˮ�ұ���
            except Del_List.Add(E_Sheet.Name+#9+'ˮ�ұ���'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('K_e').Value:=E_Sheet.cells[k,8].value;		//ˮ�ұ���
            except Del_List.Add(E_Sheet.Name+#9+'ˮ�ұ���'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('Z_S').Value:=E_Sheet.cells[k,9].value;		//ע������
            except Del_List.Add(E_Sheet.Name+#9+'ע������'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('Z_E').Value:=E_Sheet.cells[k,10].value;		//ע������
            except Del_List.Add(E_Sheet.Name+#9+'ע������'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_a').Value:=E_Sheet.cells[k,11].value;		//ˮ������
            except Del_List.Add(E_Sheet.Name+#9+'����'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_b').Value:=E_Sheet.cells[k,12].value;		//ˮ������
            except Del_List.Add(E_Sheet.Name+#9+'����'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_c').Value:=E_Sheet.cells[k,13].value;		//ˮ������
            except Del_List.Add(E_Sheet.Name+#9+'�ϻ�'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_v').Value:=E_Sheet.cells[k,14].value;		//ˮ������
            except Del_List.Add(E_Sheet.Name+#9+'�ϼƻ�'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_m').Value:=E_Sheet.cells[k,15].value;		//��λע����
            except Del_List.Add(E_Sheet.Name+#9+'��λע����'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('V_p').Value:=E_Sheet.cells[k,16].value;		//�ཬѹ��
            except Del_List.Add(E_Sheet.Name+#9+'ѹ��'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('T_S').Value:=E_Sheet.cells[k,17].value;		//ʱ����
            except Del_List.Add(E_Sheet.Name+#9+'��ʱ'+#9+inttostr(k)) end;
        		try Parameters.ParamByName('T_E').Value:=E_Sheet.cells[k,18].value;	 	//ʱ����
            except Del_List.Add(E_Sheet.Name+#9+'��ʱ'+#9+inttostr(k)) end;
         		try Parameters.ParamByName('T_T').Value:=E_Sheet.cells[k,19].value;	 	//S ʱ��
            except Del_List.Add(E_Sheet.Name+#9+'ʱ��'+#9+inttostr(k)) end;
        		Parameters.ParamByName('K_B').Value:=E_Sheet.cells[k,20].value;		//T��ע
            ExecSQL;                    //�����ݼ���ִ��SQL���
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
     E_App.WorkBooks.Close;         //�رչ�����
     E_App.Quit;                    //�˳� Excel
     E_App:=Unassigned;             //�ͷ�excel����
     DATA_Connect(StatusBar1.Panels[0].Text,'Caran_GJL',True);           //�������ݿ�
     StatusBar1.Panels[1].Text:=ListBox2.Items[j];
    except;
   	end;
  ListBox2.Width:=Form1.Width-320;
  Memo1.Lines:=List;
  Memo2.Lines:=Del_List;
  Button4.Enabled:=False;
//  Memo1.Visible:=True;
//  Memo2.Visible:=True;
  ShowMessage(IntToStr(List.Count)+'���ļ�ȫ����ɣ�');
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
   DATA_Connect(OpenDialog1.FileName,'Caran_GJL',True);           //�������ݿ�
   Data_EhH(DbgEh_S);         //����������ʾ��ʽ

   F_Name:=Copy(OpenDialog1.FileName,1,Length(OpenDialog1.FileName)-3)+'TXT';
   AssignFile(F,F_Name);
//   try Append(f) except Rewrite(F) end;  //�½��ļ�������Ѵ�����׷��,�����½�
   if not FileExists(F_Name) then Rewrite(F) else Append(F);   //�򿪴��ڵ��ļ������ļ�ָ�붨λ���ļ�β��
   CloseFile(F);
   TT:='Synopsis';
   Button1.Enabled:=True;
 //  Button4.Enabled:=True;
   Button5.Enabled:=True;
   ListBox1.Enabled:=True;
   ListBox2.Enabled:=True;
   StatusBar1.Panels[2].Text:=F_Name;
   StatusBar1.Panels[0].Text:=OpenDialog1.FileName;  
 except ShowMessage('�����ݿ�ʧ��!');
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
//   try Append(f) except Rewrite(F) end;  //�½��ļ�������Ѵ�����׷��,�����½�
   if not FileExists(Trim(SaveDialog1.FileName)+'.TXT') then Rewrite(F) else Append(F);   //�򿪴��ڵ��ļ������ļ�ָ�붨λ���ļ�β��
   CloseFile(F);
   TT:='Synopsis';
   Button1.Enabled:=True;
   Button4.Enabled:=True;
   Button5.Enabled:=True;
   ListBox1.Enabled:=True;
   ListBox2.Enabled:=True;
   StatusBar1.Panels[0].Text:=Trim(SaveDialog1.FileName)+'.Mdb';     //��ȡ����·��
   StatusBar1.Panels[2].Text:=Trim(SaveDialog1.FileName)+'.TXT';
 except ShowMessage('�½����ݿ�ʧ��!');
 end;
end;

procedure TForm1.Button5Click(Sender: TObject);
var Str_Q:string;
begin
 if TT='Synopsis' then           //���
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

  Str_Q:='�׺�,׮��,����,����,�׿ڸ߳�,��������,�ཬ����,�ཬ����,�ཬ�γ�,ƽ��͸ˮ��,ˮ��ע����,ˮ��ע����,ˮ�������,ˮ��ϼ�,��λע����,��ע from Caran_T where �׺�';
  DATA_re(Qur_S,DbgEh_S,DS_S,Str_Q);    //ˢ������
  Data_AdoR(Qur_S);         //����������ʾ��ʽ
  Data_EhR(DbgEh_S);         //����������ʾ��ʽ
  TT:='Statistics';
  StatusBar1.Panels[1].Text:='�ཬ�׼��';
 end else if TT='Statistics' then          //ͳ��
 begin
  Str_Q:='�׺�,�δ�,�ཬ����,�ཬ����,�ཬ�γ�,�׾�,͸ˮ��,ˮ�ұ���,ˮ�ұ���,ע������,ע������,ˮ��ע����,ˮ��ע����,ˮ�������,ˮ��ϼ�,��λע����,�ཬѹ��,ʱ����,ʱ����,ʱ��,��ע from Caran_GJT where �׺�';
  DATA_re(Qur_S,DbgEh_S,DS_S,Str_Q);    //ˢ������
  Data_Ado(Qur_S);         //����������ʾ��ʽ
  Data_EhA(DbgEh_S);         //����������ʾ��ʽ
  StatusBar1.Panels[1].Text:='�ཬ��ͳ������';

  PopupMenu1.Items[2].Visible:=True;
  PopupMenu1.Items[3].Visible:=True;
  PopupMenu1.Items[4].Visible:=True;
  PopupMenu1.Items[5].Visible:=True;
  TT:='ALL';
 end else if TT='ALL' then
 begin
  Str_Q:='�׺�,�δ�,�ཬ����,�ཬ����,�ཬ�γ�,�׾�,͸ˮ��,ˮ�ұ���,ˮ�ұ���,ע������,ע������,ˮ��ע����,ˮ��ע����,ˮ�������,ˮ��ϼ�,��λע����,�ཬѹ��,ʱ����,ʱ����,ʱ��,��ע from Caran_GJL where �׺�';
  DATA_re(Qur_S,DbgEh_S,DS_S,Str_Q);    //ˢ������
  Data_Ado(Qur_S);         //����������ʾ��ʽ
  Data_EhA(DbgEh_S);         //����������ʾ��ʽ
  StatusBar1.Panels[1].Text:='�ཬ����ϸ����';
  PopupMenu1.Items[2].Visible:=True;
  PopupMenu1.Items[3].Visible:=True;
  PopupMenu1.Items[4].Visible:=True;
  PopupMenu1.Items[5].Visible:=True;
  TT:='Synopsis';
 end;
end;

procedure TForm1.N1Click(Sender: TObject);       //��ӡ���
begin
//  PrintDBGridEh2.Title.Text:='��ϸ��';

  PrintDBGridEh2.PageHeader.CenterText.Clear;
  PrintDBGridEh2.PageHeader.CenterText.Add(trim(StatusBar1.Panels[1].Text));    //
  PrintDBGridEh2.PageHeader.Font.Style:=[fsBold];
  PrintDBGridEh2.PageHeader.Font.Name:='����';
  PrintDBGridEh2.PageHeader.Font.Size:=12;

  PrintDBGridEh2.PageFooter.CenterText.Clear;
  PrintDBGridEh2.PageFooter.CenterText.Add('�� &[Page] ҳ / �� &[Pages] ҳ');
  PrintDBGridEh2.PageFooter.RightText.Add(SysUtils.DateTimeToStr(Now()));
  PrintDBGridEh2.PageFooter.Font.Size:=7;
  PrintDBGridEh2.Preview; //��ӡԤ��
  //PrintDBGridEh2.Print; //ֱ�����͵���ӡ���ϴ�ӡ
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
      Application.MessageBox(PChar('û�пɵ���������'), PChar('��ʾ'), MB_OK + MB_ICONINFORMATION);
      exit;
    end;
    FSaveDialog := TSaveDialog.Create(Self);
    FSaveDialog.Filter :=
      'Excel �ĵ� (*.xls)|*.XLS|Text files (*.txt)|*.TXT|Comma separated values (*.csv)|*.CSV|HTML file (*.htm)|*.HTM|Word �ĵ� (*.rtf)|*.RTF';
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
          if application.MessageBox('�ļ����Ѵ��ڣ��Ƿ񸲸�   ', '��ʾ',
            MB_ICONASTERISK or MB_OKCANCEL) <> idok then
            exit;
        end;
        Screen.Cursor := crHourGlass;
        SaveDBGridEhToExportFile(ExpClass, DbgEh_S, FSaveDialog.FileName, true);
        Screen.Cursor := crDefault;
        MessageBox(Handle, '�����ɹ�  ', '��ʾ', MB_OK +
          MB_ICONINFORMATION);
      end;
    end;
    FSaveDialog.Destroy;
  except
    on e: exception do
    begin
      Application.MessageBox(PChar(e.message), '����', MB_OK + MB_ICONSTOP);
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

  DBChart1.Title.Text.Add('ѹˮ���(m)----����ֵ(Lu)');
  DBChart1.Series[0].XValues.ValueSource:='�ཬ����';
  if StatusBar1.Panels[1].Text='�ཬ�׼��' then
    DBChart1.Series[0].YValues.ValueSource:='ƽ��͸ˮ��'
  else DBChart1.Series[0].YValues.ValueSource:='͸ˮ��';
//  DbChart1.Series[0].XLabelsSource:='͸ˮ��';       //����X��
  DbChart1.LeftAxis.Title.Caption:='����ֵ Lu';
  DBChart1.RefreshDataSet(Qur_S,DBChart1.Series[0]);
end;

procedure TForm1.N3Click(Sender: TObject);
begin
  DbgEh_S.Anchors:=[akLeft,akTop,akRight];
  DbgEh_S.Height:=170;
  DbChart1.Visible:=True;
  DBChart1.Series[0].DataSource:=Qur_S;

  DbChart1.Title.Text.Clear;
  DBChart1.Title.Text.Add('ѹˮ���(m)----��λ�ཬ��(kg/m)');
  DBChart1.Series[0].XValues.ValueSource:='�ཬ����';
  DBChart1.Series[0].YValues.ValueSource:='��λע����';
//  DbChart1.Series[0].XLabelsSource:='��λע����';
  DbChart1.LeftAxis.Title.Caption:='��λע���� kg/m';
  DBChart1.RefreshDataSet(Qur_S,DBChart1.Series[0]);
end;

procedure TForm1.N4Click(Sender: TObject);
begin
  DbgEh_S.Anchors:=[akLeft,akTop,akRight];
  DbgEh_S.Height:=170;
  DbChart1.Visible:=True;
  DBChart1.Series[0].DataSource:=Qur_S;

  DbChart1.Title.Text.Clear;
  DBChart1.Title.Text.Add('����ֵ(Lu)----��λ�ཬ��(kg/m)');
  if StatusBar1.Panels[1].Text='�ཬ�׼��' then
    DBChart1.Series[0].YValues.ValueSource:='ƽ��͸ˮ��'
  else  DBChart1.Series[0].XValues.ValueSource:='͸ˮ��';
  
 DBChart1.Series[0].YValues.ValueSource:='��λע����';
//  DbChart1.Series[0].XLabelsSource:='��λע����';
  DbChart1.LeftAxis.Title.Caption:='��λע���� kg/m';
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
var maxR,j,i,k: Integer;//��¼���ݱ�ĵ�ǰ��¼��
  E_Target,E_TTarg: OleVariant;
begin
  opendialog1.filter:='Excel(*.xls)|*.xls|Excel2007(*.xlsx)|*.xlsx';
 if opendialog1.Execute then
 try
   E_App := CreateOleObject('Excel.Application');
   E_App.Visible := true;                       //��ʾ�򿪵�Excel�ļ�
   E_App.WorkBooks.Open(OpenDialog1.FileName);   //���ļ�
   E_Sheet:=E_App.worksheets['Caran_GJL'];
   E_Target:=E_App.worksheets['Caran_GJT'];
   E_TTarg:=E_App.worksheets['Caran_T'];
//   E_Sheet.activate;
   if E_Target.usedrange.rows.count=1 then
    maxR:= E_Sheet.usedrange.rows.count     //�����ݵ�����
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
     E_Target.cells[j,2].value:=E_Sheet.cells[i,2].value;		//�׺�
     E_Target.cells[j,3].value:=E_Sheet.cells[i,3].value;		//�δ�
     E_Target.cells[j,4].value:=E_Sheet.cells[i,4].value;		//�ཬ����
     E_Target.cells[j,5].value:=E_Sheet.cells[i,5].value;		//�ཬ����
     E_Target.cells[j,6].value:=E_Sheet.cells[i,6].value;		//�ཬ�γ�
     E_Target.cells[j,7].value:=E_Sheet.cells[i,7].value;		//�׾�
     try if strtofloat(E_Target.cells[j,8].value)=0 then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value    	//͸ˮ��
     except if Trim(E_Target.cells[j,8].value)='' then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value end;
     try if strtofloat(E_Target.cells[j,9].value)=0 then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value     //ˮ�ұ���
     except if Trim(E_Target.cells[j,9].value)='' then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value end;
     E_Target.cells[j,10].value:=E_Sheet.cells[i,10].value;		//ˮ�ұ���
     try if strtofloat(E_Target.cells[j,11].value)=0 then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value    //ע������
     except if Trim(E_Target.cells[j,11].value)='' then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value end;
     E_Target.cells[j,12].value:=E_Sheet.cells[i,12].value;		//ע������
     try if strtofloat(E_Target.cells[j,13].value)=0 then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value		//ˮ������
     except if Trim(E_Target.cells[j,13].value)='' then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,14].value)=0 then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value		//ˮ������
     except if Trim(E_Target.cells[j,14].value)='' then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,15].value)=0 then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value		//��λע����
     except if Trim(E_Target.cells[j,15].value)='' then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,16].value)=0 then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value		//ˮ������
     except if Trim(E_Target.cells[j,16].value)='' then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,17].value)=0 then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value		//ˮ������
     except if Trim(E_Target.cells[j,17].value)='' then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value end;		//��λע����
     E_Target.cells[j,18].value:=E_Sheet.cells[i,18].value;		//�ཬѹ��
     try if strtofloat(E_Target.cells[j,19].value)=0 then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value    	//ʱ����
     except if Trim(E_Target.cells[j,19].value)='' then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value end;
     E_Target.cells[j,20].value:=E_Sheet.cells[i,20].value;		//ʱ����
     try if strtofloat(E_Target.cells[j,21].value)=0 then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value   	//ʱ��
     except if Trim(E_Target.cells[j,21].value)='' then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value end;
     E_Target.cells[j,22].value:=E_Sheet.cells[i,22].value;		//��ע
    j:=j+1;
   end else
  begin if E_Sheet.cells[i,3].value<>E_Sheet.cells[i+1,3].value then
    begin
     E_Target.cells[j,1].value:=E_Sheet.cells[i,1].value;		//ID
     E_Target.cells[j,2].value:=E_Sheet.cells[i,2].value;		//�׺�
     E_Target.cells[j,3].value:=E_Sheet.cells[i,3].value;		//�δ�
     E_Target.cells[j,4].value:=E_Sheet.cells[i,4].value;		//�ཬ����
     E_Target.cells[j,5].value:=E_Sheet.cells[i,5].value;		//�ཬ����
     E_Target.cells[j,6].value:=E_Sheet.cells[i,6].value;		//�ཬ�γ�
     E_Target.cells[j,7].value:=E_Sheet.cells[i,7].value;		//�׾�
     try if strtofloat(E_Target.cells[j,8].value)=0 then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value    	//͸ˮ��
     except if Trim(E_Target.cells[j,8].value)='' then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value end;
     try if strtofloat(E_Target.cells[j,9].value)=0 then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value     //ˮ�ұ���
     except if Trim(E_Target.cells[j,9].value)='' then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value end;
     E_Target.cells[j,10].value:=E_Sheet.cells[i,10].value;		//ˮ�ұ���
     try if strtofloat(E_Target.cells[j,11].value)=0 then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value    //ע������
     except if Trim(E_Target.cells[j,11].value)='' then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value end;
     E_Target.cells[j,12].value:=E_Sheet.cells[i,12].value;		//ע������
     try if strtofloat(E_Target.cells[j,13].value)=0 then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value		//ˮ������
     except if Trim(E_Target.cells[j,13].value)='' then E_Target.cells[j,13].value:=E_Sheet.cells[i,13].value else E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,14].value)=0 then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value		//ˮ������
     except if Trim(E_Target.cells[j,14].value)='' then E_Target.cells[j,14].value:=E_Sheet.cells[i,14].value else E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,15].value)=0 then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value		//��λע����
     except if Trim(E_Target.cells[j,15].value)='' then E_Target.cells[j,15].value:=E_Sheet.cells[i,15].value else E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,16].value)=0 then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value		//ˮ������
     except if Trim(E_Target.cells[j,16].value)='' then E_Target.cells[j,16].value:=E_Sheet.cells[i,16].value else E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value end;		//ˮ������
     try if strtofloat(E_Target.cells[j,17].value)=0 then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value		//ˮ������
     except if Trim(E_Target.cells[j,17].value)='' then E_Target.cells[j,17].value:=E_Sheet.cells[i,17].value else E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value end;		//��λע����
     E_Target.cells[j,18].value:=E_Sheet.cells[i,18].value;		//�ཬѹ��
     try if strtofloat(E_Target.cells[j,19].value)=0 then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value    	//ʱ����
     except if Trim(E_Target.cells[j,19].value)='' then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value end;
     E_Target.cells[j,20].value:=E_Sheet.cells[i,20].value;		//ʱ����
     try if strtofloat(E_Target.cells[j,21].value)=0 then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value   	//ʱ��
     except if Trim(E_Target.cells[j,21].value)='' then E_Target.cells[j,21].value:=E_Sheet.cells[i,21].value else E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value end;
     E_Target.cells[j,22].value:=E_Sheet.cells[i,22].value;		//��ע
     j:=j+1;
    end else
   begin
    E_Target.cells[j,1].value:=E_Sheet.cells[i,1].value;		//ID
    E_Target.cells[j,2].value:=E_Sheet.cells[i,2].value;		//�׺�
    E_Target.cells[j,3].value:=E_Sheet.cells[i,3].value;		//�δ�
    E_Target.cells[j,4].value:=E_Sheet.cells[i,4].value;		//�ཬ����
    E_Target.cells[j,5].value:=E_Sheet.cells[i,5].value;		//�ཬ����
    E_Target.cells[j,6].value:=E_Sheet.cells[i,6].value;		//�ཬ�γ�
    E_Target.cells[j,7].value:=E_Sheet.cells[i,7].value;		//�׾�
     try if strtofloat(E_Target.cells[j,8].value)=0 then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value    	//͸ˮ��
     except if Trim(E_Target.cells[j,8].value)='' then E_Target.cells[j,8].value:=E_Sheet.cells[i,8].value end;
     try if strtofloat(E_Target.cells[j,9].value)=0 then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value     //ˮ�ұ���
     except if Trim(E_Target.cells[j,9].value)='' then E_Target.cells[j,9].value:=E_Sheet.cells[i,9].value end;
     E_Target.cells[j,10].value:=E_Sheet.cells[i,10].value;		//ˮ�ұ���
     try if strtofloat(E_Target.cells[j,11].value)=0 then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value    //ע������
     except if Trim(E_Target.cells[j,11].value)='' then E_Target.cells[j,11].value:=E_Sheet.cells[i,11].value end;
    E_Target.cells[j,12].value:=E_Sheet.cells[i,12].value;		//ע������
    E_Target.cells[j,13].value:=E_Target.cells[j,13].value+E_Sheet.cells[i,13].value;		//ˮ������
    E_Target.cells[j,14].value:=E_Target.cells[j,14].value+E_Sheet.cells[i,14].value;		//ˮ������
    E_Target.cells[j,15].value:=E_Target.cells[j,15].value+E_Sheet.cells[i,15].value;		//ˮ������
    E_Target.cells[j,16].value:=E_Target.cells[j,16].value+E_Sheet.cells[i,16].value;		//ˮ������
    E_Target.cells[j,17].value:=E_Target.cells[j,17].value+E_Sheet.cells[i,17].value;		//��λע����
    E_Target.cells[j,18].value:=E_Sheet.cells[i,18].value;		//�ཬѹ��
     try if strtofloat(E_Target.cells[j,19].value)=0 then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value    	//ʱ����
     except if Trim(E_Target.cells[j,19].value)='' then E_Target.cells[j,19].value:=E_Sheet.cells[i,19].value end;
    E_Target.cells[j,20].value:=E_Sheet.cells[i,20].value;		//ʱ����
    E_Target.cells[j,21].value:=E_Target.cells[j,21].value+E_Sheet.cells[i,21].value;		//ʱ��
    E_Target.cells[j,22].value:=E_Sheet.cells[i,22].value;		//��ע
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
    E_TTarg.cells[j,2].value:=E_Target.cells[i,2].value;		//�׺�
    try if strtofloat(E_TTarg.cells[j,8].value)=0 then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value     //�����
    except if Trim(E_TTarg.cells[j,8].value)='' then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value end;
    E_TTarg.cells[j,9].value:=E_Target.cells[i,5].value;		//�չ���
    try if strtofloat(E_TTarg.cells[j,10].value)=0 then E_TTarg.cells[j,10].value:=E_Target.cells[i,6].value else E_TTarg.cells[j,10].value:=E_TTarg.cells[j,10].value+E_Target.cells[i,6].value		//�γ�
    except if Trim(E_TTarg.cells[j,10].value)='' then E_TTarg.cells[j,10].value:=E_Target.cells[i,6].value else E_TTarg.cells[j,10].value:=E_TTarg.cells[j,10].value+E_Target.cells[i,6].value end;
    try if strtofloat(E_TTarg.cells[j,11].value)=0 then E_TTarg.cells[j,11].value:=E_Target.cells[i,8].value else E_TTarg.cells[j,11].value:=(E_TTarg.cells[j,11].value+E_Target.cells[i,8].value)/k		//ƽ��͸ˮ��
    except if Trim(E_TTarg.cells[j,11].value)='' then E_TTarg.cells[j,11].value:=E_Target.cells[i,8].value else E_TTarg.cells[j,11].value:=(E_TTarg.cells[j,11].value+E_Target.cells[i,8].value)/k end;
    try if strtofloat(E_TTarg.cells[j,12].value)=0 then E_TTarg.cells[j,12].value:=E_Target.cells[i,13].value else E_TTarg.cells[j,12].value:=E_TTarg.cells[j,12].value+E_Target.cells[i,13].value		//ˮ������
    except if Trim(E_TTarg.cells[j,12].value)='' then E_TTarg.cells[j,12].value:=E_Target.cells[i,13].value else E_TTarg.cells[j,12].value:=E_TTarg.cells[j,12].value+E_Target.cells[i,13].value end;
    try if strtofloat(E_TTarg.cells[j,13].value)=0 then E_TTarg.cells[j,13].value:=E_Target.cells[i,14].value else E_TTarg.cells[j,13].value:=E_TTarg.cells[j,13].value+E_Target.cells[i,14].value		//ˮ������
    except if Trim(E_TTarg.cells[j,13].value)='' then E_TTarg.cells[j,13].value:=E_Target.cells[i,14].value else E_TTarg.cells[j,13].value:=E_TTarg.cells[j,13].value+E_Target.cells[i,14].value end;
    try if strtofloat(E_TTarg.cells[j,14].value)=0 then E_TTarg.cells[j,14].value:=E_Target.cells[i,15].value else E_TTarg.cells[j,14].value:=E_TTarg.cells[j,14].value+E_Target.cells[i,15].value		//ˮ������
    except if Trim(E_TTarg.cells[j,14].value)='' then E_TTarg.cells[j,14].value:=E_Target.cells[i,15].value else E_TTarg.cells[j,14].value:=E_TTarg.cells[j,14].value+E_Target.cells[i,15].value end;
    try if strtofloat(E_TTarg.cells[j,15].value)=0 then E_TTarg.cells[j,15].value:=E_Target.cells[i,16].value else E_TTarg.cells[j,15].value:=E_TTarg.cells[j,15].value+E_Target.cells[i,16].value		//ˮ������
    except if Trim(E_TTarg.cells[j,15].value)='' then E_TTarg.cells[j,15].value:=E_Target.cells[i,16].value else E_TTarg.cells[j,15].value:=E_TTarg.cells[j,15].value+E_Target.cells[i,16].value end;
    try if strtofloat(E_TTarg.cells[j,16].value)=0 then E_TTarg.cells[j,16].value:=E_Target.cells[i,17].value else E_TTarg.cells[j,16].value:=E_TTarg.cells[j,13].value/E_TTarg.cells[j,10].value		//��λע����
    except if Trim(E_TTarg.cells[j,16].value)='' then E_TTarg.cells[j,16].value:=E_Target.cells[i,17].value else E_TTarg.cells[j,16].value:=E_TTarg.cells[j,13].value/E_TTarg.cells[j,10].value end;
    E_TTarg.cells[j,17].value:=IntToStr(k)+'��';		//��ע
    j:=j+1; k:=1;
    end else
   begin
    E_TTarg.cells[j,1].value:=E_Target.cells[i,1].value;		//ID
    E_TTarg.cells[j,2].value:=E_Target.cells[i,2].value;		//�׺�
    try if strtofloat(E_TTarg.cells[j,8].value)=0 then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value     //�����
    except if Trim(E_TTarg.cells[j,8].value)='' then E_TTarg.cells[j,8].value:=E_Target.cells[i,4].value end;
    E_TTarg.cells[j,9].value:=E_Target.cells[i,5].value;		//�չ���
    E_TTarg.cells[j,10].value:=E_TTarg.cells[j,10].value+E_Target.cells[i,6].value;		//�γ�  *
    E_TTarg.cells[j,11].value:=E_TTarg.cells[j,11].value+E_Target.cells[i,8].value;		//ƽ��͸ˮ��
    E_TTarg.cells[j,12].value:=E_TTarg.cells[j,12].value+E_Target.cells[i,13].value;		//ƽ��͸ˮ��
    E_TTarg.cells[j,13].value:=E_TTarg.cells[j,13].value+E_Target.cells[i,14].value;		//ƽ��͸ˮ��
    E_TTarg.cells[j,14].value:=E_TTarg.cells[j,14].value+E_Target.cells[i,15].value;		//ƽ��͸ˮ��
    E_TTarg.cells[j,15].value:=E_TTarg.cells[j,15].value+E_Target.cells[i,16].value;		//ƽ��͸ˮ��
    E_TTarg.cells[j,16].value:=E_TTarg.cells[j,16].value+E_Target.cells[i,17].value;		//ƽ��͸ˮ��
    k:=k+1;
   end;
 end;
 E_App.ActiveWorkBook.Save;
 E_App.WorkBooks.Close;         //�رչ�����
 E_App.Quit;                    //�˳� Excel
 E_App:=Unassigned;             //�ͷ�excel����
 ShowMessage('����������ɣ�');
end;

end.
