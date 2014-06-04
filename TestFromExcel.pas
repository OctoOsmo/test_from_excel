unit TestFromExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, IBDatabase, DB, IBCustomDataSet, IBQuery, StdCtrls, ShellAPI,
  IBUpdateSQL;


type
  TFormMain = class(TForm)
    IBQueryImport: TIBQuery;
    IBDatabaseImport: TIBDatabase;
    IBTransactionImport: TIBTransaction;
    ButtonOpen: TButton;
    LabelCurrentQuestion: TLabel;
    Memo1: TMemo;
    LabelFileName: TLabel;
    LabelFilePath: TLabel;
    IBUpdateSQLQuestion: TIBUpdateSQL;
    IBQueryQuestion: TIBQuery;
    DataSourceQuestion: TDataSource;
    IBQueryQuestionN_VOPR: TIntegerField;
    IBQueryQuestionN_TEMA: TIntegerField;
    IBQueryQuestionNAME_VOPR_L: TIBStringField;
    IBQueryQuestionNAME_VOPR_P: TIBStringField;
    IBQueryQuestionTIP_VOPR: TSmallintField;
    IBQueryQuestionKOL_OTV: TIntegerField;
    IBQueryQuestionVOPR_PIC: TBlobField;
    Button1: TButton;
    IBTransactionQuestion: TIBTransaction;
    procedure ButtonOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  protected
    procedure WMDropFiles (var Msg: TMessage);
    message wm_DropFiles;
  public
    { Public declarations }
  end;

var
  FormMain: TFormMain;
  Excel: Variant;
  filePath: string;
  themeName: string;

implementation

{$R *.dfm}

{function LastPos(SubStr, S: string): Integer;
var
 Found, Len, Pos: integer;
begin
 Pos := Length(S);
 Len := Length(SubStr);
 Found := 0;
 while (Pos > 0) and (Found = 0) do
 begin
   if Copy(S, Pos, Len) = SubStr then
     Found := Pos;
   Dec(Pos);
 end;
 LastPos := Found;
end;}

procedure TFormMain.WMDropFiles(var Msg: TMessage);
Var
Filename: array[0..256] of char;
begin 

DragQueryFile(THandle(Msg.WParam),0,Filename,SizeOf(Filename));
// ������ � ���������� Filename ����� ���������� ���� 
// � ���������������� �����. ����� �� ������ ��������� � ���� ������, ���� 
// ��� ����, ��� ��� ������. 

//��������: ��������� ��� � Memo 
Memo1.lines.loadfromfile(Filename);

//����� ����� �����
filePath := Filename;
LabelFilePath.Caption := '���� � �����: '+filePath;
themeName := copy(filePath, LastDelimiter('/\', filePath)+1 ,Length(filePath));
themeName := copy(themeName, 1 , LastDelimiter('.', themeName)-1);
LabelFileName.Caption := '��� �������� �����: '+themeName;
FormMain.Width := FormMain.LabelFilePath.Width + FormMain.LabelFilePath.Left + 20;

//�������� �� ��������� ������������� 
DragFinish(THandle(Msg.WParam));
end;

procedure SetSQL(q: TIBQuery; sqlText: string);
begin
q.Close;
q.SQL.Clear;
q.SQL.Append(sqlText);
//q.Open;
end;

function AddTheme(themeName: string): integer;
var
  sqlText: string;
begin
//INSERT INTO TEMA (NAME_TEMA, KOL_VOPR) VALUES ('NewTheme', 0)
  sqlText := 'INSERT INTO TEMA (NAME_TEMA, KOL_VOPR) VALUES ('''+themeName+''', 0)';
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.ExecSQL;
  sqlText := 'select max(N_TEMA) from TEMA';
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.Open;
  Result := FormMain.IBQueryImport.FieldByName('MAX').AsInteger;
end;

function AddQuestion(themeNum: integer; questionText: string): integer;
var
  sqlText: string;
begin
  sqlText := 'insert into vopros (n_tema, name_vopr_l, name_vopr_p, tip_vopr, kol_otv) values('+IntToStr(themeNum)+', '''+questionText+''', '''', 1, 0)';
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.ExecSQL;
  sqlText := 'select max(N_VOPR) from VOPROS';
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.Open;
  Result := FormMain.IBQueryImport.FieldByName('MAX').AsInteger;

 //sqlText := 'INSERT INTO TEMA (NAME_TEMA, KOL_VOPR) VALUES ('''+themeName+''', 0)';
end;

function AddAnswer(questionNum: integer; answer: string; isRight: integer): integer;
var
  sqlText: string;
begin
  sqlText := 'insert into otvet (n_vopr, name_otv, pr_pr) values ('+IntToStr(questionNum)+', '''+answer+''', '+IntToStr(isRight)+')';
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.ExecSQL;
  sqlText := 'select max(n_otv) from otvet';  
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.Open;
  Result := FormMain.IBQueryImport.FieldByName('MAX').AsInteger;
end;

procedure SetAnsCout(questionNum: integer; ansCount: integer);
var  
  sqlText: string;
begin
  sqlText:= 'update vopros v set v.kol_otv ='+IntToStr(ansCount)+' where v.n_vopr = '+IntToStr(questionNum)+'';
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.ExecSQL;
end;

procedure SetQuestCout(themeNum: integer; questCount: integer);
var  
  sqlText: string;
begin
  sqlText:= 'update tema t set t.kol_vopr ='+IntToStr(questCount)+' where t.n_tema = '+IntToStr(themeNum)+'';
  SetSQL(FormMain.IBQueryImport, sqlText);
  FormMain.IBQueryImport.ExecSQL;
end;

procedure TFormMain.ButtonOpenClick(Sender: TObject);
var
 Question, tmp: string;
 colCount, i, j, curRow, curCol, questionNum, themeNum, questCount, ansCount, isRight: integer;
 ansArray : Array[1..20] of string;
begin
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[filePath];//, 0, True];
  //Excel.Visible := True;
  //�������������� ������ � ������ ��� ����
  FormMain.IBDatabaseImport.Connected := true;

  //�������� ����� ���� ������������
  themeNum := AddTheme(themeName);
  questCount := 0;
  curCol := 2;
  curRow := 2;
  Question := Trim(Excel.Cells[curRow, curCol]);
  while('' <> Question) do
  begin
    //--------------------------

    //--------------------------
    //���� � ���� ������ �������
    i:=0;
    questionNum := AddQuestion(themeNum, Question);
    questCount := questCount+1;
    //���� ��������� ������� � �������� �������
    //������������ ������� �������
    ansCount := 0;
    repeat
    begin
      i:=i+1;//counting size of answers array
      curCol := curCol+1;//������� � ���������� ������
      ansArray[i] := Trim(Excel.Cells[curRow, curCol]);
    end
    until('' = ansArray[i]);
    //���� ������� � ���� ������
    i:= 1;
    while('' <> ansArray[i]) do
    begin
      if('+' = ansArray[i][1]) then
        begin
          isRight := 1;
          ansArray[i] := Trim(copy(ansArray[i], 2, Length(ansArray[i])-1));
        end
        else
          isRight := 0;
      AddAnswer(questionNum, ansArray[i], isRight);
      ansCount := ansCount+1;//counting answers
      i:= i+1;
    end;
    SetAnsCout(questionNum, ansCount);//��������� ���������� ������� ��� �������
    //������� � ���������� �������
    curRow := curRow+1;
    curCol := 2;
    Question := Trim(Excel.Cells[curRow, curCol]);
    FormMain.LabelCurrentQuestion.Caption := '����� �������� �������: ' + IntToStr(questCount);
    Application.ProcessMessages;
  end;//end of theme
  SetQuestCout(themeNum, questCount);//��������� ���������� �������� � ����
  //�������� Excel
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;
  FormMain.IBTransactionImport.CommitRetaining;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  DragAcceptFiles(Handle,True);
end;

procedure TFormMain.Button1Click(Sender: TObject);
begin
  IBQueryQuestion.Open;
  IBQueryQuestion.Insert;
  IBQueryQuestionN_TEMA.Value := 123;
  IBQueryQuestionTIP_VOPR.Value := 1;
  IBQueryQuestionKOL_OTV.Value := 0;
  IBQueryQuestion.ApplyUpdates;
  IBTransactionQuestion.CommitRetaining;
  ShowMessage(IBQueryQuestionN_VOPR.Text);
  IBQueryQuestion.Close;
end;

end.
