unit TestFromExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, IBDatabase, DB, IBCustomDataSet, IBQuery, StdCtrls, ShellAPI,
  IBUpdateSQL;


type
  TFormMain = class(TForm)
    IBDatabaseImport: TIBDatabase;
    ButtonOpen: TButton;
    LabelCurrentQuestion: TLabel;
    Memo1: TMemo;
    LabelFileName: TLabel;
    LabelFilePath: TLabel;
    IBUpdateSQLQuestion: TIBUpdateSQL;
    IBQueryQuestion: TIBQuery;
    IBQueryQuestionN_VOPR: TIntegerField;
    IBQueryQuestionN_TEMA: TIntegerField;
    IBQueryQuestionNAME_VOPR_L: TIBStringField;
    IBQueryQuestionNAME_VOPR_P: TIBStringField;
    IBQueryQuestionTIP_VOPR: TSmallintField;
    IBQueryQuestionKOL_OTV: TIntegerField;
    IBQueryQuestionVOPR_PIC: TBlobField;
    IBQueryAnswer: TIBQuery;
    IBUpdateSQLAnswer: TIBUpdateSQL;
    IBQueryTheme: TIBQuery;
    IBUpdateSQLTheme: TIBUpdateSQL;
    IBQueryAnswerN_OTV: TIntegerField;
    IBQueryAnswerN_VOPR: TIntegerField;
    IBQueryAnswerNAME_OTV: TIBStringField;
    IBQueryAnswerPR_PR: TSmallintField;
    IBQueryAnswerN_OTV_S: TIntegerField;
    IBQueryAnswerOTV_PIC: TBlobField;
    IBQueryThemeN_TEMA: TIntegerField;
    IBQueryThemeNAME_TEMA: TIBStringField;
    IBQueryThemeKOL_VOPR: TIntegerField;
    IBTransactionTheme: TIBTransaction;
    LabelStatus: TLabel;
    ListBoxFileList: TListBox;
    procedure ButtonOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
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

procedure GetAllFiles( Path: string; Lb: TListBox );
var
sRec: TSearchRec;
isFound: boolean;
begin
isFound := FindFirst( Path + ‘\*.*’, faAnyFile, sRec ) = 0;
while isFound do
begin
if ( sRec.Name <> ‘.’ ) and ( sRec.Name <> ‘..’ ) then
begin
if ( sRec.Attr and faDirectory ) = faDirectory then
GetAllFiles( Path + ‘\’ + sRec.Name, Lb );
Lb.Items.Add( Path + ‘\’ + sRec.Name );
end;
Application.ProcessMessages;
isFound := FindNext( sRec ) = 0;
end;
FindClose( sRec );
end;

procedure TFormMain.WMDropFiles(var Msg: TMessage);
Var
Filename: array[0..256] of char;
begin 

DragQueryFile(THandle(Msg.WParam),0,Filename,SizeOf(Filename));
// Теперь в переменной Filename будет находиться путь 
// к перетаскиваемому файлу. Далее вы можете выполнять с этим файлом, зная 
// его путь, все что угодно.

//Например: Загрузить его в Memo 
//Memo1.lines.loadfromfile(Filename);

//Вывод имени файла
filePath := Filename;
LabelFilePath.Caption := 'Путь к файлу: '+filePath;
themeName := copy(filePath, LastDelimiter('/\', filePath)+1 ,Length(filePath));
themeName := copy(themeName, 1 , LastDelimiter('.', themeName)-1);
LabelFileName.Caption := 'Имя текущего файла: '+themeName;
FormMain.Width := FormMain.LabelFilePath.Width + FormMain.LabelFilePath.Left + 20;

//Сообщаем об окончании претаскивания 
DragFinish(THandle(Msg.WParam));
GetAllFiles(filePath, FormMain.ListBoxFileList);
end;

function AddTheme(themeName: string): integer;
begin
  FormMain.IBQueryTheme.Open;
  FormMain.IBQueryTheme.Insert;
  FormMain.IBQueryThemeNAME_TEMA.Value := themeName;
  FormMain.IBQueryThemeKOL_VOPR.Value := 0;
  Result := FormMain.IBQueryThemeN_TEMA.Value;
end;

function AddQuestion(themeNum: integer; questionText: string): integer;
begin
  FormMain.IBQueryQuestion.Open;
  FormMain.IBQueryQuestion.Insert;
  FormMain.IBQueryQuestionN_TEMA.Value := themeNum;
  FormMain.IBQueryQuestionNAME_VOPR_L.Value := questionText;
  FormMain.IBQueryQuestionKOL_OTV.Value := 0;
  FormMain.IBQueryQuestionTIP_VOPR.Value := 1;
  Result := FormMain.IBQueryQuestionN_VOPR.Value;
end;

function AddAnswer(questionNum: integer; answer: string; isRight: integer): integer;
begin
  FormMain.IBQueryAnswer.Open;
  FormMain.IBQueryAnswer.Insert;
  FormMain.IBQueryAnswerN_VOPR.Value := questionNum;
  FormMain.IBQueryAnswerNAME_OTV.Value := answer;
  FormMain.IBQueryAnswerPR_PR.Value := isRight;
  Result := FormMain.IBQueryAnswerN_OTV.Value;
end;

procedure SetAnsCout(questionNum: integer; ansCount: integer);
begin
  FormMain.IBQueryQuestion.Open;
  FormMain.IBQueryQuestion.Edit;
  FormMain.IBQueryQuestionKOL_OTV.Value := ansCount;
end;

procedure SetQuestCout(themeNum: integer; questCount: integer);
begin
  FormMain.IBQueryTheme.Open;
  FormMain.IBQueryTheme.Edit;
  FormMain.IBQueryThemeKOL_VOPR.Value := questCount;
end;

procedure TFormMain.ButtonOpenClick(Sender: TObject);
var
 Question: string;
 i, curRow, curCol, questionNum, themeNum, questCount, ansCount, isRight: integer;
 ansArray : Array[1..20] of string;
begin
  FormMain.LabelStatus.Caption := 'Статус: обработка...';
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[filePath];//, 0, True];
  //Excel.Visible := True;
  //Подготавливаем вопрос и ответы для базы
  FormMain.IBDatabaseImport.Connected := true;

  //Создание новой темы тестирования
  themeNum := AddTheme(themeName);
  questCount := 0;
  curCol := 2;
  curRow := 2;
  Question := Trim(Excel.Cells[curRow, curCol]);
  while('' <> Question) do
  begin
    //Ввод в тему нового вопроса
    i:=0;
    questionNum := AddQuestion(themeNum, Question);
    questCount := questCount+1;
    //Ввод вариантов ответов к текущему вопросу
    //Формирование массива ответов
    ansCount := 0;
    repeat
    begin
      i:=i+1;//counting size of answers array
      curCol := curCol+1;//Переход к следующему ответу
      ansArray[i] := Trim(Excel.Cells[curRow, curCol]);
    end
    until('' = ansArray[i]);
    //Ввод ответов в базу данных
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
    SetAnsCout(questionNum, ansCount);//Установка количества ответов для вопроса
    //Переход к следующему вопросу
    curRow := curRow+1;
    curCol := 2;
    Question := Trim(Excel.Cells[curRow, curCol]);
    FormMain.LabelCurrentQuestion.Caption := 'Номер текущего вопроса: ' + IntToStr(questCount);
    Application.ProcessMessages;
  end;//end of theme
  SetQuestCout(themeNum, questCount);//Установка количества вопросов в теме
  //Закрытие Excel
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;
  FormMain.IBQueryTheme.ApplyUpdates;
  FormMain.IBQueryQuestion.ApplyUpdates;
  FormMain.IBQueryAnswer.ApplyUpdates;
  FormMain.IBTransactionTheme.CommitRetaining;
  FormMain.IBQueryTheme.Close;
  FormMain.IBQueryQuestion.Close;
  FormMain.IBQueryAnswer.Close;
  FormMain.LabelStatus.Caption := 'Статус: ожидание.';
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  DragAcceptFiles(Handle,True);
end;

end.
