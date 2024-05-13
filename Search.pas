// *** Classe para pesquisas - opção a não utilizar o componente DBLoockupCombo
// *** Analista/Desenvolvedor - Hélio Oliveira
// *** Data: 11/05/2024
// ***    O desenho do formulário tem uma medida padrão 500x600 que poderá ser
// *** alterada quando necessário - as larguras das colunas no grid irá respeitar
// *** o tamanho do campo no SGDB.

unit Search;

interface

uses
  Vcl.Forms, Vcl.ExtCtrls, Vcl.Controls, Vcl.StdCtrls, Vcl.DBGrids, System.Classes,
  Vcl.Graphics, System.SysUtils, System.UITypes, FireDAC.Comp.Client, Data.DB;

type
  TSearch = class
  private
    FJoinTable2: String;
    FFieldsNameList: String;
    FJoinTable3: String;
    FJoinTable1: String;
    FFieldsNameConditionList: String;
    FFieldsDescriptionList: String;
    FJoinFieldsCondition2: String;
    FJoinFieldsCondition3: String;
    FJoinFieldsCondition1: String;
    FMasterTable: String;
    FValuesSearch: String;
    FFieldsOrdering: String;
    FValueReturn: String;
    FSearchTitle: String;
    FQuery : TFDQuery;
    FSource : TDataSource;
    FFormSearsh : TForm;
    FListFieldsName : TStringList;
    FListFieldsDescription : TStringList;
    FListFieldsNameCondition : TStringList;
    FListValuesSearch : TStringList;
    FDBGrid : TDBGrid;
    FPanelTop : TPanel;
    FpanelTop_x : TPanel;
    FPanelButton : TPanel;
    FPanelCenter : TPanel;
    FBtnSelect : TButton;
    FBtnClose : TButton;
    FStrB : TStringBuilder;
    FEditSearch : TEdit;
    FFieldFilter: String;
    procedure SetFieldsNameConditionList(const Value: String);
    procedure SetFieldsNameList(const Value: String);
    procedure SetJoinFieldsCondition1(const Value: String);
    procedure SetJoinFieldsCondition2(const Value: String);
    procedure SetJoinFieldsCondition3(const Value: String);
    procedure SetJoinTable1(const Value: String);
    procedure SetJoinTable2(const Value: String);
    procedure SetJoinTable3(const Value: String);
    procedure SetMasterTable(const Value: String);
    procedure SetValuesSearch(const Value: String);
    procedure SetFieldsDescriptionList(const Value: String);
    procedure DrawFormSearch(aForm : TForm);
    procedure SetGridColumns(aGrid : TDBGrid);
    Procedure GetDados;
    procedure FormSearchShow(Sender : TObject);
    procedure FEditSearchChange(Sender : TObject);
    procedure FBtnCloseClick(Sender : TObject);
    procedure FBtnSelectClick(Sender : TObject);
    procedure SetFieldsOrdering(const Value: String);
    procedure SetValueReturn(const Value: String);
    procedure SetSearchTitle(const Value: String);
    procedure SetFieldFilter(const Value: String);
  public
    procedure StartSearch; //Método que irá invocar a pesquisa - deverá ser executado após a criação da instância de TSearch
    property MasterTable : String read FMasterTable write SetMasterTable; //Nome da tabela principal da consulta
    property JoinTable1 : String read FJoinTable1 write SetJoinTable1; //Tabela de junção
    property JoinFieldsCondition1 : String read FJoinFieldsCondition1 write SetJoinFieldsCondition1; //Condição da junção
    property JoinTable2 : String read FJoinTable2 write SetJoinTable2;
    property JoinFieldsCondition2 : String read FJoinFieldsCondition2 write SetJoinFieldsCondition2;
    property JoinTable3 : String read FJoinTable3 write SetJoinTable3;
    property JoinFieldsCondition3 : String read FJoinFieldsCondition3 write SetJoinFieldsCondition3;
    property FieldsNameList : String read FFieldsNameList write SetFieldsNameList; //Nome do campos que serão listados na grid de consulta (separados por virgula)
    property FieldsDescriptionList : String read FFieldsDescriptionList write SetFieldsDescriptionList; //Titulo das colunas a ser exibida no titulo da grid (separados por virgula)
    property FieldsNameConditionList : String read FFieldsNameConditionList write SetFieldsNameConditionList; //Nome dos campos que farão parte da condição de pesquisa (separados por virgula)
    property FieldsOrdering : String read FFieldsOrdering write SetFieldsOrdering; //Campo(s) pelo(s) qual(ais) a consulta será ordenada (separados por virgula)
    property ValuesSearch : String read FValuesSearch write SetValuesSearch; //Valores a serem pesquisados (separados por virgula)
    property ValueReturn : String read FValueReturn write SetValueReturn; //Valor retornado após a seleção no grid - aqui retornará o valor da coluna 0 (zero)
    property SearchTitle : String read FSearchTitle write SetSearchTitle; //Titulo da tela de pesquisa
    property FieldFilter : String read FFieldFilter write SetFieldFilter; //Campo pelo qual será feito o filtro na tela de consulta.
    constructor Create(Connection : TFDConnection);
    destructor Destroy; override;
  end;

implementation

{ TSearch }

procedure TSearch.FBtnCloseClick(Sender: TObject);
begin
  Self.FFormSearsh.Close;
end;

constructor TSearch.Create(Connection: TFDConnection);
begin
  Self.FQuery                   := TFDQuery.Create(nil);
  Self.FQuery.Connection        := Connection;
  Self.FSource                  := TDataSource.Create(nil);
  Self.FSource.DataSet          := Self.FQuery;
  Self.FFormSearsh              := TForm.Create(nil);
  Self.FFormSearsh.Color        := $00999999;
  Self.FFormSearsh.OnShow       := Self.FormSearchShow;
  Self.FDBGrid                  := TDBGrid.Create(nil);
  Self.FListFieldsName          := TStringList.Create;
  Self.FListFieldsDescription   := TStringList.Create;
  Self.FListFieldsNameCondition := TStringList.Create;
  Self.FListValuesSearch        := TStringList.Create;
  FStrB                         := TStringBuilder.Create;
end;

procedure TSearch.DrawFormSearch(aForm: TForm);
begin
  aForm.ClientHeight               := 400;
  aForm.ClientWidth                := 500;
  aForm.BorderStyle                := bsNone;
  aForm.Position                   := poOwnerFormCenter;
  Self.FPanelTop                   := TPanel.Create(nil);
  Self.FpanelTop_x                 := TPanel.Create(nil);
  Self.FPanelButton                := TPanel.Create(nil);
  Self.FPanelCenter                := TPanel.Create(nil);
  Self.FPanelTop.Font.Color        := clWhite;
  Self.FPanelTop.Parent            := aForm;
  Self.FPanelButton.Parent         := aForm;
  Self.FPanelCenter.Parent         := aForm;
  Self.FpanelTop_x.Parent          := aForm;
  Self.FPanelTop.Align             := alTop;
  Self.FPanelButton.Align          := alBottom;
  Self.FPanelCenter.Align          := alClient;
  Self.FpanelTop_x.Align           := alTop;
  Self.FpanelTop_x.Height          := 25;
  Self.FPanelButton.Padding.Left   := 3;
  Self.FPanelButton.Padding.Right  := 3;
  Self.FPanelButton.Padding.Top    := 3;
  Self.FPanelButton.Padding.Bottom := 3;
  Self.FPanelTop.Caption           := Self.SearchTitle;
  Self.FPanelTop.Height            := 30;
  Self.FPanelTop.Font.Size         := 12;
  Self.FPanelTop.Font.Style        := [TFontStyle.fsBold];
  Self.FPanelTop.Color             := clGray;
  Self.FPanelTop.ParentBackground  := False;
  Self.FEditSearch                 := TEdit.Create(nil);
  Self.FEditSearch.Parent          := Self.FpanelTop_x;
  Self.FEditSearch.Align           := alClient;
  Self.FEditSearch.OnChange        := Self.FEditSearchChange;
  Self.FBtnSelect                  := TButton.Create(nil);
  Self.FBtnClose                   := TButton.Create(nil);
  Self.FBtnSelect.Font.Style       := [TFontStyle.fsBold];
  Self.FBtnClose.Font.Style        := [TFontStyle.fsBold];
  Self.FBtnSelect.OnClick          := Self.FBtnSelectClick;
  Self.FBtnClose.OnClick           := Self.FBtnCloseClick;
  Self.FBtnSelect.Parent           := FPanelButton;
  Self.FBtnClose.Parent            := FPanelButton;
  Self.FBtnSelect.Align            := alRight;
  Self.FBtnClose.Align             := alRight;
  Self.FBtnSelect.Caption          := 'Selecionar';
  Self.FBtnClose.Caption           := 'Cancelar';
  Self.FDBGrid.Parent              := Self.FPanelCenter;
  Self.FDBGrid.Align               := alClient;
  Self.FDBGrid.DataSource          := Self.FSource;
  Self.FDBGrid.Options             := Self.FDBGrid.Options + [dgRowSelect];
  Self.FDBGrid.Options             := Self.FDBGrid.Options - [dgEditing, dgIndicator];

  Self.SetGridColumns(Self.FDBGrid);
  Self.GetDados;
  aForm.ShowModal;
end;

procedure TSearch.FBtnSelectClick(Sender: TObject);
begin
  Self.ValueReturn := Self.FDBGrid.Columns[0].Field.Value;
  Self.FBtnCloseClick(Sender);
end;

procedure TSearch.FEditSearchChange(Sender: TObject);
begin
  Self.FQuery.Filtered      := False;
  Self.FQuery.FilterOptions := [foCaseInsensitive];
  Self.FQuery.Filter        := Self.FieldFilter + ' Like ' +
                               QuotedStr('%' + Self.FEditSearch.Text + '%');
  Self.FQuery.Filtered      := True;
end;

procedure TSearch.FormSearchShow(Sender: TObject);
begin
  Self.FEditSearch.SetFocus;
end;

procedure TSearch.GetDados;
var
  aIndex: Integer;
begin
  Self.FStrB.AppendLine('SELECT ').Append(Self.FFieldsNameList)
       .AppendLine(' FROM ').Append(FMasterTable);

  if not Self.FJoinTable1.IsEmpty then
    Self.FStrB.AppendLine(Self.JoinTable1).Append(Self.FJoinFieldsCondition1);

  if not Self.FieldsNameConditionList.IsEmpty then
  begin

    FListFieldsNameCondition.StrictDelimiter := True;
    FListFieldsNameCondition.Delimiter       := ',';
    FListFieldsNameCondition.DelimitedText   := Self.FieldsNameConditionList;

    FListValuesSearch.StrictDelimiter := True;
    FListValuesSearch.Delimiter       := ',';
    FListValuesSearch.DelimitedText   := Self.FValuesSearch;

    if FListFieldsNameCondition.Count <> FListValuesSearch.Count then
      raise Exception.Create('A quantidade de valores a serem pesquisados devem ser iguais a quantidade de campos [FIELDS] condicionais.');

    Self.FStrB.AppendLine('WHERE ');
    for aIndex := 0 to Pred(Self.FListFieldsNameCondition.Count) do
    begin
      if aIndex = 0 then
        Self.FStrB.Append(Self.FListFieldsNameCondition[aIndex]).Append(' = ').Append(Self.FListValuesSearch[aIndex].QuotedString)
      else
        Self.FStrB.Append(' AND ').Append(Self.FListFieldsNameCondition[aIndex]).Append(' = ').Append(Self.FListValuesSearch[aIndex].QuotedString);
    end;
  end;

  Self.FStrB.AppendLine(Self.FFieldsOrdering).Append(';');

  Self.FQuery.Connection.StartTransaction;
  try
    Self.FQuery.SQL.Text := Self.FStrB.ToString;
    Self.FQuery.Open
  except
    on e : Exception do
    begin
      Self.FQuery.Connection.Rollback;
      raise Exception.Create(e.Message);
    end;
  end;
end;

destructor TSearch.Destroy;
begin
  FreeAndNil(FQuery);
  FreeAndNil(FSource);
  FreeAndNil(FFormSearsh);
  FreeAndNil(FListFieldsName);
  FreeAndNil(FListFieldsNameCondition);
  FreeAndNil(FListValuesSearch);
  FreeAndNil(FListFieldsDescription);
  FreeAndNil(FStrB);
  inherited;
end;

procedure TSearch.SetFieldFilter(const Value: String);
begin
  FFieldFilter := Value;
end;

procedure TSearch.SetFieldsDescriptionList(const Value: String);
begin
  FFieldsDescriptionList := Value;
end;

procedure TSearch.SetFieldsNameConditionList(const Value: String);
begin
  FFieldsNameConditionList := Value;
end;

procedure TSearch.SetFieldsNameList(const Value: String);
begin
  FFieldsNameList := Value;
end;

procedure TSearch.SetFieldsOrdering(const Value: String);
begin
  FFieldsOrdering := Value;
end;

procedure TSearch.SetGridColumns(aGrid: TDBGrid);
var
  aIndex: Integer;
begin
  for aIndex := 0 to Pred(FListFieldsName.Count) do
  begin
    aGrid.Columns.Add;
    aGrid.Columns.Items[aIndex].FieldName        := Copy(FListFieldsName[aIndex], Pos('.', FListFieldsName[aIndex])+1, FListFieldsName[aIndex].Length).Trim;
    aGrid.Columns.Items[aIndex].Title.Caption    := FListFieldsDescription[aIndex].Trim;
    aGrid.Columns.Items[aIndex].Title.Font.Style := [TFontStyle.fsBold];
  end;
end;

procedure TSearch.SetJoinFieldsCondition1(const Value: String);
begin
  FJoinFieldsCondition1 := Value;
end;

procedure TSearch.SetJoinFieldsCondition2(const Value: String);
begin
  FJoinFieldsCondition2 := Value;
end;

procedure TSearch.SetJoinFieldsCondition3(const Value: String);
begin
  FJoinFieldsCondition3 := Value;
end;

procedure TSearch.SetJoinTable1(const Value: String);
begin
  FJoinTable1 := Value;
end;

procedure TSearch.SetJoinTable2(const Value: String);
begin
  FJoinTable2 := Value;
end;

procedure TSearch.SetJoinTable3(const Value: String);
begin
  FJoinTable3 := Value;
end;

procedure TSearch.SetMasterTable(const Value: String);
begin
  FMasterTable := Value;
end;

procedure TSearch.SetSearchTitle(const Value: String);
begin
  FSearchTitle := Value;
end;

procedure TSearch.SetValueReturn(const Value: String);
begin
  FValueReturn := Value;
end;

procedure TSearch.SetValuesSearch(const Value: String);
begin
  FValuesSearch := Value;
end;

procedure TSearch.StartSearch;
begin
  if String(Self.FMasterTable).IsEmpty then
    raise Exception.Create('Informe o nome da tabela master [principal] da consulta.');

  if String(Self.FFieldsNameList).IsEmpty then
    raise Exception.Create('Informe o nome dos campos [FIELDS] a serem retornados para pesquisa.');

  Self.FListFieldsName.StrictDelimiter        := True;
  Self.FListFieldsName.Delimiter              := ',';
  Self.FListFieldsName.DelimitedText          := Self.FFieldsNameList;

  if String(Self.FFieldsDescriptionList).IsEmpty then
    raise Exception.Create('Informe o titulo das colunas a serem exibidas no grid.');

  Self.FListFieldsDescription.StrictDelimiter := True;
  Self.FListFieldsDescription.Delimiter       := ',';
  Self.FListFieldsDescription.DelimitedText   := Self.FFieldsDescriptionList;

  if Self.FListFieldsName.Count <> Self.FListFieldsDescription.Count then
    raise Exception.Create('A quantidade de campos e suas descrições devem ser iguais.');

  Self.DrawFormSearch(FFormSearsh);
end;

end.
