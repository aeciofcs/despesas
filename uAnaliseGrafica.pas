unit uAnaliseGrafica;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, bsSkinCtrls, TeEngine, Series, ExtCtrls, TeeProcs, Chart,
  bsSkinBoxCtrls, StdCtrls, DB, ADODB, Mask, cxPropertiesStore;

type
  TfrmAnaliseGrafica = class(TForm)
    ChtGraficos: TChart;
    bsSkinStatusBar1: TbsSkinStatusBar;
    Series1: TBarSeries;
    QryGrafico: TADOQuery;
    qryCadContas: TADOQuery;
    Series2: TPieSeries;
    bsSkinExPanel1: TbsSkinExPanel;
    BtnAtualizar: TbsSkinSpeedButton;
    bsSkinGroupBox1: TbsSkinGroupBox;
    lbl01: TbsSkinStdLabel;
    cmbano: TbsSkinComboBox;
    cmbMes: TbsSkinComboBox;
    dtpData_Ini: TbsSkinDateEdit;
    lblTurma: TbsSkinStdLabel;
    dtpData_Fim: TbsSkinDateEdit;
    cmbPeriodo: TbsSkinComboBox;
    bsSkinGroupBox2: TbsSkinGroupBox;
    CmbTipoGrafico: TbsSkinComboBox;
    Series3: TFastLineSeries;
    cxPropertiesStore1: TcxPropertiesStore;
    CmbTipoAnalise: TbsSkinComboBox;
    lblParcelas: TbsSkinStdLabel;
    edtQtdeParcelas: TbsSkinSpinEdit;
    chkMostraLegenda: TbsSkinCheckRadioBox;
    procedure BtnAtualizarClick(Sender: TObject);
    procedure cmbPeriodoChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CmbTipoAnaliseChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAnaliseGrafica: TfrmAnaliseGrafica;

implementation

uses uprincipal,ufuncoes;

{$R *.dfm}

procedure TfrmAnaliseGrafica.cmbPeriodoChange(Sender: TObject);
begin
   ListaPeriodo( TbsSkinDateEdit(dtpData_Ini), TbsSkinDateEdit( dtpData_Fim), cmbPeriodo.ItemIndex, Now );
end;

procedure TfrmAnaliseGrafica.CmbTipoAnaliseChange(Sender: TObject);
begin

   case CmbTipoAnalise.itemindex of
      0 :
      Begin
         lblParcelas.Visible     := False;
         edtQtdeParcelas.visible := False;
      End;
      1,2 :
      Begin
        lblParcelas.Visible     := True;
        edtQtdeParcelas.visible := True;
        edtQtdeParcelas.value   := 10;
      End;
   end;
end;

procedure TfrmAnaliseGrafica.FormClose(Sender: TObject;
  var Action: TCloseAction);
 var  lstmArquivo: TFileStream;
begin
   lstmArquivo := TFileStream.Create(gspath + '\Config\Config_' + TForm(Sender).name, fmCreate);
   cxPropertiesStore1.StorageStream := lstmArquivo;
   cxPropertiesStore1.StoreTo;
   lstmArquivo.Free;
end;

procedure TfrmAnaliseGrafica.FormShow(Sender: TObject);
var lstmArquivo: TFileStream;
begin
   cmbPeriodoChange(cmbPeriodo);
   if FileExists(gspath + '\Config\Config_' + TForm(Sender).name) then
   Begin
      lstmArquivo := TFileStream.Create(gspath + '\Config\Config_' + TForm(Sender).name, fmOpenRead);
      cxPropertiesStore1.StorageStream := lstmArquivo;
      cxPropertiesStore1.RestoreFrom;
      lstmArquivo.Free;
   End;
   cmbPeriodoChange(cmbPeriodo);
   CmbTipoAnaliseChange(CmbTipoAnalise);
end;

procedure TfrmAnaliseGrafica.BtnAtualizarClick(Sender: TObject);
var lsGrupo  : String;
    lsConta  : String;
    lrTotal  : Double;
    liconta  : Integer;
    lsFiltro  : String;
    lsGroupBy : String;
    lsSelect  : String;
begin
   // top das despesas em grafico
   lsGroupBy := 'Group by month(Data_Lancamento)';
   lsSelect  := 'month(Data_Lancamento) as mes, Sum(Desp.Valor) as Total ';
   ChtGraficos.Legend.Visible := chkMostraLegenda.Checked;

   case CmbTipoAnalise.ItemIndex of
      0 : lsFiltro := '';
      1 : lsFiltro := 'nr_Parcela<=:parNR_Parcela AND nr_Parcela<>:parNR_Parcela2 AND ';
      2 :
      Begin
         lsFiltro := 'nr_Parcela<=:parNR_Parcela AND nr_Parcela<>:parNR_Parcela2 AND Month(Data_Lancamento)=:parData_Lancamento AND ';
         lsGroupBy := '';
         lsSelect  := 'Historico as mes, Desp.Valor as Total ';
      End;
   end;

   qryGrafico.close;
   qryGrafico.SQL.Text :='Select '+lsSelect+' '+
                           'from T_Despesas Desp, T_ContaCorrente Tpg '+
                           'where Desp.Data_Lancamento>=:pardata_Ini and '+
                           '      Desp.Data_Lancamento<=:ParData_Fim and  '+lsFiltro+' '+
                           '      Desp.D_C=:parD_C and Tpg.codigo=Desp.Cod_ContaCorrente and  tpg.exiberesumo=:parExiberesumo '+
                           ' '+lsGroupBy+' '+
                           'Order by month(Data_Lancamento) ';
   qryGrafico.Parameters.ParamValues['parData_Ini']          := Strtodate(dtpData_Ini.text);
   qryGrafico.Parameters.ParamValues['parData_Fim']          := Strtodate(dtpData_Fim.text);
   qryGrafico.Parameters.ParamValues['parD_C']               := 'D';
   qryGrafico.Parameters.ParamValues['parExiberesumo']       := 'S';

   case CmbTipoAnalise.ItemIndex of
      1 :
      Begin
         qryGrafico.Parameters.ParamValues['parNR_Parcela']   := edtQtdeParcelas.value;
         qryGrafico.Parameters.ParamValues['parNR_Parcela2']  := 1;
      End;
      2 :
      Begin
         qryGrafico.Parameters.ParamValues['parData_Lancamento'] := intToStr(cmbMes.ItemIndex);
         qryGrafico.Parameters.ParamValues['parNR_Parcela']      := edtQtdeParcelas.value;
         qryGrafico.Parameters.ParamValues['parNR_Parcela2']     := 1;
      End;
   end;

   qryGrafico.Open;
   for liConta := 1 To ChtGraficos.SeriesCount do
   begin
      ChtGraficos.Series[Liconta-1].Active := False;
      ChtGraficos.Series[Liconta-1].Clear;
   end;
   ChtGraficos.Series[CmbTipoGrafico.ItemIndex].Active :=True;
   lrtotal := 0;
   While not qryGrafico.Eof do
   Begin
      lrtotal   := lrtotal + qryGrafico.FieldByname('Total').AsFloat;
      ChtGraficos.Series[CmbTipoGrafico.ItemIndex].AddY(qryGrafico.FieldByname('Total').AsFloat,  ( qryGrafico.FieldByname( 'Mes' ).AsString ));
      qryGrafico.Next;
   End;
   ChtGraficos.Title.Text[0] := ' Analise de Estatisco'+FormatFloat('0.00',lrTotal );
end;

end.
