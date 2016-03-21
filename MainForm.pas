unit MainForm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TfrmMain = class(TForm)
    edNumbertoGenerate: TEdit;
    btnGenerate: TButton;
    MemoAddresses: TMemo;
    procedure btnGenerateClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    function GetBaseAddress( lBaseSet : Integer = 0) : String;
    function CheckValid(Address: Integer) : Boolean;
    function GenerateExtraFromBase(BaseAddress, DoNumber : Integer): String;
  public
    { Public declarations }
    Increment : Integer;
  end;

var
  frmMain: TfrmMain;
Const Min_Address = 16777216;
Const MAX_Address = 2147418112;
Const c64K_Bytes = 65536       ;
Const DefaultIncrement = c64K_Bytes * 4;
implementation

{$R *.DFM}

function TfrmMain.CheckValid(Address: Integer): Boolean;
begin
    Result := false;
    If Address >= Min_Address Then
        Result := True

end;

function TfrmMain.GenerateExtraFromBase(BaseAddress,
  DoNumber: Integer): String;

var
    i : integer;
begin
   For i := 1 To DoNumber do
   begin
     //'Make the increment large enough to accommodate most dlls.
     BaseAddress := BaseAddress + (Increment);
     If BaseAddress < MAX_Address Then
         Result := Result + #13 + #10 + format('%.8x',[BaseAddress])
       //GenerateFromBase = GenerateFromBase & vbCrLf & "&H" & Hex$(lBase)
     Else
       Exit;
   end;
end;

Function TFrmMain.GetBaseAddress( lBaseSet : Integer = 0) : String;
var tempbase : Integer;
begin
    Randomize;
//    RandSeed := lbaseset;
    tempbase := 0;

  // 'Make sure the random number is above the minimum 16MB
   While Not CheckValid(tempbase) do
   begin
     tempbase := Random(MAX_Address );
   end;

  // 'Make sure the number is a multiple of 64K
   While tempbase Mod c64K_Bytes <> 0 do
   begin
     tempbase := tempbase - 1  ;
   end;

   //'Add &H to begining for copying and pasting into VB
   Result := Format('%.8x',[tempbase]) + GenerateExtraFromBase(tempbase,lbaseset);  //'$' + Hex$(tempbase) & GenerateFromBase(tempbase, lBaseSet);

End;
procedure TfrmMain.btnGenerateClick(Sender: TObject);
begin
    MemoAddresses.Lines.add( GetBaseAddress(strtoint(edNumbertoGenerate.text)));
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
    Increment := DefaultIncrement ;
end;

end.
 