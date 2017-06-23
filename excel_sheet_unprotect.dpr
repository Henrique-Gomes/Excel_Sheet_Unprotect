program excel_sheet_unprotect;

{$APPTYPE CONSOLE}

uses
  SysUtils, ComObj, ActiveX, Windows, Dialogs,

  Messages, Variants, Classes, Graphics, Controls, Forms,
 StdCtrls;

function unprotectSheet(P : Variant): boolean;
    var i, j, k, l, m, i1, i2, i3, i4, i5, i6, n : integer;
    begin
      result := true;
      For i := 65 To 66 do
        For j := 65 To 66 do
          For k := 65 To 66 do
            For l := 65 To 66 do
              For m := 65 To 66 do
                For i1 := 65 To 66 do
                  For i2 := 65 To 66 do
                    For i3 := 65 To 66 do
                      For i4 := 65 To 66 do
                        For i5 := 65 To 66 do
                          For i6 := 65 To 66 do
                            For n := 32 To 126 do
                            begin
                              if P.ProtectContents then
                                try
                                  P.Unprotect(Chr(i) + Chr(j) + Chr(k) + Chr(l) + Chr(m) + Chr (i1) + Chr(i2) + Chr(i3) + Chr(i4) + Chr(i5) + Chr(i6) + Chr(n));
                                  writeln(#13#10'Here is your password:');
                                  writeln(Chr(i) + Chr(j) + Chr(k) + Chr(l) + Chr(m) + Chr (i1) + Chr(i2) + Chr(i3) + Chr(i4) + Chr(i5) + Chr(i6) + Chr(n));
                                except
                                end
                              else
                                exit;
                            end;
      writeln(#13#10'Operation terminated unsuccessfully');
      result := false;
  end;

var
  VExcel, VWB, VWS, vip: OleVariant;
  dialog : TOpenDialog;
  s: string;

begin
  dialog := TOpenDialog.Create(nil);

  writeln('Enter Excel file');
  if not dialog.Execute then
    exit;

  CoInitialize(nil);

  try
    VWB := CreateOleObject('Excel.Application');
  except
    writeln(#13#10'Cannot initiate Excel');
    Exit;
  end;

  VWB.DisplayAlerts := False;
  VWB.Visible := False;
  VWB.WorkBooks.open(dialog.filename);
  VExcel := VWB.Application;
  VWS := VWB.Workbooks[1].WorkSheets[1];

  writeln(#13#10'Working...');
  if VWS.ProtectContents then
    unprotectSheet(VWS)
  else writeln(#13#10'Sheet is not locked.');

  // Uncomment one of the lines bellow if you want to save the unprotected file
  //VWB.Workbooks[1].Save;
  //VWB.Workbooks[1].SaveAs (dialog.filename+' new', $00000027 {xlExcel7 = $00000027; on Excel.pas Unit}, vip, vip, false, false, $00000001 {xlNoChange = $00000001; on Excel_TLB.pas Unit}, false, false, vip, vip, vip, 0);

  // Closing Excel...
  VWS := Unassigned;  
  VWB := Unassigned;
  VExcel.Workbooks.Close;
  VExcel.Quit;
  VExcel := Unassigned;

  writeln(#13#10'Press Enter to Exit');
  s:=#0;
  while(s=#0) do
    readln(s);
end.
